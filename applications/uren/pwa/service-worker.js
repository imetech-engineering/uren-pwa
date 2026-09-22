const CACHE = "imtech-uren-pwa-v55";
const ASSETS = [
  "./",
  "./index.html",
  "./js/melding.js",
  "./js/imetech-apps.js",
  "./js/terug.js",
  "./css/app.css",
  "./manifest.webmanifest",
  "./config.js",
  "./branding/logo-zwart.png",
  "./branding/logo-wit.png",
  "./icons/icon-192.png",
  "./icons/icon-512.png",
  "./js/auth.js",
  "./js/graph.js",
  "./js/uren_excel.js",
  "./js/uren_estimates.js",
  "./js/uren_graph_excel.js",
  "./js/uren_graph_estimates.js",
  "./js/uren_analyse.js",
  "./js/uren_inzichten.js",
  "./js/uren_invoer.js",
  "./js/combobox.js",
  "./js/data_cache.js",
  "./js/offline_queue.js",
  "./js/install.js",
  "./js/app.js",
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    // cache: "reload" = langs de HTTP-cache van de browser: anders kan een net gepubliceerde versie
    // de oude index.html (GitHub Pages cachet 10 min) in de nieuwe cache zetten.
    caches.open(CACHE).then((cache) => cache.addAll(ASSETS.map((u) => new Request(u, { cache: "reload" })))).then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches
      .keys()
      .then((keys) =>
        Promise.all(
          keys
            .filter((k) => k !== CACHE && k !== CDN_CACHE)
            .map((k) => caches.delete(k))
        )
      )
      .then(() => self.clients.claim())
  );
});


// Cache voor bibliotheken van de CDN (MSAL, Chart.js). Die staan op een vaste
// versie in de URL, dus cache-first is veilig: een nieuwe versie is een andere
// URL. Deze cache overleeft een versiebump van de app.
const CDN_CACHE = "imtech-uren-cdn-v1";
const CDN_HOSTS = ["cdn.jsdelivr.net", "cdnjs.cloudflare.com"];

/**
 * Same-origin: direct uit de cache (app start meteen) en op de achtergrond
 * bijwerken, zodat de volgende start de nieuwe versie heeft.
 */
async function uitCacheEnBijwerken(request) {
  const cache = await caches.open(CACHE);
  const cached = await cache.match(request, { ignoreSearch: true });
  const netwerk = fetch(request)
    .then((res) => {
      if (res && res.ok) cache.put(request, res.clone());
      return res;
    })
    .catch(() => null);
  if (cached) return cached;
  const res = await netwerk;
  return res || new Response("Offline en niet in de cache", { status: 503 });
}

/** CDN-bibliotheek: uit de cache, anders ophalen en bewaren. */
async function cdnUitCache(request) {
  const cache = await caches.open(CDN_CACHE);
  const cached = await cache.match(request);
  if (cached) return cached;
  try {
    // Met cors krijgen we een bruikbare (niet-ondoorzichtige) response terug.
    const res = await fetch(new Request(request.url, { mode: "cors", credentials: "omit" }));
    if (res && res.ok) cache.put(request, res.clone());
    return res;
  } catch (_) {
    return fetch(request);
  }
}

// === Wachtrij wegwerken zodra er weer verbinding is (Background Sync) ===
// De service worker heeft zelf geen inlog-token, dus hij port de app: staat die
// ergens open, dan werkt die de wachtrij af. Is er niets open, dan een seintje
// (alleen als je meldingen al hebt toegestaan).
self.addEventListener("sync", (event) => {
  if (event.tag !== "imtech-queue-sync") return;
  event.waitUntil(
    (async () => {
      const clients = await self.clients.matchAll({ type: "window", includeUncontrolled: true });
      if (clients.length) {
        for (const c of clients) c.postMessage({ type: "flush-queue" });
        return;
      }
      if (self.registration.showNotification && Notification?.permission === "granted") {
        await self.registration.showNotification("Wijzigingen klaarzetten", {
          body: "Er staan wijzigingen in de wachtrij. Open de app om ze op te slaan.",
          icon: "./icons/icon-192.png",
          tag: "imtech-queue-sync",
        });
      }
    })()
  );
});

self.addEventListener("notificationclick", (event) => {
  event.notification.close();
  event.waitUntil(self.clients.openWindow("./"));
});

self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") return;
  const url = new URL(event.request.url);
  if (url.origin === self.location.origin) {
    event.respondWith(uitCacheEnBijwerken(event.request));
    return;
  }
  if (CDN_HOSTS.includes(url.hostname)) {
    event.respondWith(cdnUitCache(event.request));
  }
});
