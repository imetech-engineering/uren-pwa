/**
 * De uren-PWA is verhuisd naar https://apps.imetech.nl/uren/ (grote omschakeling, november 2026).
 * Deze service worker bewaart niets meer: hij gooit de oude cache weg, meldt zichzelf af en laadt open vensters
 * opnieuw, zodat iedereen de doorverwijzing in index.html ziet in plaats van de oude app uit de cache.
 */
self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", (event) => {
  event.waitUntil((async () => {
    for (const naam of await caches.keys()) await caches.delete(naam);
    await self.registration.unregister();
    for (const venster of await self.clients.matchAll({ type: "window" })) venster.navigate(venster.url);
  })());
});
