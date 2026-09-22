/**
 * PWA shell: sync state, routing, bottom nav, OneDrive read/write with etag.
 */
(function () {
  const state = {
    tab: "invoer",
    tt: null,
    wbso: null,
    entries: [],
    intel: null,
    etag: null,
    meta: null,
    loading: false,
    quietRefresh: false,
    syncStatus: "Nog niet geladen",
    editRow: null,
    selectedHistoryRow: null,
    weekTarget: 0,
    lastOg: "",
    lastProj: "",
    lastLoc: "",
    analyseFilters: {
      periodMode: "custom_month", // volgt de maandbalkjes in de periodekaart
      keyword: "",
      selectedOgs: [],
      selectedProjs: [],
      selectedTarieven: [],
      tariefNonZero: false,
      groupMode: "none",
      customYear: new Date().getFullYear(),
      customMonth: new Date().getMonth() + 1,
      customWeekYear: new Date().getFullYear(),
      customWeek: 1,
    },
    inzichten: { rangVeld: "opdrachtgever" },
    chartFilters: {
      year: new Date().getFullYear(),
      chartMode: "week_year",
      cumulativeEuro: false,
    },
    chartInstance: null,
    darkMode: false,
    estimates: [],
    estimateFilters: { statuses: [], search: "" },
    estimateEditRow: null,
  };

  const $ = (sel) => document.querySelector(sel);

  function setStatus(msg, isError) {
    state.syncStatus = msg;
    const el = $("#sync-status");
    if (el) {
      el.textContent = msg;
      el.classList.toggle("error", !!isError);
    }
  }

  function haptic(pattern) {
    if (navigator.vibrate) {
      try {
        navigator.vibrate(pattern);
      } catch (_) {}
    }
  }

  function showToast(msg, isError, onClick = null) {
    const t = $("#toast");
    if (!t) return;
    t.textContent = msg;
    t.classList.toggle("hidden", false);
    t.classList.toggle("error", !!isError);
    t.classList.toggle("toast-klikbaar", !!onClick);
    t.onclick = onClick
      ? () => {
          t.classList.add("hidden");
          t.onclick = null;
          onClick();
        }
      : null;
    haptic(isError ? [40, 60, 40] : 30);
    clearTimeout(showToast._tid);
    showToast._tid = setTimeout(() => t.classList.add("hidden"), onClick ? 10000 : 5000);
  }

  function applyDarkMode(on) {
    state.darkMode = !!on;
    document.documentElement.dataset.theme = on ? "dark" : "";
    const meta = $("#meta-theme-color");
    if (meta) meta.content = on ? "#121210" : "#2563EB";
    const logo = $("#header-logo");
    if (logo) logo.src = on ? "branding/logo-wit.png" : "branding/logo-zwart.png";
    const cb = $("#toggle-dark-mode");
    if (cb) cb.checked = on;
    try {
      localStorage.setItem("imtech-uren-dark", on ? "1" : "0");
    } catch (_) {}
    if (state.chartInstance) renderGrafiekenChart(state.entries);
  }

  function loadDarkPreference() {
    try {
      applyDarkMode(localStorage.getItem("imtech-uren-dark") === "1");
    } catch (_) {}
  }

  function loadWeekTarget() {
    try {
      const v = Number(localStorage.getItem("imtech-uren-week-target"));
      state.weekTarget = Number.isFinite(v) && v >= 0 ? v : 0;
    } catch (_) {
      state.weekTarget = 0;
    }
    const inp = $("#week-target-input");
    if (inp) inp.value = state.weekTarget ? String(state.weekTarget) : "";
  }

  function saveWeekTarget(val) {
    const n = Math.max(0, Number(val) || 0);
    state.weekTarget = n;
    try {
      localStorage.setItem("imtech-uren-week-target", String(n));
    } catch (_) {}
    updateInvoerStats();
  }

  function fieldsToEntry(fields, rowIndex) {
    const d = new Date(fields.datumStr + "T12:00:00");
    return {
      datum: d,
      datumStr: fields.datumStr,
      opdrachtgever: (fields.opdrachtgever || "").trim(),
      project: (fields.project || "").trim(),
      werkzaamheden: (fields.werkzaamheden || "").trim(),
      locatie: (fields.locatie || "").trim(),
      uren: Number(fields.uren) || 0,
      tarief: Number(fields.tarief) || 0,
      row_index: rowIndex,
    };
  }

  // Alleen de zichtbare tab tekenen; de rest krijgt een vlaggetje en wordt
  // getekend zodra je hem opent. Scheelt werk bij elke wijziging.
  const vuil = { invoer: true, projecten: true, analyse: true, grafieken: true };
  let invoerNogVullen = false; // eerste keer laden: formulier voorvullen

  function tekenTab(tab, initialForm = false) {
    if (tab === "invoer") {
      renderInvoer(initialForm || invoerNogVullen);
      invoerNogVullen = false;
    } else if (tab === "projecten") renderProjecten();
    else if (tab === "analyse") renderAnalyse();
    else if (tab === "grafieken") renderGrafieken();
    vuil[tab] = false;
  }

  function renderAll(initialForm = false) {
    for (const k of Object.keys(vuil)) vuil[k] = true;
    // Sta je niet op invoer, dan wordt het formulier gevuld zodra je er komt.
    if (initialForm && state.tab !== "invoer") invoerNogVullen = true;
    tekenTab(state.tab, initialForm);
  }

  function optimisticAdd(fields, tempRow = -Date.now()) {
    const entry = fieldsToEntry(fields, tempRow);
    const snapshot = { entries: [...state.entries] };
    state.entries = [...state.entries, entry];
    state.intel = UrenExcel.buildIntel(state.entries);
    return () => {
      state.entries = snapshot.entries;
      state.intel = UrenExcel.buildIntel(state.entries);
    };
  }

  function optimisticUpdate(rowIndex, fields) {
    const idx = state.entries.findIndex((e) => e.row_index === rowIndex);
    if (idx < 0) return null;
    const snapshot = { entries: [...state.entries], idx, prev: { ...state.entries[idx] } };
    const next = [...state.entries];
    next[idx] = fieldsToEntry(fields, rowIndex);
    state.entries = next;
    state.intel = UrenExcel.buildIntel(state.entries);
    return () => {
      const rollback = [...state.entries];
      rollback[snapshot.idx] = snapshot.prev;
      state.entries = rollback;
      state.intel = UrenExcel.buildIntel(state.entries);
    };
  }

  function optimisticDelete(rowIndex) {
    const idx = state.entries.findIndex((e) => e.row_index === rowIndex);
    if (idx < 0) return null;
    const snapshot = { entries: [...state.entries], removed: state.entries[idx] };
    state.entries = state.entries.filter((e) => e.row_index !== rowIndex);
    state.intel = UrenExcel.buildIntel(state.entries);
    return () => {
      state.entries = snapshot.entries;
      state.intel = UrenExcel.buildIntel(state.entries);
    };
  }

  async function updateQueueBadge() {
    const el = $("#queue-status");
    if (!el) return;
    try {
      const n = await UrenOfflineQueue.count();
      if (n > 0) {
        el.textContent = `${n} wijziging${n > 1 ? "en" : ""} wachten op sync`;
        el.classList.remove("hidden");
      } else {
        el.classList.add("hidden");
      }
    } catch (_) {
      el.classList.add("hidden");
    }
  }

  function showConflictModal(message) {
    const modal = $("#conflict-modal");
    const msg = $("#conflict-message");
    if (!modal) return;
    if (msg) {
      msg.textContent =
        (message || "") +
        " Het bestand is op een ander apparaat gewijzigd. Ververs om de nieuwste versie te laden en probeer opnieuw.";
    }
    modal.classList.remove("hidden");
  }

  function closeConflictModal() {
    $("#conflict-modal")?.classList.add("hidden");
  }

  function isNetworkError(e) {
    if (!UrenOfflineQueue.isOnline()) return true;
    const m = (e?.message || "").toLowerCase();
    return (
      e?.name === "TypeError" ||
      m.includes("failed to fetch") ||
      m.includes("network") ||
      m.includes("load failed")
    );
  }

  async function executeMutation(descriptor) {
    const token = await ensureLoggedIn();
    const path = drivePath();
    const { kind, fields, rowIndex } = descriptor;
    if (kind === "hours_add") {
      await UrenGraphExcel.withSession(path, token, (sid) =>
        UrenGraphExcel.addEntry(path, token, sid, fields)
      );
    } else if (kind === "hours_update") {
      await UrenGraphExcel.withSession(path, token, (sid) =>
        UrenGraphExcel.updateEntry(path, token, sid, rowIndex, fields)
      );
    } else if (kind === "hours_delete") {
      await UrenGraphExcel.withSession(path, token, (sid) =>
        UrenGraphExcel.deleteEntry(path, token, sid, rowIndex)
      );
    } else if (kind === "estimate_add") {
      await UrenGraphEstimates.withSession(path, token, (sid) =>
        UrenGraphEstimates.addEstimate(path, token, sid, fields)
      );
    } else if (kind === "estimate_update") {
      await UrenGraphEstimates.withSession(path, token, (sid) =>
        UrenGraphEstimates.updateEstimate(path, token, sid, rowIndex, fields)
      );
    } else if (kind === "estimate_delete") {
      await UrenGraphEstimates.withSession(path, token, (sid) =>
        UrenGraphEstimates.deleteEstimate(path, token, sid, rowIndex)
      );
    }
  }

  async function flushOfflineQueue() {
    if (!UrenOfflineQueue.isOnline() || !UrenAuth.isLoggedIn()) return;
    const items = await UrenOfflineQueue.getAll();
    if (!items.length) return;
    setStatus(`Sync ${items.length} wachtende wijziging(en)…`);
    for (const item of items) {
      try {
        await executeMutation(item);
        await UrenOfflineQueue.remove(item.id);
      } catch (e) {
        if (e.name === "GraphConflictError") {
          showConflictModal(e.message);
          break;
        }
        if (isNetworkError(e)) break;
        await UrenOfflineQueue.remove(item.id);
        showToast(e.message || String(e), true);
      }
    }
    await updateQueueBadge();
    await refreshFromCloudQuiet();
  }

  async function reloadFromCloud() {
    const token = await ensureLoggedIn();
    const path = drivePath();
    const meta = await UrenGraph.getDriveItemMeta(path, token);
    const entries = await UrenGraphExcel.readAllEntries(path, token);
    const estimates = await UrenGraphEstimates.readAllEstimates(path, token);
    state.entries = entries;
    state.estimates = estimates;
    state.intel = UrenExcel.buildIntel(entries);
    state.etag = meta.etag;
    state.meta = meta;
    return meta;
  }

  async function refreshFromCloudQuiet() {
    if (state.loading) return;
    while (state.quietRefresh) {
      await new Promise((r) => setTimeout(r, 40));
    }
    state.quietRefresh = true;
    try {
      const meta = await reloadFromCloud();
      setStatus(`Bijgewerkt ${new Date(meta.lastModified).toLocaleString("nl-NL")}`);
      renderAll(false);
      renderAccount();
    } catch (e) {
      showToast(e.message || String(e), true);
      throw e;
    } finally {
      state.quietRefresh = false;
    }
  }

  function applyOptimistic(optimisticRollback) {
    if (!optimisticRollback) return null;
    const rollback = optimisticRollback();
    renderAll();
    return rollback;
  }

  async function queueOfflineMutation(descriptor) {
    await UrenOfflineQueue.add(descriptor);
    meldSyncAan();
    await updateQueueBadge();
    showToast("Offline — wijziging in wachtrij");
  }

  /**
   * Opslaan zonder dat het scherm onder je handen verandert: de wijziging staat
   * meteen in de lijst, daarna gaat hij naar OneDrive. Mislukt dat, dan draaien
   * we hem terug. De verversing achteraf raakt het invulformulier niet aan.
   */
  async function persistMutation(descriptor, optimisticRollback) {
    const rollback = applyOptimistic(optimisticRollback);
    if (!UrenOfflineQueue.isOnline()) {
      await queueOfflineMutation(descriptor);
      return;
    }

    setStatus("Opslaan in OneDrive…");
    try {
      await executeMutation(descriptor);
    } catch (e) {
      if (e.name === "GraphConflictError") {
        showConflictModal(e.message);
        throw e;
      }
      if (isNetworkError(e)) {
        await queueOfflineMutation(descriptor);
        return;
      }
      // Niet opgeslagen → de regel weer uit het scherm halen.
      if (rollback) {
        rollback();
        renderAll();
      }
      if (e.name === "GraphLockError") {
        showToast(e.message, true);
      } else {
        showToast(e.message || String(e), true);
      }
      throw e;
    }
    try {
      await refreshFromCloudQuiet();
      showToast("Opgeslagen in OneDrive");
      return true;
    } catch (e) {
      showToast("Opgeslagen in OneDrive — ververs handmatig als de lijst niet klopt.", true);
      throw e;
    }
  }

  function updateGrafiekControlsVisibility() {
    const mode = state.chartFilters.chartMode;
    $("#grafiek-euro-row")?.classList.toggle("hidden", mode !== "cumulative");
  }

  const MONTH_LABELS = ["Jan", "Feb", "Mrt", "Apr", "Mei", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dec"];
  const PIE_COLORS = ["#2563eb", "#16a34a", "#dc2626", "#ca8a04", "#9333ea", "#0891b2", "#ea580c", "#64748b"];

  function renderGrafiekenChart(allEntries) {
    const cf = state.chartFilters;
    const canvas = $("#grafiek-chart");
    if (!canvas || typeof Chart === "undefined") return;
    const base = UrenAnalyse.filterRowsForCharts(allEntries, state.analyseFilters);
    const year = cf.year || UrenAnalyse.chartYearFromRows(base);
    const yearRows = UrenAnalyse.rowsForChartYear(base, year);
    let labels = [];
    let data = [];
    let data2 = [];
    let chartType = "bar";
    let title = "";
    const mode = cf.chartMode;

    if (mode === "week_year") {
      const weeks = UrenAnalyse.aggregateHoursPerIsoWeek(yearRows, year);
      labels = weeks.map((w) => `W${w.week}`);
      data = weeks.map((w) => w.uren);
      title = `Uren per ISO-week ${year}`;
    } else if (mode === "month_year") {
      const months = UrenAnalyse.aggregateHoursPerMonth(yearRows, year);
      labels = months.map((m) => MONTH_LABELS[m.month - 1]);
      data = months.map((m) => m.uren);
      title = `Uren per maand ${year}`;
    } else if (mode === "og_bar" || mode === "og_pie") {
      const ogs = UrenAnalyse.aggregateHoursPerOg(yearRows, 10);
      labels = ogs.map((o) => o.og);
      data = ogs.map((o) => o.uren);
      title = `Uren per opdrachtgever ${year}`;
      chartType = mode === "og_pie" ? "pie" : "bar";
    } else if (mode === "revenue_og_bar" || mode === "revenue_og_pie") {
      const ogs = UrenAnalyse.aggregateRevenuePerOg(yearRows, 10);
      labels = ogs.map((o) => o.og);
      data = ogs.map((o) => o.bedrag);
      title = `Omzet (€) per opdrachtgever ${year}`;
      chartType = mode === "revenue_og_pie" ? "pie" : "bar";
    } else if (mode === "revenue_month") {
      const months = UrenAnalyse.aggregateRevenuePerMonth(yearRows, year);
      labels = months.map((m) => MONTH_LABELS[m.month - 1]);
      data = months.map((m) => m.bedrag);
      title = `Omzet (€) per maand ${year}`;
    } else if (mode === "locatie") {
      const locs = UrenAnalyse.aggregateHoursPerLocatie(yearRows, 12);
      labels = locs.map((l) => l.loc);
      data = locs.map((l) => l.uren);
      title = `Uren per locatie ${year}`;
    } else if (mode === "cumulative") {
      const cum = UrenAnalyse.aggregateCumulativeForYear(yearRows, year);
      labels = cum.map((c) => c.label.slice(5));
      data = cum.map((c) => c.uren);
      if (cf.cumulativeEuro) data2 = cum.map((c) => c.bedrag);
      chartType = "line";
      title = cf.cumulativeEuro
        ? `Cumulatief uren en omzet ${year}`
        : `Cumulatief uren ${year}`;
    }

    const grid = getComputedStyle(document.documentElement).getPropertyValue("--chart-grid").trim();
    const accent = getComputedStyle(document.documentElement).getPropertyValue("--accent").trim();
    const text = getComputedStyle(document.documentElement).getPropertyValue("--text-secondary").trim();
    const datasets =
      chartType === "line" && data2.length
        ? [
            {
              label: "Uren",
              data,
              backgroundColor: "transparent",
              borderColor: accent,
              borderWidth: 2,
              fill: false,
              tension: 0.2,
              yAxisID: "y",
            },
            {
              label: "Omzet €",
              data: data2,
              backgroundColor: "transparent",
              borderColor: "#16a34a",
              borderWidth: 2,
              fill: false,
              tension: 0.2,
              yAxisID: "y1",
            },
          ]
        : [
            {
              label: title,
              data,
              backgroundColor: chartType === "pie" ? PIE_COLORS : accent,
              borderColor: chartType === "pie" ? "#ffffff" : accent,
              borderWidth: chartType === "pie" ? 1 : 0,
            },
          ];

    if (state.chartInstance) state.chartInstance.destroy();
    const options = {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: chartType === "pie" || data2.length > 0, labels: { color: text } }, title: { display: !!title, text: title, color: text } },
    };
    if (chartType !== "pie") {
      options.scales = {
        x: { ticks: { color: text, maxRotation: 45 }, grid: { color: grid } },
        y: { ticks: { color: text }, grid: { color: grid }, position: "left" },
      };
      if (data2.length) {
        options.scales.y1 = {
          ticks: { color: text },
          grid: { drawOnChartArea: false },
          position: "right",
        };
      }
    }
    state.chartInstance = new Chart(canvas, { type: chartType, data: { labels, datasets }, options });
  }

  function renderGrafieken() {
    updateGrafiekControlsVisibility();
    const summary = $("#grafiek-summary");
    if (!state.entries.length) {
      if (summary) summary.textContent = "Geen data — ververs uit OneDrive.";
      if (state.chartInstance) {
        state.chartInstance.destroy();
        state.chartInstance = null;
      }
      return;
    }
    const base = UrenAnalyse.filterRowsForCharts(state.entries, state.analyseFilters);
    const year = state.chartFilters.year || UrenAnalyse.chartYearFromRows(base);
    const yearRows = UrenAnalyse.rowsForChartYear(base, year);
    const sum = UrenAnalyse.summarize(yearRows);
    if (summary) {
      summary.textContent = `${year}: ${sum.totU.toFixed(1)} u · €${sum.totE.toFixed(2)} · ${sum.count} regels`;
    }
    renderGrafiekenChart(state.entries);
  }

  async function ensureLoggedIn() {
    if (!UrenAuth.isLoggedIn()) {
      throw new Error("Log eerst in met je Microsoft-account (Instellingen).");
    }
    return UrenAuth.acquireToken();
  }

  function drivePath() {
    return UrenAuth.getConfig().graph.drivePath;
  }

  /** Laatste stand lokaal bewaren, zodat de app volgende keer meteen gevuld is. */
  function bewaarSnapshot() {
    UrenCache?.bewaar("data", {
      entries: state.entries,
      estimates: state.estimates,
      etag: state.etag,
      lastModified: state.meta?.lastModified || null,
    });
  }

  /** Bij het starten: eerst de bewaarde stand tonen, daarna pas de cloud. */
  async function toonSnapshot() {
    const snap = await UrenCache?.lees("data");
    if (!snap?.entries?.length || state.entries.length) return false;
    state.entries = snap.entries;
    state.estimates = snap.estimates || [];
    state.intel = UrenExcel.buildIntel(state.entries);
    state.etag = snap.etag || null;
    const wanneer = snap.lastModified || snap.bewaardOp;
    setStatus(
      `Laatst bekend ${new Date(wanneer).toLocaleString("nl-NL")} — bijwerken…`
    );
    renderAll(true);
    return true;
  }

  async function refreshFromCloud() {
    laadTimetick();
    laadWbso();
    state.loading = true;
    setStatus("Laden uit OneDrive…");
    try {
      await flushOfflineQueue();
      const token = await ensureLoggedIn();
      const path = drivePath();
      // Eerst de kleine bestand-info: is het werkboek niet gewijzigd en hebben we
      // de gegevens al, dan hoeven de zware tabellen niet opnieuw opgehaald.
      const meta = await UrenGraph.getDriveItemMeta(path, token);
      if (meta.etag && meta.etag === state.etag && state.entries.length) {
        state.meta = meta;
        setStatus(`Bij (ongewijzigd) ${new Date(meta.lastModified).toLocaleString("nl-NL")}`);
        renderAll(true);
        renderAccount();
        await updateQueueBadge();
        return;
      }
      const [entries, estimates] = await Promise.all([
        UrenGraphExcel.readAllEntries(path, token),
        UrenGraphEstimates.readAllEstimates(path, token),
      ]);
      state.entries = entries;
      state.estimates = estimates;
      state.intel = UrenExcel.buildIntel(entries);
      state.etag = meta.etag;
      state.meta = meta;
      bewaarSnapshot();
      setStatus(`Bijgewerkt ${new Date(meta.lastModified).toLocaleString("nl-NL")}`);
      renderAll(true);
      renderAccount();
      await updateQueueBadge();
    } catch (e) {
      setStatus(e.message || String(e), true);
      showToast(e.message || String(e), true);
      throw e;
    } finally {
      state.loading = false;
    }
  }

  function getEstimateFormFields() {
    return {
      datumStr: $("#est-datum")?.value,
      opdrachtgever: $("#est-og")?.value,
      project: $("#est-project")?.value,
      ureninschatting: $("#est-planned")?.value,
      status: $("#est-status")?.value,
      opmerking: $("#est-opmerking")?.value,
    };
  }

  function fillEstimateStatusSelect() {
    const sel = $("#est-status");
    if (!sel) return;
    sel.innerHTML = "";
    for (const s of UrenEstimates.PROJECT_STATUSES) {
      const o = document.createElement("option");
      o.value = s;
      o.textContent = s;
      sel.appendChild(o);
    }
  }

  function openProjectModal(entry) {
    fillEstimateStatusSelect();
    state.estimateEditRow = entry?.row_index ?? null;
    $("#project-modal-title").textContent = entry ? "Project bewerken" : "Project toevoegen";
    $("#est-datum").value = entry?.datumStr || UrenExcel.formatDateIso(new Date());
    $("#est-og").value = entry?.opdrachtgever || "";
    $("#est-project").value = entry?.project || "";
    $("#est-planned").value = entry?.ureninschatting ?? "";
    $("#est-status").value = entry?.status || UrenEstimates.DEFAULT_STATUS;
    $("#est-opmerking").value = entry?.opmerking || "";
    const actual = entry?.gemaakte_uren;
    const delta = entry ? UrenEstimates.displayDelta(entry) : null;
    $("#est-actual").textContent = entry ? `${Number(actual || 0).toFixed(1)} u` : "—";
    $("#est-delta").textContent =
      delta != null ? `${Number(delta).toFixed(1)} u` : entry ? "—" : "—";
    $("#btn-est-delete")?.classList.toggle("hidden", !entry);
    const dl = $("#dl-og-est");
    if (dl && state.intel) {
      dl.innerHTML = "";
      for (const n of comboOptionsOg()) {
        const o = document.createElement("option");
        o.value = n;
        dl.appendChild(o);
      }
    }
    $("#project-modal")?.classList.remove("hidden");
  }

  function closeProjectModal() {
    state.estimateEditRow = null;
    $("#project-modal")?.classList.add("hidden");
  }

  async function onEstimateSave() {
    const fields = getEstimateFormFields();
    if (!fields.project?.trim()) {
      showToast("Project is verplicht.", true);
      return;
    }
    if (!fields.datumStr) {
      showToast("Datum is verplicht.", true);
      return;
    }
    const bewerkRij = state.estimateEditRow;
    closeProjectModal(); // meteen dicht; opslaan gebeurt daarna
    try {
      await persistMutation(
        bewerkRij
          ? { kind: "estimate_update", fields, rowIndex: bewerkRij }
          : { kind: "estimate_add", fields, rowIndex: null }
      );
    } catch (_) {}
  }

  async function onEstimateDelete() {
    if (!state.estimateEditRow) return;
    if (!confirm("Projectrij verwijderen uit Excel?")) return;
    const rij = state.estimateEditRow;
    closeProjectModal();
    try {
      await persistMutation({ kind: "estimate_delete", fields: null, rowIndex: rij });
    } catch (_) {}
  }

  function renderProjecten() {
    const summaryEl = $("#projecten-summary");
    const listEl = $("#projecten-list");
    if (!summaryEl || !listEl) return;
    if (!state.estimates?.length) {
      summaryEl.textContent = "Geen projecten — ververs uit OneDrive.";
      listEl.innerHTML = "";
      $("#projecten-status-cards").innerHTML = "";
      return;
    }
    const summary = UrenEstimates.buildStatusSummary(state.estimates);
    const activeLine =
      summary.activePlanned > 0 || summary.activeActual > 0
        ? `Actief: ${summary.activeActual} / ${summary.activePlanned} u (resterend ${summary.activeRemaining} u)`
        : "";
    const overLine = summary.overBudget.length
      ? ` | ${summary.overBudget.length} over budget`
      : "";
    summaryEl.textContent = `${state.estimates.length} projecten${activeLine ? " · " + activeLine : ""}${overLine}`;

    const cardsEl = $("#projecten-status-cards");
    if (cardsEl) {
      cardsEl.innerHTML = "";
      for (const st of UrenEstimates.PROJECT_STATUSES) {
        const count = summary.counts[st] || 0;
        if (!count) continue;
        const card = document.createElement("button");
        card.type = "button";
        card.className =
          "status-card" +
          (state.estimateFilters.statuses.includes(st) ? " active" : "");
        card.innerHTML = `<div class="sc-count">${count}</div><div class="sc-label">${st}</div>`;
        card.addEventListener("click", () => {
          const f = state.estimateFilters;
          const i = f.statuses.indexOf(st);
          if (i >= 0) f.statuses.splice(i, 1);
          else f.statuses.push(st);
          renderProjecten();
        });
        cardsEl.appendChild(card);
      }
    }

    const chipsEl = $("#chips-project-status");
    if (chipsEl) {
      chipsEl.innerHTML = "";
      const mk = (label, val) => {
        const b = document.createElement("button");
        b.type = "button";
        const active =
          val == null
            ? state.estimateFilters.statuses.length === 0
            : state.estimateFilters.statuses.includes(val);
        b.className = "chip" + (active ? " active" : "");
        b.textContent = label;
        b.addEventListener("click", () => {
          if (val == null) state.estimateFilters.statuses = [];
          else {
            const i = state.estimateFilters.statuses.indexOf(val);
            if (i >= 0) state.estimateFilters.statuses.splice(i, 1);
            else state.estimateFilters.statuses.push(val);
          }
          renderProjecten();
        });
        chipsEl.appendChild(b);
      };
      mk("Alles", null);
      for (const st of UrenEstimates.PROJECT_STATUSES) mk(st, st);
    }

    const filtered = UrenEstimates.filterEstimates(
      UrenEstimates.sortEstimates(state.estimates, state.entries),
      state.estimateFilters
    );
    listEl.innerHTML = "";
    for (const row of filtered) {
      const li = document.createElement("li");
      li.className = "project-card";
      const planned = Number(row.ureninschatting) || 0;
      const actual = Number(row.gemaakte_uren) || 0;
      const delta = UrenEstimates.displayDelta(row);
      const pct = planned > 0 ? Math.min(100, (actual / planned) * 100) : 0;
      const over = delta != null && delta < 0;
      const deltaHtml =
        delta != null
          ? `<span class="${over ? "delta-negative" : ""}">${delta > 0 ? "+" : ""}${Number(delta).toFixed(1)} u</span>`
          : "";
      li.innerHTML = `<div class="project-card-head">
          <span class="project-card-title">${row.project}</span>
          <span class="status-badge ${UrenEstimates.statusClass(row.status)}">${row.status}</span>
        </div>
        <div class="project-card-og">${row.opdrachtgever || "—"}</div>
        <div class="project-progress"><div class="project-progress-bar${over ? " over" : ""}" style="width:${pct}%"></div></div>
        <div class="project-stats">
          <span>${actual.toFixed(1)} / ${planned.toFixed(1)} u</span>
          ${deltaHtml}
        </div>`;
      li.addEventListener("click", () => openProjectModal(row));
      listEl.appendChild(li);
    }
  }

  function getFormFields() {
    return {
      datumStr: $("#field-datum")?.value,
      opdrachtgever: $("#field-og")?.value,
      project: $("#field-project")?.value,
      werkzaamheden: $("#field-werk")?.value,
      locatie: $("#field-locatie")?.value,
      uren: $("#field-uren")?.value,
      tarief: $("#field-tarief")?.value,
    };
  }

  function fillForm(entry) {
    $("#field-datum").value = entry?.datumStr || UrenExcel.formatDateIso(new Date());
    $("#field-og").value = entry?.opdrachtgever || "";
    $("#field-project").value = entry?.project || "";
    $("#field-werk").value = entry?.werkzaamheden || "";
    $("#field-locatie").value = entry?.locatie || "";
    $("#field-uren").value = entry?.uren ?? "";
    $("#field-tarief").value = entry?.tarief ?? "";
    onComboChange("og");
    updateInvoerStats();
  }

  function resetFormAfterSave() {
    const og = ($("#field-og")?.value || "").trim() || state.lastOg;
    const proj = ($("#field-project")?.value || "").trim() || state.lastProj;
    const loc = ($("#field-locatie")?.value || "").trim() || state.lastLoc;
    state.lastOg = og;
    state.lastProj = proj;
    state.lastLoc = loc;
    $("#field-datum").value = UrenExcel.formatDateIso(new Date());
    $("#field-og").value = og;
    $("#field-project").value = proj;
    $("#field-locatie").value = loc;
    $("#field-werk").value = "";
    $("#field-uren").value = "1";
    // Historie-filter leegmaken zodat de zojuist opgeslagen regel meteen zichtbaar is.
    const search = $("#history-search");
    if (search && search.value) {
      search.value = "";
      renderHistory();
    }
    onComboChange("og");
    updateInvoerStats();
    // Bewust géén focus: dit draait ná de OneDrive-sync, en dan springt het
    // toetsenbord onverwacht open terwijl je al iets anders aan het doen bent.
  }

  function invoerContext() {
    return {
      og: ($("#field-og")?.value || "").trim(),
      proj: ($("#field-project")?.value || "").trim(),
      loc: ($("#field-locatie")?.value || "").trim(),
    };
  }

  /** Bij opdrachtgever alleen sorteren; aanvullen pas bij project/locatie. */
  function onComboChange(trigger) {
    renderDatalists();
    const { og, proj, loc } = invoerContext();
    if (og && proj) {
      const t = UrenInvoer.suggestTarief(state.intel, og, proj);
      if (t !== "") $("#field-tarief").value = t;
    }
    if (trigger === "project" && og && proj) {
      const combo = state.intel?.last_combo?.[`${og}\0${proj}`];
      if (combo?.locatie && !($("#field-locatie")?.value || "").trim()) {
        $("#field-locatie").value = combo.locatie;
      }
    }
    if (trigger === "loc" && og && proj && loc) {
      const wzEl = $("#field-werk");
      if (wzEl && !(wzEl.value || "").trim()) {
        const opts = UrenInvoer.smartWerkzaamheden(state.intel, og, proj, loc);
        if (opts.length) wzEl.value = opts[0];
      }
    }
  }

  function comboOptionsOg() {
    if (!state.intel) return [];
    return UrenInvoer.sortByUsage(state.intel.og_usage, state.intel.all_opdrachtgevers);
  }

  function comboOptionsProj() {
    if (!state.intel) return [];
    return UrenInvoer.smartProjects(state.intel, $("#field-og")?.value);
  }

  function comboOptionsLoc() {
    if (!state.intel) return [];
    return UrenInvoer.smartLocaties(
      state.intel,
      $("#field-og")?.value,
      $("#field-project")?.value
    );
  }

  function comboOptionsWerk() {
    if (!state.intel) return [];
    const { og, proj, loc } = invoerContext();
    return UrenInvoer.smartWerkzaamheden(state.intel, og, proj, loc);
  }

  function renderDatalists() {
    if (!state.intel) return;
    const fill = (id, items) => {
      const dl = document.getElementById(id);
      if (!dl) return;
      dl.innerHTML = "";
      for (const n of items) {
        const o = document.createElement("option");
        o.value = n;
        dl.appendChild(o);
      }
    };
    fill("dl-og", comboOptionsOg());
    fill("dl-project", comboOptionsProj());
    fill("dl-locatie", comboOptionsLoc());
    fill("dl-werk", comboOptionsWerk());
  }

  function adjustHours(delta) {
    const el = $("#field-uren");
    if (!el) return;
    let v = parseFloat(el.value);
    if (!Number.isFinite(v)) v = 0;
    v = Math.max(0, Math.round((v + delta) * 2) / 2);
    el.value = v === 0 ? "" : String(v);
  }

  function applyHistoryToForm(entry, focusWerk = false) {
    if (!entry) return;
    state.editRow = null;
    $("#btn-save").textContent = "Opslaan";
    const today = UrenExcel.formatDateIso(new Date());
    $("#field-datum").value = today;
    $("#field-og").value = entry.opdrachtgever || "";
    $("#field-project").value = entry.project || "";
    $("#field-locatie").value = entry.locatie || "";
    $("#field-werk").value = entry.werkzaamheden || "";
    $("#field-uren").value = entry.uren ?? "";
    $("#field-tarief").value = entry.tarief ?? "";
    onComboChange();
    switchTab("invoer");
    showToast("Regel overgenomen — datum blijft vandaag");
    if (focusWerk) $("#field-werk")?.focus();
    else $("#field-uren")?.focus();
  }

  function isoWeekInfo(d) {
    const date = d instanceof Date ? d : new Date(d);
    const tmp = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    tmp.setUTCDate(tmp.getUTCDate() + 4 - (tmp.getUTCDay() || 7));
    const yearStart = new Date(Date.UTC(tmp.getUTCFullYear(), 0, 1));
    const week = Math.ceil(((tmp - yearStart) / 86400000 + 1) / 7);
    return { year: tmp.getUTCFullYear(), week };
  }

  function updateInvoerStats() {
    const dayEl = $("#invoer-day-stats");
    const weekEl = $("#invoer-week-stats");
    if (!dayEl || !weekEl) return;
    const datumStr = $("#field-datum")?.value;
    if (!datumStr || !state.entries?.length) {
      dayEl.textContent = "—";
      weekEl.textContent = "—";
      return;
    }
    const sel = new Date(datumStr + "T12:00:00");
    const { year: weekYear, week: weekNo } = isoWeekInfo(sel);
    let dayH = 0;
    let weekH = 0;
    for (const e of state.entries) {
      const ed = e.datum instanceof Date ? e.datum : new Date(e.datum);
      if (UrenExcel.formatDateIso(ed) === datumStr) dayH += e.uren;
      const iw = isoWeekInfo(ed);
      if (iw.year === weekYear && iw.week === weekNo) weekH += e.uren;
    }
    dayEl.textContent = `${dayH.toFixed(1)} u`;
    weekEl.textContent = `${weekH.toFixed(1)} u (week ${weekNo})`;

    const targetWrap = $("#week-target-wrap");
    const targetLabel = $("#week-target-label");
    const targetFill = $("#week-target-fill");
    const target = state.weekTarget;
    if (targetWrap && target > 0) {
      targetWrap.classList.remove("hidden");
      const pct = Math.min(100, (weekH / target) * 100);
      if (targetLabel) {
        targetLabel.textContent = `${weekH.toFixed(1)} / ${target} u`;
      }
      if (targetFill) {
        targetFill.style.width = `${pct}%`;
        targetFill.classList.toggle("over", weekH > target);
      }
    } else if (targetWrap) {
      targetWrap.classList.add("hidden");
    }
  }

  const ICON_APPLY =
    '<svg viewBox="0 0 24 24" width="17" height="17" fill="none" stroke="currentColor" ' +
    'stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
    '<path d="M12 3v10"/><path d="m8 11 4 4 4-4"/>' +
    '<path d="M4 17v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2"/></svg>';
  const ICON_PENCIL =
    '<svg viewBox="0 0 24 24" width="17" height="17" fill="none" stroke="currentColor" ' +
    'stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
    '<path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4Z"/></svg>';
  const ICON_TRASH =
    '<svg viewBox="0 0 24 24" width="17" height="17" fill="none" stroke="currentColor" ' +
    'stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
    '<path d="M3 6h18"/><path d="M8 6V4h8v2"/><path d="M19 6l-1 14H6L5 6"/>' +
    '<path d="M10 11v6M14 11v6"/></svg>';

  function esc(s) {
    const d = document.createElement("div");
    d.textContent = s == null ? "" : String(s);
    return d.innerHTML;
  }

  // Twee regels per item: wie/wat + uren boven, de rest eronder.
  const historyTitel = (e) =>
    [e.opdrachtgever, e.project].filter(Boolean).join(" · ") || "(geen project)";
  const historyUren = (e) =>
    `${String(e.uren).replace(".", ",")} u × €${String(e.tarief).replace(".", ",")}`;
  const historySub = (e) =>
    [e.datumStr, e.locatie, e.werkzaamheden].filter(Boolean).join(" · ");
  const kortDatum = (iso) => (/^\d{4}-\d{2}-\d{2}/.test(iso || "") ? `${iso.slice(8, 10)}-${iso.slice(5, 7)}` : iso || "");
  const ICON_TT =
    '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/></svg>';

  function ttBadge(st) {
    if (st === "in") return '<span class="tt-badge tt-in" title="Staat in Timetick">TT ✓</span>';
    if (st === "verstuurd") return '<span class="tt-badge tt-verstuurd" title="Verstuurd, wordt ingevoerd">TT …</span>';
    if (st === "tarief0") return '<span class="tt-badge tt-tarief0">TT ?</span>';
    return "";
  }

  function renderHistory() {
    const list = $("#history-list");
    if (!list || !state.intel) return;
    list.innerHTML = "";
    const q = ($("#history-search")?.value || "").toLowerCase();
    let items = state.intel.history;
    if (q) {
      items = items.filter((e) =>
        UrenInvoer.formatHistoryLine(e).toLowerCase().includes(q)
      );
    }
    const totalMatched = items.length;
    items = items.slice(0, 80);
    if (state.selectedHistoryRow != null && !items.some((e) => e.row_index === state.selectedHistoryRow)) {
      state.selectedHistoryRow = null;
    }
    renderTimetickBalk();
    for (const e of items) {
      const li = document.createElement("li");
      const tt = ttStatus(e);
      const open = e.row_index === state.selectedHistoryRow;
      const wbso = isWbso(e);
      li.className = "history-item" + (open ? " selected" : "") + (tt ? " tt-" + tt : "") + (wbso ? " wbso" : "");
      li.dataset.row = String(e.row_index);
      const kanTT = isR2R(e) && getal(e.tarief) > 0 && getal(e.uren) > 0 && !tt;
      li.innerHTML = `<span class="hi-head">
          <span class="hi-title">${esc(e.project || "(geen project)")}${e.opdrachtgever ? `<span class="hi-og"> · ${esc(e.opdrachtgever)}</span>` : ""}</span>
          <span class="hi-uren">${wbso ? '<span class="tt-badge wbso-badge" title="Telt voor de WBSO">WBSO</span>' : ""}${ttBadge(tt)}${esc(historyUren(e))}</span>
        </span>
        <span class="hi-werk"><span class="hi-datum">${esc(kortDatum(e.datumStr))}</span>${esc(e.werkzaamheden || "(geen omschrijving)")}</span>
        ${
          open
            ? `${tt === "tarief0" ? '<span class="hi-tt-uitleg">Staat in Timetick, maar hier op tarief 0</span>' : ""}
        ${e.locatie ? `<span class="hi-loc">${esc(e.locatie)}</span>` : ""}
        <span class="history-actions">
          <button type="button" class="btn-icon" data-act="apply" data-row="${e.row_index}" aria-label="Overnemen in formulier" title="Overnemen">${ICON_APPLY}</button>
          <button type="button" class="btn-icon" data-act="edit" data-row="${e.row_index}" aria-label="Bewerken" title="Bewerken">${ICON_PENCIL}</button>
          ${kanTT ? `<button type="button" class="btn-icon btn-icon-tt" data-act="tt" data-row="${e.row_index}" aria-label="Naar Timetick" title="Naar Timetick">${ICON_TT}</button>` : ""}
          <button type="button" class="btn-icon btn-icon-danger" data-act="del" data-row="${e.row_index}" aria-label="Verwijderen" title="Verwijderen">${ICON_TRASH}</button>
        </span>`
            : ""
        }`;
      bindHistorySwipe(li, e);
      li.addEventListener("dblclick", (ev) => {
        if (ev.target.closest("button")) return;
        applyHistoryToForm(e, true);
      });
      li.addEventListener("click", (ev) => {
        if (ev.target.closest("button")) return;
        const now = Date.now();
        if (
          renderHistory._lastTap?.row === e.row_index &&
          now - renderHistory._lastTap.t < 450
        ) {
          applyHistoryToForm(e, true);
          renderHistory._lastTap = null;
          return;
        }
        renderHistory._lastTap = { row: e.row_index, t: now };
        // Nog eens tikken op een open regel klapt hem weer dicht.
        state.selectedHistoryRow = open ? null : e.row_index;
        renderHistory();
      });
      list.appendChild(li);
    }
    list.querySelectorAll("button").forEach((btn) => {
      btn.addEventListener("click", async (ev) => {
        ev.stopPropagation();
        const row = Number(btn.dataset.row);
        const entry = state.entries.find((x) => x.row_index === row);
        if (btn.dataset.act === "apply" && entry) {
          applyHistoryToForm(entry, true);
        } else if (btn.dataset.act === "tt" && entry) {
          stuurNaarTimetick([entry]);
        } else if (btn.dataset.act === "edit" && entry) {
          state.editRow = row;
          state.selectedHistoryRow = row;
          fillForm(entry);
          $("#btn-save").textContent = "Bijwerken";
          switchTab("invoer");
        } else if (btn.dataset.act === "del") {
          if (!confirm("Regel verwijderen uit Excel?")) return;
          // Bewerkte je net die regel? Dan het formulier meteen vrijgeven.
          if (state.editRow === row) {
            state.editRow = null;
            $("#btn-save").textContent = "Opslaan";
            resetFormAfterSave();
          }
          await verwijderRegel(entry || { row_index: row });
        }
      });
    });
  }

  function bindHistorySwipe(li, entry) {
    let startX = 0;
    let startY = 0;
    let tracking = false;
    const threshold = 72;

    li.addEventListener(
      "touchstart",
      (ev) => {
        if (ev.target.closest("button")) return;
        const t = ev.touches[0];
        startX = t.clientX;
        startY = t.clientY;
        tracking = true;
      },
      { passive: true }
    );

    li.addEventListener(
      "touchmove",
      (ev) => {
        if (!tracking) return;
        const t = ev.touches[0];
        const dx = t.clientX - startX;
        const dy = t.clientY - startY;
        if (Math.abs(dy) > Math.abs(dx)) {
          tracking = false;
          li.style.transform = "";
          li.classList.remove("swiping");
          return;
        }
        if (Math.abs(dx) > 8) {
          li.classList.add("swiping");
          li.style.transform = `translateX(${dx}px)`;
        }
      },
      { passive: true }
    );

    li.addEventListener("touchend", async (ev) => {
      if (!tracking) return;
      tracking = false;
      const t = ev.changedTouches[0];
      const dx = t.clientX - startX;
      li.style.transform = "";
      li.classList.remove("swiping");
      if (dx > threshold) {
        applyHistoryToForm(entry, true);
      } else if (dx < -threshold) {
        if (!confirm("Regel verwijderen uit Excel?")) return;
        await verwijderRegel(entry);
        resetFormAfterSave();
      }
    });
  }

  function renderInvoer(initialForm = false) {
    renderDatalists();
    renderHistory();
    updateInvoerStats();
    if (initialForm && !state.editRow) {
      fillForm(null);
      const og = $("#field-og")?.value;
      const pr = $("#field-project")?.value;
      if (state.intel && og && pr) {
        const t = UrenInvoer.suggestTarief(state.intel, og, pr);
        if (t !== "" && !$("#field-tarief").value) $("#field-tarief").value = t;
      }
    }
  }

  function renderChipRow(containerId, values, selected, onToggle) {
    const el = $(containerId);
    if (!el) return;
    el.innerHTML = "";
    const mk = (label, val) => {
      const b = document.createElement("button");
      b.type = "button";
      const active =
        val == null ? selected.length === 0 : selected.includes(val);
      b.className = "chip" + (active ? " active" : "");
      b.textContent = label;
      b.addEventListener("click", () => {
        onToggle(val);
        renderAnalyse();
      });
      el.appendChild(b);
    };
    mk("Alles", null);
    for (const v of values) mk(String(v), fieldKey(v));
  }

  function fieldKey(v) {
    return typeof v === "number" ? v : v;
  }

  function pruneProjectSelection() {
    const f = state.analyseFilters;
    if (!f.selectedOgs.length) return;
    const valid = new Set(
      state.entries
        .filter((r) => r.project && f.selectedOgs.includes(r.opdrachtgever))
        .map((r) => r.project)
    );
    f.selectedProjs = f.selectedProjs.filter((p) => valid.has(p));
  }

  function projectChipSource() {
    const f = state.analyseFilters;
    const rows = state.entries.filter(
      (r) =>
        r.project &&
        (f.selectedOgs.length === 0 || f.selectedOgs.includes(r.opdrachtgever))
    );
    return UrenAnalyse.sortFilterValues(rows, "project");
  }

  function tariefChipSource() {
    return UrenAnalyse.sortFilterValues(state.entries, "tarief").map((v) =>
      typeof v === "number" ? v : Number(v)
    );
  }

  function renderAnalyse() {
    const gefilterd = UrenAnalyse.filterRows(state.entries, state.analyseFilters);
    renderPeriode(gefilterd);
    renderJaarcijfers();
    if (!state.entries.length) {
      $("#analyse-summary").textContent = "geen data — ververs uit OneDrive";
      $("#analyse-list").innerHTML = "";
      $("#analyse-loc-list").innerHTML = "";
      return;
    }
    const f = state.analyseFilters;
    const ogs = UrenAnalyse.sortFilterValues(state.entries, "opdrachtgever");
    renderChipRow("#chips-og", ogs, f.selectedOgs, (val) => {
      if (val == null) f.selectedOgs = [];
      else {
        const i = f.selectedOgs.indexOf(val);
        if (i >= 0) f.selectedOgs.splice(i, 1);
        else f.selectedOgs.push(val);
      }
      pruneProjectSelection();
      renderAnalyse();
    });
    renderChipRow("#chips-proj", projectChipSource(), f.selectedProjs, (val) => {
      if (val == null) f.selectedProjs = [];
      else {
        const i = f.selectedProjs.indexOf(val);
        if (i >= 0) f.selectedProjs.splice(i, 1);
        else f.selectedProjs.push(val);
      }
    });
    const chipTr = $("#chips-tarief");
    if (chipTr) {
      chipTr.innerHTML = "";
      const mkTr = (label, onClick, active) => {
        const b = document.createElement("button");
        b.type = "button";
        b.className = "chip" + (active ? " active" : "");
        b.textContent = label;
        b.addEventListener("click", () => {
          onClick();
          const nz = $("#filter-tarief-nonzero");
          if (nz) nz.checked = f.tariefNonZero;
          renderAnalyse();
        });
        chipTr.appendChild(b);
      };
      mkTr(
        "Alles",
        () => {
          f.selectedTarieven = [];
          f.tariefNonZero = false;
        },
        !f.selectedTarieven.length && !f.tariefNonZero
      );
      mkTr(
        "Alles behalve 0",
        () => {
          f.selectedTarieven = [];
          f.tariefNonZero = !f.tariefNonZero;
        },
        f.tariefNonZero
      );
      for (const t of tariefChipSource()) {
        const active = f.selectedTarieven.includes(t);
        mkTr(
          `€${t}`,
          () => {
            f.tariefNonZero = false;
            const i = f.selectedTarieven.indexOf(t);
            if (i >= 0) f.selectedTarieven.splice(i, 1);
            else f.selectedTarieven.push(t);
          },
          active
        );
      }
    }
    const filtered = gefilterd;
    const sum = UrenAnalyse.summarize(filtered);
    const uniqueDays = UrenAnalyse.countUniqueDays(filtered);
    $("#analyse-summary").textContent =
      `${sum.count} regel${sum.count === 1 ? "" : "s"} · ${fmtU(sum.totU)} · ${fmtE0(sum.totE)} · ` +
      `${uniqueDays} dag${uniqueDays === 1 ? "" : "en"}`;
    const grouped = UrenAnalyse.groupRows(filtered, f.groupMode);
    const list = $("#analyse-list");
    list.innerHTML = "";
    for (const item of grouped) {
      const li = document.createElement("li");
      li.className = "analyse-row";
      if (item.werk != null) {
        li.innerHTML = `<div class="ar-main">${item.label} — ${item.sub}</div>
          <div class="ar-sub">${item.detail}</div>
          <div class="ar-num">${item.uren.toFixed(2)} u · €${item.bedrag.toFixed(2)}</div>
          <div class="ar-werk">${item.werk || ""}</div>`;
      } else {
        li.innerHTML = `<div class="ar-main">${item.label}</div>
          <div class="ar-num">${item.uren.toFixed(2)} u · €${item.bedrag.toFixed(2)} (${item.count} regels)</div>`;
      }
      list.appendChild(li);
    }

    const locBody = $("#analyse-loc-list");
    if (locBody) {
      locBody.innerHTML = "";
      for (const row of UrenAnalyse.aggregateLocations(filtered)) {
        const tr = document.createElement("tr");
        tr.innerHTML = `<td>${row.loc}</td><td class="num">${row.days}</td><td class="num">${row.uren.toFixed(2)}</td><td class="num">${row.bedrag.toFixed(2)}</td>`;
        locBody.appendChild(tr);
      }
    }
    renderInzichten(filtered);
    renderGrafieken();
  }

  // === Inzichten: declarabiliteit, tarieven, ranglijst en jaarbeeld ===
  const fmtU = (u) => `${(u ?? 0).toFixed(1).replace(".", ",")} u`;
  const fmtE0 = (v) => `€ ${Math.round(v ?? 0).toLocaleString("nl-NL")}`;

  function renderInzichten(filtered) {
    const I = window.UrenInzichten;
    if (!I || !$("#inz-declarabel")) return;
    const k = I.kerncijfers(filtered);
    $("#inz-periode").textContent = periodeLabel();
    const decl = $("#inz-declarabel");
    decl.textContent = k.declarabel == null ? "—" : `${String(k.declarabel).replace(".", ",")}%`;
    decl.classList.toggle("laag", k.declarabel != null && k.declarabel < 60);
    $("#inz-uren").textContent = fmtU(k.uren);
    $("#inz-omzet").textContent = fmtE0(k.omzet);
    $("#inz-effectief").textContent = k.effectiefTarief == null ? "—" : fmtE0(k.effectiefTarief);
    $("#inz-tarief").textContent = k.gemTarief == null ? "—" : fmtE0(k.gemTarief);
    $("#inz-perdag").textContent = k.urenPerDag == null ? "—" : fmtU(k.urenPerDag);
    $("#inz-uitleg").textContent = k.uren
      ? `${fmtU(k.urenBetaald)} betaald en ${fmtU(k.urenOnbetaald)} onbetaald over ${k.dagen} ` +
        `dag${k.dagen === 1 ? "" : "en"}. Effectief tarief = omzet ÷ álle gewerkte uren.`
      : "Geen uren in deze selectie.";

    // Waar gaan de onbetaalde uren heen?
    const onbetaald = I.onbetaaldPer(filtered, "werkzaamheden");
    $("#inz-onbetaald-wrap").classList.toggle("hidden", onbetaald.totaal <= 0);
    const maxOnb = Math.max(1, ...onbetaald.top.map((r) => r.uren));
    $("#inz-onbetaald").innerHTML = onbetaald.top
      .map(
        (r) => `<div class="rank-row">
          <span class="rank-naam">${esc(r.naam)}</span>
          <span class="rank-cijfer">${fmtU(r.uren)}</span>
          <span class="rank-bar"><i class="onbetaald" style="width:${(r.uren / maxOnb) * 100}%"></i></span>
        </div>`
      )
      .join("") +
      (onbetaald.aantal > onbetaald.top.length
        ? `<div class="rank-sub">+ ${onbetaald.aantal - onbetaald.top.length} andere, samen ${fmtU(onbetaald.totaal)}</div>`
        : "");

    // Ranglijst per opdrachtgever / project / locatie
    const veld = state.inzichten.rangVeld;
    document.querySelectorAll("#inz-rang-toggle .chip").forEach((c) => {
      c.classList.toggle("active", c.dataset.veld === veld);
    });
    const rang = I.ranglijst(filtered, veld);
    const maxUren = Math.max(1, ...rang.map((r) => r.uren));
    $("#inz-rang").innerHTML = rang.length
      ? rang
          .map(
            (r) => `<div class="rank-row">
              <span class="rank-naam">${esc(r.naam)}</span>
              <span class="rank-cijfer">${fmtU(r.uren)} · ${fmtE0(r.omzet)}</span>
              <span class="rank-bar"><i style="width:${(r.uren / maxUren) * 100}%"></i></span>
              <span class="rank-sub">${r.aandeel.toString().replace(".", ",")}% van de uren · ` +
              `${r.declarabel == null ? "—" : r.declarabel + "% declarabel"} · ` +
              `effectief ${r.effectiefTarief == null ? "—" : fmtE0(r.effectiefTarief)}/u</span>
            </div>`
          )
          .join("")
      : '<p class="sub">Geen uren in deze selectie.</p>';
  }

  // === Periodekaart: één plek waar je kiest waar je naar kijkt ===
  const MAAND_NAMEN = [
    "Januari", "Februari", "Maart", "April", "Mei", "Juni",
    "Juli", "Augustus", "September", "Oktober", "November", "December",
  ];

  /** Waar kijk je nu naar? In gewone taal, voor de koppen en de uitleg. */
  function periodeLabel() {
    const f = state.analyseFilters;
    if (f.periodMode === "week") return "deze week";
    if (f.periodMode === "custom_week") return `week ${f.customWeek} van ${f.customWeekYear}`;
    if (f.periodMode === "custom_month") return `${MAAND_NAMEN[(f.customMonth || 1) - 1]} ${f.customYear}`;
    if (f.periodMode === "custom_year") return `heel ${f.customYear}`;
    return "alle uren";
  }

  /** Jaar waar de jaarcijfers over gaan (bij "deze week" en "alles": dit jaar). */
  function gekozenJaar() {
    const f = state.analyseFilters;
    if (f.periodMode === "custom_week") return f.customWeekYear;
    if (f.periodMode === "custom_month" || f.periodMode === "custom_year") {
      return f.customYear || new Date().getFullYear();
    }
    return new Date().getFullYear();
  }

  /** Week vooruit/terug in de losse-week-kiezer (met jaaroverloop). */
  function stapWeek(delta) {
    const f = state.analyseFilters;
    let week = (f.customWeek || 1) + delta;
    let jaar = f.customWeekYear || new Date().getFullYear();
    if (week < 1) {
      jaar -= 1;
      week = 52;
    } else if (week > 53) {
      jaar += 1;
      week = 1;
    }
    f.customWeek = week;
    f.customWeekYear = jaar;
    const wn = $("#filter-week-num");
    const wy = $("#filter-week-year");
    if (wn) wn.value = String(week);
    if (wy) wy.value = String(jaar);
    if (f.periodMode === "custom_week") renderAnalyse();
  }

  function renderPeriode(gefilterd) {
    const f = state.analyseFilters;
    const jaar = gekozenJaar();
    const jaarInput = $("#periode-jaar");
    if (jaarInput) jaarInput.value = String(jaar);

    // "Deze maand" licht alleen op als het écht de huidige maand is; blader je
    // naar een andere maand, dan wijst het balkje de selectie aan.
    const nu = new Date();
    const isDezeMaand =
      f.periodMode === "custom_month" &&
      f.customYear === nu.getFullYear() &&
      f.customMonth === nu.getMonth() + 1;
    const preset =
      f.periodMode === "week" ? "week"
      : isDezeMaand ? "maand"
      : f.periodMode === "custom_year" ? "jaar"
      : f.periodMode === "alles" ? "alles"
      : "";
    document.querySelectorAll("#periode-presets .chip").forEach((c) => {
      c.classList.toggle("active", c.dataset.preset === preset);
    });

    // Maandbalkjes van het gekozen jaar; tikken kiest die maand.
    const maanden = window.UrenInzichten
      ? window.UrenInzichten.maandCijfers(state.entries, jaar)
      : [];
    const max = Math.max(1, ...maanden.map((m) => m.uren));
    const actief = f.periodMode === "custom_month" ? (f.customMonth || 0) - 1 : -1;
    const wrap = $("#periode-maanden");
    if (wrap) {
      wrap.innerHTML = maanden
        .map(
          (m, i) => `<div class="inz-month${actief === i ? " selected" : ""}" data-maand="${i + 1}"
              title="${MAAND_NAMEN[i]}: ${fmtU(m.uren)} · ${fmtE0(m.omzet)}">
            <span class="inz-month-bar"><i class="${m.uren ? "" : "leeg"}" style="height:${
              m.uren ? Math.max(3, (m.uren / max) * 100) : 3
            }%"></i></span>
            <span class="inz-month-label">${MONTH_LABELS[i][0]}</span>
          </div>`
        )
        .join("");
    }

    const sum = UrenAnalyse.summarize(
      gefilterd || UrenAnalyse.filterRows(state.entries, f)
    );
    const uitleg = $("#periode-uitleg");
    if (uitleg) {
      uitleg.innerHTML =
        `Je kijkt naar <strong>${periodeLabel()}</strong>: ${fmtU(sum.totU)} · ${fmtE0(sum.totE)}. ` +
        (f.periodMode === "custom_month"
          ? "<em>Tik dezelfde maand nog eens aan voor het hele jaar.</em>"
          : "<em>Tik een maand aan voor die maand.</em>");
    }
  }

  /** Jaarcijfers: urencriterium én prognose, op één plek. */
  function renderJaarcijfers() {
    const I = window.UrenInzichten;
    const jaar = gekozenJaar();
    const label = $("#uc-jaar");
    if (label) label.textContent = String(jaar);
    if (!I || !$("#uc-tekst")) return;

    const p = I.jaarPrognose(state.entries, jaar);
    const DOEL = p.doelUren;
    const fill = $("#uc-fill");
    if (fill) {
      fill.style.width = `${Math.min(100, (p.urenTotNu / DOEL) * 100).toFixed(1)}%`;
      fill.classList.toggle("uc-achter", p.isPrognose && !p.haaltCriterium);
    }
    $("#uc-tekst").textContent =
      p.urenTotNu >= DOEL
        ? `Urencriterium: ${Math.floor(p.urenTotNu)} van ${DOEL} uur — gehaald ✓`
        : p.isPrognose
          ? `Urencriterium: ${Math.floor(p.urenTotNu)} van ${DOEL} uur · ` +
            (p.haaltCriterium
              ? "in dit tempo haal je het ✓"
              : `komt ${fmtU(p.urenTekort)} tekort in dit tempo`)
          : `Urencriterium: ${Math.floor(p.urenTotNu)} van ${DOEL} uur in ${jaar}`;

    $("#inz-prog-uren").textContent = fmtU(p.uren);
    $("#inz-prog-omzet").textContent = fmtE0(p.omzet);
    $("#inz-prog-uitleg").textContent = p.isPrognose
      ? `Tot nu toe ${fmtU(p.urenTotNu)} en ${fmtE0(p.omzetTotNu)}, doorgetrokken naar heel ${jaar} ` +
        `(${Math.round(p.deel * 100)}% van het jaar is om).`
      : `Werkelijke cijfers van ${jaar}.`;
  }

  function renderAccount() {
    const acc = UrenAuth.getAccountLabel();
    $("#account-label").textContent = acc || "Niet ingelogd";
    const link = $("#onedrive-link");
    if (link && state.meta?.webUrl) {
      link.href = state.meta.webUrl;
      link.classList.remove("hidden");
    }
  }

  function switchTab(tab) {
    state.tab = tab;
    document.querySelectorAll(".tab-panel").forEach((p) => p.classList.add("hidden"));
    document.querySelectorAll(".bottom-nav button").forEach((b) => {
      b.classList.toggle("active", b.dataset.tab === tab);
    });
    const panel = document.getElementById(`panel-${tab}`);
    if (panel) panel.classList.remove("hidden");
    const sticky = $("#invoer-sticky-bar");
    const mainEl = document.querySelector("main");
    if (sticky) sticky.classList.toggle("hidden", tab !== "invoer");
    if (mainEl) mainEl.classList.toggle("has-sticky-save", tab === "invoer");
    if (vuil[tab]) tekenTab(tab);
    else if (tab === "grafieken") renderGrafieken(); // canvas hertekenen na tonen
  }

  async function onSave() {
    const fields = getFormFields();
    const err = UrenInvoer.validateForm(fields);
    if (err) {
      showToast(err, true);
      return;
    }
    if (!state.editRow) {
      const similar = UrenInvoer.findSimilarEntries(state.entries, fields);
      if (similar.length && !confirm(UrenInvoer.formatSimilarWarning(similar))) return;
    }
    const prevEntry = state.editRow
      ? state.entries.find((x) => x.row_index === state.editRow)
      : null;
    const budgetFields = { ...fields, _prevHours: prevEntry?.uren };
    const budgetMsg = UrenInvoer.budgetWarning(state.estimates, budgetFields, state.editRow);
    if (budgetMsg && !confirm(budgetMsg)) return;

    // Alles wat je ziet gebeurt nú: regel in de lijst, formulier klaar voor de
    // volgende. Het opslaan in OneDrive loopt daarna; als je ondertussen alweer
    // typt, blijft dat staan.
    const bewerkRij = state.editRow;
    state.editRow = null;
    $("#btn-save").textContent = "Opslaan";
    resetFormAfterSave();
    // Nog niet geüploade regel (offline toegevoegd): de wachtende toevoeging zelf aanpassen.
    if (bewerkRij && bewerkRij < 0) {
      if (await wijzigWachtende(bewerkRij, fields)) {
        optimisticUpdate(bewerkRij, fields);
        renderAll();
        showToast("Wachtende regel bijgewerkt, gaat mee bij de volgende sync");
      }
      return;
    }
    try {
      let gelukt;
      if (bewerkRij) {
        gelukt = await persistMutation(
          { kind: "hours_update", fields, rowIndex: bewerkRij },
          () => optimisticUpdate(bewerkRij, fields)
        );
        if (gelukt && prevEntry) volgTimetick(prevEntry, { ...fields, datumStr: fields.datumStr });
      } else {
        const tempRow = -Date.now();
        gelukt = await persistMutation(
          { kind: "hours_add", fields, rowIndex: null, tempRow },
          () => optimisticAdd(fields, tempRow)
        );
      }
      if (gelukt && !bewerkRij) biedTimetickAan(fields);
    } catch (_) {}
  }

  /** Verwijderen: wachtende (offline) regel uit de wachtrij halen, anders uit Excel; Timetick volgt. */
  async function verwijderRegel(entry) {
    const row = entry.row_index;
    if (row < 0) {
      if (await wijzigWachtende(row, null)) {
        optimisticDelete(row);
        renderAll();
        showToast("Wachtende regel weggehaald");
      }
      return;
    }
    try {
      const gelukt = await persistMutation(
        { kind: "hours_delete", fields: null, rowIndex: row },
        () => optimisticDelete(row)
      );
      if (gelukt && entry.datumStr) volgTimetick(entry, null);
    } catch (_) {}
  }

  /** Offline toegevoegde regel die nog in de wachtrij staat: aanpassen of weghalen in plaats van een los verzoek. */
  async function wijzigWachtende(tempRow, fields) {
    const items = await UrenOfflineQueue.getAll();
    const item = items.find((i) => i.kind === "hours_add" && i.tempRow === tempRow);
    if (!item) {
      showToast("Deze regel wordt nog geüpload; probeer het zo nog eens", true);
      return false;
    }
    if (fields) await UrenOfflineQueue.put({ ...item, fields });
    else await UrenOfflineQueue.remove(item.id);
    await updateQueueBadge();
    return true;
  }

  /* ------------------------------------------------------------ Timetick */

  // R2R-uren met tarief kunnen via de assistent (Pi → Cowork op de pc) ook in
  // Timetick. Adres en token komen uit de assistent-app op hetzelfde domein.
  const isR2R = (f) => /r2r/i.test(f.opdrachtgever || "") || /^\s*60\d\d/.test(f.project || "");
  const getal = (v) => parseFloat(String(v ?? "").replace(",", ".")) || 0;

  function assistentInstellingen() {
    return new Promise((ok) => {
      try {
        const r = indexedDB.open("assistent", 1);
        r.onupgradeneeded = () => r.result.createObjectStore("kv");
        r.onerror = () => ok(null);
        r.onsuccess = () => {
          try {
            const t = r.result.transaction("kv", "readonly").objectStore("kv").get("instellingen");
            t.onsuccess = () => ok(t.result || null);
            t.onerror = () => ok(null);
          } catch (_) {
            ok(null);
          }
        };
      } catch (_) {
        ok(null);
      }
    });
  }

  /* ---------------------------------------------------------------- WBSO */

  // De lijst WBSO-projecten per jaar staat op de Pi (zelfde als in de
  // assistent-app), dus wat je hier of daar instelt is overal gelijk.
  const wbsoSleutel = (naam) =>
    String(naam || "").replace(/^\s*\d{3,}\s*/, "").toLowerCase().replace(/[^a-z0-9]/g, "");

  function isWbso(e) {
    if (/wbso/i.test(e.project || "")) return true;
    const jaar = String(e.datumStr || "").slice(0, 4);
    const lijst = state.wbso?.perJaar?.[jaar];
    if (!lijst) return false;
    const k = wbsoSleutel(e.project);
    return !!k && lijst.some((n) => wbsoSleutel(n) === k);
  }

  async function piCall(pad, opties = {}) {
    const inst = await assistentInstellingen();
    if (!inst?.adres || !inst?.token) throw new Error("Stel eerst de assistent-app in (adres en token)");
    const res = await fetch(inst.adres.replace(/\/$/, "") + "/api" + pad, {
      method: opties.body ? "POST" : "GET",
      headers: { Authorization: "Bearer " + inst.token, ...(opties.body ? { "Content-Type": "application/json" } : {}) },
      body: opties.body ? JSON.stringify(opties.body) : undefined,
      cache: "no-store",
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.detail || `Fout ${res.status} van de assistent`);
    return data;
  }

  async function laadWbso() {
    const dit = new Date().getFullYear();
    state.wbso = state.wbso || { perJaar: {}, info: {} };
    for (const jaar of [dit, dit - 1]) {
      try {
        const w = await piCall("/instellingen/wbso?jaar=" + jaar);
        state.wbso.perJaar[String(jaar)] = w.projecten || [];
        state.wbso.info[String(jaar)] = w;
      } catch (_) {
        return;
      }
    }
    renderHistory();
    tekenWbsoSamenvatting();
  }

  function tekenWbsoSamenvatting() {
    const el = $("#wbso-samenvatting");
    if (!el) return;
    const jaar = String(new Date().getFullYear());
    const w = state.wbso?.info?.[jaar];
    const n = (w?.projecten || []).length;
    el.textContent = w ? `${n} ${n === 1 ? "project" : "projecten"}` : "—";
  }

  async function tekenWbsoInstellingen(jaar) {
    const el = $("#wbso-inhoud");
    if (!el) return;
    const dit = new Date().getFullYear();
    jaar = jaar || dit;
    el.innerHTML = '<p class="inst-noot">Laden…</p>';
    let w;
    try {
      w = await piCall("/instellingen/wbso?jaar=" + jaar);
    } catch (e) {
      el.innerHTML = `<p class="inst-noot">${esc(e.message)}</p>`;
      return;
    }
    const gekozen = new Set(w.projecten || []);
    const namen = [...new Set([...(w.beschikbaar || []), ...(w.projecten || [])])];
    const stand = w.stand && String(w.stand.jaar) === String(jaar)
      ? `<p class="inst-noot">Stand ${jaar}: ${String(w.stand.totaal ?? 0).replace(".", ",")} van ${w.doel} uur</p>`
      : "";
    el.innerHTML = `
      <label class="inst-rij wbso-jaar"><span class="inst-label">Jaar</span><select id="wbso-jaar" class="inst-select">${[dit - 1, dit, dit + 1]
        .map((j) => `<option value="${j}"${j === Number(w.jaar) ? " selected" : ""}>${j}</option>`)
        .join("")}</select></label>
      ${stand}
      <div id="wbso-lijst">${
        namen.length
          ? namen.map((n) => `<label class="check-row"><input type="checkbox" value="${esc(n)}"${gekozen.has(n) ? " checked" : ""} /> ${esc(n)}</label>`).join("")
          : '<p class="inst-noot">Nog geen projecten uit de urenadministratie voor dit jaar.</p>'
      }</div>
      <div class="wbso-rij"><input id="wbso-extra" type="text" placeholder="Ander project" autocomplete="off" /><button type="button" id="wbso-opslaan" class="btn-primary">Opslaan</button></div>`;
    $("#wbso-jaar").addEventListener("change", (ev) => tekenWbsoInstellingen(Number(ev.target.value)));
    $("#wbso-opslaan").addEventListener("click", async () => {
      const lijst = [...el.querySelectorAll("#wbso-lijst input:checked")].map((i) => i.value);
      const extra = $("#wbso-extra").value.trim();
      if (extra) lijst.push(extra);
      const j = Number($("#wbso-jaar").value);
      try {
        const uit = await piCall("/instellingen/wbso", { body: { jaar: j, projecten: lijst } });
        state.wbso = state.wbso || { perJaar: {}, info: {} };
        state.wbso.perJaar[String(j)] = uit.projecten || lijst;
        showToast("WBSO-projecten opgeslagen, ook in de assistent");
        tekenWbsoInstellingen(j);
        laadWbso();
      } catch (e) {
        showToast(e.message || String(e), true);
      }
    });
  }

  /** Wat de assistent weet over Timetick: verstuurd en wat er echt in staat. */
  async function laadTimetick() {
    const inst = await assistentInstellingen();
    if (!inst?.adres || !inst?.token) return;
    try {
      const res = await fetch(inst.adres.replace(/\/$/, "") + "/api/uren/timetick", {
        headers: { Authorization: "Bearer " + inst.token },
        cache: "no-store",
      });
      if (!res.ok) return;
      state.tt = await res.json();
      renderHistory();
    } catch (_) {
      /* assistent niet bereikbaar: dan zonder aanduiding */
    }
  }

  const zelfde = (r, datum, uren, nr) =>
    r.datum === datum && Math.abs(getal(r.uren) - uren) < 0.01 && (!r.nr || !nr || r.nr === nr);

  const ttNr = (e) => (/^\s*(\d{4})\b/.exec(e.project || "") || [])[1] || null;
  const ttNorm = (t) => String(t || "").toLowerCase().replace(/[^\p{L}\p{N}]+/gu, " ").trim();

  /**
   * Koppel elke regel uit Timetick (ook zelf ingevoerd, buiten de assistent om) aan hooguit één urenregel.
   * Volgorde: zelfde projectnummer, dan zelfde omschrijving, dan de rest op datum + uren.
   * Zo kan één Timetick-regel van 0:30 niet twee regels van 0:30 op dezelfde dag afvinken.
   */
  // Timetick-regels die niet in de urenadministratie horen (verlof, feestdag, prive).
  const ttNegeer = (r) => /^\s*_/.test(r.project || "") || /verlof|ziekte|feestdag/i.test(r.project || "");
  let ttCache = { tt: null, entries: null, kaart: new Map(), mist: [] };
  function ttKaart() {
    if (ttCache.tt === state.tt && ttCache.entries === state.entries) return ttCache.kaart;
    const kaart = new Map();
    const kand = (state.entries || []).filter((e) => isR2R(e) && getal(e.tarief) > 0 && getal(e.uren) > 0);
    // Regels met tarief 0 gaan normaal niet naar Timetick. Staat zo'n regel er toch in, dan is er iets mis:
    // die krijgt geen vinkje maar een waarschuwing.
    const kandNul = (state.entries || []).filter((e) => isR2R(e) && getal(e.tarief) <= 0 && getal(e.uren) > 0);
    const regels = (state.tt?.stand?.regels || []).map((r) => ({ ...r, _nr: r.nr || null, _o: ttNorm(r.omschrijving) }));
    const rondes = [
      (r, e) => r._nr && r._nr === ttNr(e),
      (r, e) => !r._nr && r._o && r._o === ttNorm(e.werkzaamheden),
      (r, e) => !r._nr || !ttNr(e),
    ];
    const vrij = new Set(regels.map((_, i) => i));
    for (const past of rondes) {
      for (const i of [...vrij]) {
        const r = regels[i];
        const e = kand.find((x) => !kaart.has(x.row_index) && x.datumStr === r.datum &&
          Math.abs(getal(r.uren) - getal(x.uren)) < 0.01 && past(r, x));
        if (e) { kaart.set(e.row_index, "in"); vrij.delete(i); }
      }
    }
    for (const i of [...vrij]) {
      const r = regels[i];
      const e = kandNul.find((x) => !kaart.has(x.row_index) && x.datumStr === r.datum &&
        Math.abs(getal(r.uren) - getal(x.uren)) < 0.01 && (!r._nr || !ttNr(x) || r._nr === ttNr(x)));
      if (e) { kaart.set(e.row_index, "tarief0"); vrij.delete(i); }
    }
    // Wat overblijft staat wel in Timetick, maar niet in de urenadministratie.
    const grens = new Date(Date.now() - 30 * 864e5).toISOString().slice(0, 10);
    const mist = [...vrij].map((i) => regels[i]).filter((r) => r.datum >= grens && !ttNegeer(r))
      .sort((a, b) => (a.datum < b.datum ? 1 : -1));
    ttCache = { tt: state.tt, entries: state.entries, kaart, mist };
    return kaart;
  }

  /** Regels die in Timetick staan zonder tegenhanger in de urenadministratie (laatste 30 dagen). */
  function ttOntbreektHier() {
    ttKaart();
    return ttCache.mist;
  }

  function ttStatus(e) {
    if (!state.tt || !isR2R(e)) return null;
    const uit = ttKaart().get(e.row_index);
    if (uit) return uit;
    if (getal(e.tarief) <= 0) return null;
    const v = (state.tt.verstuurd || []).find((r) => zelfde(r, e.datumStr, getal(e.uren), ttNr(e)));
    return v ? v.status : null;
  }

  /** Staat deze regel (datum + uren, en projectnummer als beide dat hebben) al in Timetick? */
  function staatAlInTimetick(f) {
    const nr = ttNr(f);
    return (state.tt?.stand?.regels || []).some((r) => zelfde(r, f.datumStr || f.datum, getal(f.uren), nr)) ||
      (state.tt?.verstuurd || []).some((r) => zelfde(r, f.datumStr || f.datum, getal(f.uren), nr));
  }

  function openVoorTimetick(datum) {
    return (state.entries || []).filter(
      (e) => e.datumStr === datum && isR2R(e) && getal(e.tarief) > 0 && getal(e.uren) > 0 && !ttStatus(e)
    );
  }

  function renderTimetickBalk() {
    renderTimetickMistBalk();
    const lijst = $("#history-list");
    if (!lijst) return;
    let balk = $("#tt-balk");
    if (!balk) {
      balk = document.createElement("div");
      balk.id = "tt-balk";
      balk.className = "tt-balk hidden";
      lijst.parentNode.insertBefore(balk, lijst);
    }
    const vandaag = UrenExcel.formatDateIso(new Date());
    const open = state.tt ? openVoorTimetick(vandaag) : [];
    balk.classList.toggle("hidden", !open.length);
    if (!open.length) return;
    const uren = open.reduce((t, e) => t + getal(e.uren), 0);
    balk.innerHTML = `<span>${open.length} R2R-${open.length === 1 ? "regel" : "regels"} van vandaag (${String(uren).replace(".", ",")} u) nog niet in Timetick</span><button type="button" class="tt-knop">Versturen</button>`;
    balk.querySelector("button").onclick = () => stuurNaarTimetick(open);
  }

  /** Staat er iets in Timetick wat hier niet staat, dan één regel erover; uitklappen laat zien wat. */
  function renderTimetickMistBalk() {
    const lijst = $("#history-list");
    if (!lijst) return;
    let balk = $("#tt-balk-mist");
    if (!balk) {
      balk = document.createElement("div");
      balk.id = "tt-balk-mist";
      balk.className = "tt-balk tt-balk-mist hidden";
      lijst.parentNode.insertBefore(balk, lijst);
    }
    const mist = state.tt ? ttOntbreektHier() : [];
    balk.classList.toggle("hidden", !mist.length);
    if (!mist.length) return;
    const uren = mist.reduce((t, r) => t + getal(r.uren), 0);
    const open = balk.classList.contains("open");
    balk.innerHTML = `<span>${mist.length} ${mist.length === 1 ? "regel staat" : "regels staan"} in Timetick (${String(uren).replace(".", ",")} u) maar niet hier</span>
      <button type="button" class="tt-knop tt-knop-plat">${open ? "Verberg" : "Bekijk"}</button>
      ${open ? `<ul class="tt-mist-lijst">${mist.map((r) => `<li><span>${esc(kortDatum(r.datum))}</span> ${esc(r.project || "")} · ${esc(r.omschrijving || "")} <b>${String(getal(r.uren)).replace(".", ",")} u</b></li>`).join("")}</ul>` : ""}`;
    balk.querySelector("button").onclick = () => {
      balk.classList.toggle("open");
      renderTimetickMistBalk();
    };
  }

  async function stuurNaarTimetick(entries) {
    const inst = await assistentInstellingen();
    if (!inst?.adres || !inst?.token) {
      showToast("Stel eerst de assistent-app in (adres en token)", true);
      return;
    }
    const regels = entries.map((e) => ({
      datum: e.datumStr,
      project: e.project || "",
      opdrachtgever: e.opdrachtgever || "",
      werkzaamheden: e.werkzaamheden || "",
      locatie: e.locatie || "",
      uren: getal(e.uren),
      tarief: getal(e.tarief),
    }));
    try {
      const res = await fetch(inst.adres.replace(/\/$/, "") + "/api/uren/timetick", {
        method: "POST",
        headers: { Authorization: "Bearer " + inst.token, "Content-Type": "application/json" },
        body: JSON.stringify({ regels }),
      });
      const data = await res.json().catch(() => ({}));
      if (!res.ok || !data.ok) throw new Error(data.detail || "Assistent gaf geen akkoord");
      // Meteen tonen als verstuurd; de assistent bevestigt later.
      state.tt = state.tt || { verstuurd: [], stand: { regels: [] } };
      for (const r of regels) {
        state.tt.verstuurd.push({ datum: r.datum, uren: r.uren, nr: (/^\s*(\d{4})\b/.exec(r.project) || [])[1] || null, status: "verstuurd" });
      }
      renderHistory();
      showToast(`${regels.length === 1 ? "Regel" : regels.length + " regels"} naar Timetick, komt binnen een uur`);
    } catch (e) {
      showToast("Timetick lukte niet: " + (e.message || e), true);
    }
  }

  /**
   * Een regel die al naar Timetick was gestuurd is aangepast of verwijderd. Wacht de Timetick-opdracht
   * nog, dan past de Pi die aan; staat hij al in Timetick, dan vragen we of het daar ook moet.
   */
  async function volgTimetick(oud, nieuw) {
    if (!isR2R(oud) || !ttStatus(oud)) return;
    const regel = (e) =>
      e && {
        datum: e.datumStr,
        project: e.project || "",
        opdrachtgever: e.opdrachtgever || "",
        werkzaamheden: e.werkzaamheden || "",
        locatie: e.locatie || "",
        uren: getal(e.uren),
        tarief: getal(e.tarief),
      };
    const body = { oud: regel(oud), nieuw: regel(nieuw) };
    try {
      const r = await piCall("/uren/timetick/wijzig", { body });
      if (r.status === "bijgewerkt") {
        showToast(nieuw ? "Ook aangepast in de Timetick-wachtrij" : "Ook uit de Timetick-wachtrij gehaald");
      } else if (r.status === "staat_al_in") {
        showToast(nieuw ? "Staat al in Timetick. Daar ook aanpassen?" : "Staat al in Timetick. Daar ook weghalen?", false, async () => {
          try {
            await piCall("/uren/timetick/wijzig", { body: { ...body, corrigeer: true } });
            showToast("Correctie voor Timetick staat klaar, komt binnen een uur");
          } catch (e) {
            showToast(e.message || String(e), true);
          }
        });
      }
      laadTimetick();
    } catch (_) {
      /* assistent niet bereikbaar: niets aan te doen, Timetick blijft zoals het was */
    }
  }

  async function biedTimetickAan(f) {
    if (!isR2R(f) || getal(f.uren) <= 0) return;
    if (getal(f.tarief) <= 0) {
      // Tarief 0 gaat niet naar Timetick, maar staat hij daar wel, dan klopt er iets niet.
      if (staatAlInTimetick(f)) showToast("Let op: deze regel staat wel in Timetick, maar hier op tarief 0", true);
      return;
    }
    if (staatAlInTimetick(f)) {
      showToast("Opgeslagen. Stond al in Timetick, dus niets verstuurd");
      return;
    }
    const inst = await assistentInstellingen();
    if (!inst?.adres || !inst?.token) return;
    const regel = {
      datum: f.datumStr,
      project: f.project || "",
      opdrachtgever: f.opdrachtgever || "",
      werkzaamheden: f.werkzaamheden || "",
      locatie: f.locatie || "",
      uren: getal(f.uren),
      tarief: getal(f.tarief),
    };
    showToast("Opgeslagen. Ook in Timetick zetten?", false, () =>
      stuurNaarTimetick([{ ...regel, datumStr: regel.datum }])
    );
  }

  function bindPullToRefresh() {
    const mainEl = document.querySelector("main");
    const indicator = $("#pull-indicator");
    if (!mainEl) return;
    let startY = 0;
    let pulling = false;

    mainEl.addEventListener(
      "touchstart",
      (ev) => {
        if (mainEl.scrollTop > 0 || state.loading) return;
        startY = ev.touches[0].clientY;
        pulling = true;
      },
      { passive: true }
    );

    mainEl.addEventListener(
      "touchmove",
      (ev) => {
        if (!pulling || mainEl.scrollTop > 0) return;
        const dy = ev.touches[0].clientY - startY;
        if (dy > 50) indicator?.classList.remove("hidden");
        else indicator?.classList.add("hidden");
      },
      { passive: true }
    );

    mainEl.addEventListener("touchend", async (ev) => {
      if (!pulling) return;
      pulling = false;
      const dy = ev.changedTouches[0].clientY - startY;
      indicator?.classList.add("hidden");
      if (dy > 80 && mainEl.scrollTop <= 0) {
        haptic(15);
        try {
          await refreshFromCloud();
        } catch (_) {}
      }
    });
  }

  function adjustDate(deltaDays) {
    const el = $("#field-datum");
    if (!el?.value) return;
    const d = new Date(el.value + "T12:00:00");
    d.setDate(d.getDate() + deltaDays);
    el.value = UrenExcel.formatDateIso(d);
    updateInvoerStats();
  }

  function adjustYearInput(inputId, delta) {
    const el = document.getElementById(inputId);
    if (!el) return null;
    const min = Number(el.min) || 2018;
    const max = Number(el.max) || 2035;
    let y = Number(el.value) || new Date().getFullYear();
    y = Math.min(max, Math.max(min, y + delta));
    el.value = String(y);
    return y;
  }

  function bindYearSteppers() {
    document.querySelectorAll(".btn-year-prev, .btn-year-next").forEach((btn) => {
      btn.addEventListener("click", () => {
        const inputId = btn.dataset.yearInput;
        if (!inputId) return;
        const y = adjustYearInput(inputId, btn.classList.contains("btn-year-prev") ? -1 : 1);
        if (y == null) return;
        haptic(20);
        if (inputId === "grafiek-year") {
          state.chartFilters.year = y;
          renderGrafieken();
        } else if (inputId === "filter-week-year") {
          state.analyseFilters.customWeekYear = y;
          renderAnalyse();
        } else if (inputId === "periode-jaar") {
          const f = state.analyseFilters;
          f.customYear = y;
          // Bladeren door jaren: blijf in dezelfde maand, of toon het hele jaar.
          if (f.periodMode !== "custom_month") f.periodMode = "custom_year";
          renderAnalyse();
        }
      });
    });
  }

  function bindEvents() {
    document.querySelectorAll(".bottom-nav button").forEach((btn) => {
      btn.addEventListener("click", () => switchTab(btn.dataset.tab));
    });
    // Periodekaart: presets, maandbalkjes en de extra filters
    $("#periode-presets")?.addEventListener("click", (ev) => {
      const chip = ev.target.closest(".chip");
      if (!chip) return;
      const f = state.analyseFilters;
      const nu = new Date();
      if (chip.dataset.preset === "week") {
        f.periodMode = "week";
      } else if (chip.dataset.preset === "maand") {
        f.periodMode = "custom_month";
        f.customYear = nu.getFullYear();
        f.customMonth = nu.getMonth() + 1;
      } else if (chip.dataset.preset === "jaar") {
        f.periodMode = "custom_year";
        f.customYear = gekozenJaar();
      } else {
        f.periodMode = "alles";
      }
      haptic(10);
      renderAnalyse();
    });
    $("#periode-maanden")?.addEventListener("click", (ev) => {
      const cel = ev.target.closest(".inz-month");
      if (!cel) return;
      const f = state.analyseFilters;
      const maand = Number(cel.dataset.maand);
      const zelfde = f.periodMode === "custom_month" && f.customMonth === maand;
      f.customYear = gekozenJaar();
      // Nog eens op dezelfde maand tikken = terug naar het hele jaar.
      f.periodMode = zelfde ? "custom_year" : "custom_month";
      f.customMonth = maand;
      haptic(10);
      renderAnalyse();
    });
    $("#btn-meer-filters")?.addEventListener("click", () => {
      const wrap = $("#meer-filters");
      const dicht = wrap.classList.toggle("hidden");
      $("#btn-meer-filters").textContent = dicht ? "Meer filters ▾" : "Meer filters ▴";
    });
    $("#btn-week-prev")?.addEventListener("click", () => stapWeek(-1));
    $("#btn-week-next")?.addEventListener("click", () => stapWeek(1));
    $("#btn-week-toon")?.addEventListener("click", () => {
      state.analyseFilters.periodMode = "custom_week";
      haptic(10);
      renderAnalyse();
    });
    // Inzichten: ranglijst omschakelen
    $("#inz-rang-toggle")?.addEventListener("click", (ev) => {
      const chip = ev.target.closest(".chip");
      if (!chip) return;
      state.inzichten.rangVeld = chip.dataset.veld;
      haptic(10);
      renderAnalyse();
    });
    $("#btn-save")?.addEventListener("click", onSave);
    $("#btn-clear")?.addEventListener("click", () => {
      state.editRow = null;
      state.selectedHistoryRow = null;
      $("#btn-save").textContent = "Opslaan";
      fillForm(null);
    });
    $("#btn-date-prev")?.addEventListener("click", () => adjustDate(-1));
    $("#btn-date-next")?.addEventListener("click", () => adjustDate(1));
    $("#field-datum")?.addEventListener("change", updateInvoerStats);
    $("#btn-refresh")?.addEventListener("click", () => refreshFromCloud().catch(() => {}));
    $("#btn-project-add")?.addEventListener("click", () => openProjectModal(null));
    $("#btn-est-save")?.addEventListener("click", () => onEstimateSave());
    $("#btn-est-cancel")?.addEventListener("click", closeProjectModal);
    $("#btn-est-delete")?.addEventListener("click", () => onEstimateDelete());
    $("#projecten-search")?.addEventListener("input", (e) => {
      state.estimateFilters.search = e.target.value;
      renderProjecten();
    });
    document.querySelectorAll("[data-close-modal]").forEach((el) => {
      el.addEventListener("click", closeProjectModal);
    });
    $("#btn-uren-min")?.addEventListener("click", () => adjustHours(-0.5));
    $("#btn-uren-plus")?.addEventListener("click", () => adjustHours(0.5));
    $("#field-uren")?.addEventListener("keydown", (e) => {
      if (e.key === "ArrowUp") {
        e.preventDefault();
        adjustHours(0.5);
      }
      if (e.key === "ArrowDown") {
        e.preventDefault();
        adjustHours(-0.5);
      }
    });
    $("#history-search")?.addEventListener("input", renderHistory);
    $("#field-og")?.addEventListener("input", () => onComboChange("og"));
    $("#field-project")?.addEventListener("input", () => onComboChange("og"));
    $("#field-locatie")?.addEventListener("input", () => onComboChange("project"));
    $("#btn-login")?.addEventListener("click", async () => {
      try {
        await UrenAuth.login();
        renderAccount();
        await refreshFromCloud();
      } catch (e) {
        if (e?.errorCode !== "user_cancelled") showToast(e.message, true);
      }
    });
    $("#btn-logout")?.addEventListener("click", async () => {
      await UrenAuth.logout();
      state.entries = [];
      state.estimates = [];
      state.etag = null;
      renderAccount();
      renderProjecten();
      setStatus("Uitgelogd");
    });
    const now = new Date();
    const iso = isoWeekInfo(now);
    const wy = $("#filter-week-year");
    const wn = $("#filter-week-num");
    if (wy) wy.value = iso.year;
    if (wn) wn.value = iso.week;
    state.analyseFilters.customWeekYear = iso.year;
    state.analyseFilters.customWeek = iso.week;
    state.analyseFilters.customYear = now.getFullYear();
    state.analyseFilters.customMonth = now.getMonth() + 1;
    state.chartFilters.year = now.getFullYear();
    const gy = $("#grafiek-year");
    if (gy) gy.value = now.getFullYear();
    const py = $("#periode-jaar");
    if (py) py.value = now.getFullYear();

    $("#filter-week-year")?.addEventListener("change", () => {
      state.analyseFilters.customWeekYear =
        Number($("#filter-week-year")?.value) || new Date().getFullYear();
      if (state.analyseFilters.periodMode === "custom_week") renderAnalyse();
    });
    $("#filter-tarief-nonzero")?.addEventListener("change", (e) => {
      state.analyseFilters.tariefNonZero = e.target.checked;
      if (e.target.checked) state.analyseFilters.selectedTarieven = [];
      renderAnalyse();
    });
    $("#grafiek-type")?.addEventListener("change", (e) => {
      state.chartFilters.chartMode = e.target.value;
      updateGrafiekControlsVisibility();
      renderGrafieken();
    });
    bindYearSteppers();
    $("#grafiek-cumulative-euro")?.addEventListener("change", (e) => {
      state.chartFilters.cumulativeEuro = e.target.checked;
      renderGrafieken();
    });
    $("#toggle-dark-mode")?.addEventListener("change", (e) => applyDarkMode(e.target.checked));
    $("#filter-keyword")?.addEventListener("input", (e) => {
      state.analyseFilters.keyword = e.target.value;
      renderAnalyse();
    });
    $("#filter-group")?.addEventListener("change", (e) => {
      state.analyseFilters.groupMode = e.target.value;
      renderAnalyse();
    });
    $("#week-target-input")?.addEventListener("change", (e) => saveWeekTarget(e.target.value));
    $("#btn-conflict-refresh")?.addEventListener("click", async () => {
      closeConflictModal();
      try {
        await refreshFromCloud();
      } catch (_) {}
    });
    $("#btn-conflict-dismiss")?.addEventListener("click", closeConflictModal);
    document.querySelectorAll("[data-close-conflict]").forEach((el) => {
      el.addEventListener("click", closeConflictModal);
    });
    window.addEventListener("online", () => {
      flushOfflineQueue().catch(() => {});
    });
    // Terug in beeld (van andere app teruggeschakeld): meteen bijwerken.
    document.addEventListener("visibilitychange", () => {
      if (document.visibilityState === "visible") {
        flushOfflineQueue().catch(() => {});
        laadTimetick();
      }
    });
    bindPullToRefresh();
    $("#wbso-sectie")?.addEventListener("toggle", (ev) => {
      if (ev.target.open) tekenWbsoInstellingen();
    });
    setInterval(() => {
      if (UrenAuth.isLoggedIn() && UrenOfflineQueue.isOnline() && !state.loading) {
        flushOfflineQueue().catch(() => {});
      }
    }, 120000);
  }

  async function init() {
    loadDarkPreference();
    loadWeekTarget();
    bindEvents();
    UrenCombo.createCombo("field-og", "dl-og", comboOptionsOg, () => onComboChange("og"), {
      title: "Opdrachtgever",
    });
    UrenCombo.createCombo("field-project", "dl-project", comboOptionsProj, () => onComboChange("project"), {
      title: "Project",
    });
    UrenCombo.createCombo("field-locatie", "dl-locatie", comboOptionsLoc, () => onComboChange("loc"), {
      title: "Locatie",
    });
    UrenCombo.createCombo("field-werk", "dl-werk", comboOptionsWerk, () => onComboChange("werk"), {
      title: "Werkzaamheden",
    });
    UrenInstall.init(switchTab);
    switchTab("invoer");
    await toonSnapshot();
    try {
      await UrenAuth.getMsal();
      renderAccount();
      await updateQueueBadge();
      if (UrenAuth.isLoggedIn()) await refreshFromCloud();
    } catch (e) {
      setStatus(e.message, true);
    }
    bindServiceWorker();
  }

  /**
   * De service worker start de app uit de cache en haalt updates op de
   * achtergrond op. Neemt een nieuwe versie het over, dan haal je hem met één
   * tik binnen (in plaats van bij elke start op het netwerk te wachten).
   */
  function bindServiceWorker() {
    if (!("serviceWorker" in navigator)) return;
    navigator.serviceWorker.register("./service-worker.js").catch(() => {});
    const eersteInstallatie = !navigator.serviceWorker.controller;
    let gemeld = false;
    navigator.serviceWorker.addEventListener("controllerchange", () => {
      if (eersteInstallatie || gemeld) return; // eerste keer is gewoon installeren
      gemeld = true;
      showToast("Nieuwe versie klaar, tik om te vernieuwen", false, () => (window.IMeTechHerlaad || (() => location.reload()))());
    });
    // De service worker vraagt de wachtrij weg te werken zodra er weer
    // verbinding is (Background Sync); het opslaan zelf gebeurt hier, want de
    // service worker heeft geen inlog-token.
    navigator.serviceWorker.addEventListener("message", (ev) => {
      if (ev.data?.type === "flush-queue") flushOfflineQueue().catch(() => {});
    });
  }

  /** Bij de browser aanmelden dat er nog iets te synchroniseren valt. */
  async function meldSyncAan() {
    try {
      const reg = await navigator.serviceWorker?.ready;
      await reg?.sync?.register("imtech-queue-sync");
    } catch (_) {
      /* geen Background Sync: de app werkt de wachtrij bij openen weg */
    }
  }

  document.addEventListener("DOMContentLoaded", init);

  // Kleine ingang voor debuggen en tests in de browser (geen app-logica).
  window.UrenApp = { state, switchTab, renderAll, renderAnalyse, biedTimetickAan, renderHistory, laadTimetick, laadWbso, volgTimetick, wijzigWachtende };
})();
