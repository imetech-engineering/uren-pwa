/**
 * Excel I/O via Microsoft Graph Workbook API — preserves formatting, tables and formulas.
 */
(function (global) {
  const SHEET = UrenExcel.SHEET_NAME;
  const TABLE = UrenExcel.TABLE_NAME;
  const START_ROW = UrenExcel.START_ROW;
  const COL = { DATUM: 0, WEEK: 1, JAAR: 2, OG: 3, PROJ: 4, WERK: 5, LOC: 6, UREN: 7, TARIEF: 8 };

  let cachedLayout = null;

  function encodeSheet(name) {
    return name.replace(/'/g, "''");
  }

  function wsPath(suffix) {
    return `/worksheets('${encodeSheet(SHEET)}')${suffix}`;
  }

  function workbookUrl(drivePath, suffix) {
    if (suffix.startsWith("http")) return suffix;
    return `${UrenGraph.itemUrl(drivePath)}:/workbook${suffix}`;
  }

  async function excelFetch(drivePath, token, suffix, options = {}, sessionId) {
    const url = workbookUrl(drivePath, suffix);
    const headers = { ...(options.headers || {}) };
    if (sessionId) headers["workbook-session-id"] = sessionId;
    if (options.body && !headers["Content-Type"]) {
      headers["Content-Type"] = "application/json";
    }
    const res = await fetch(url, {
      ...options,
      headers: { ...headers, Authorization: `Bearer ${token}` },
    });
    if (!res.ok) {
      const t = await res.text();
      if (res.status === 423 || res.status === 409) {
        throw new UrenGraph.GraphLockError(
          "Bestand is vergrendeld (Excel open op PC?). Sluit Excel en probeer opnieuw."
        );
      }
      throw new Error(`Excel API mislukt (${res.status}): ${t}`);
    }
    if (res.status === 204) return null;
    const ct = res.headers.get("content-type") || "";
    if (ct.includes("application/json")) return res.json();
    return res;
  }

  function parseRangeTopRow(address) {
    if (!address) return null;
    const m = address.match(/(?:'[^']+'!)?\$?([A-Z]+)\$?(\d+)/i);
    return m ? parseInt(m[2], 10) : null;
  }

  async function getTableLayout(drivePath, token, sessionId) {
    if (cachedLayout) return cachedLayout;
    const rangeData = await excelFetch(
      drivePath,
      token,
      `/tables('${encodeSheet(TABLE)}')/range`,
      {},
      sessionId
    );
    const headerRow = parseRangeTopRow(rangeData.address) || START_ROW - 1;
    const dataStartRow = headerRow + 1;
    cachedLayout = { headerRow, dataStartRow };
    return cachedLayout;
  }

  function tableIndexToExcelRow(tableIndex, dataStartRow) {
    return dataStartRow + tableIndex;
  }

  function excelRowToTableIndex(excelRow, dataStartRow) {
    return excelRow - dataStartRow;
  }

  async function withSession(drivePath, token, fn) {
    cachedLayout = null;
    const created = await excelFetch(
      drivePath,
      token,
      "/createSession",
      { method: "POST", body: JSON.stringify({ persistChanges: true }) }
    );
    const sessionId = created.id;
    try {
      return await fn(sessionId);
    } finally {
      cachedLayout = null;
      await excelFetch(
        drivePath,
        token,
        "/closeSession",
        { method: "POST", body: JSON.stringify({}) },
        sessionId
      );
    }
  }

  // De tabelrijen zijn bij het verversen toch al opgehaald; die bewaren we, zodat
  // opslaan niet nóg een keer de hele tabel hoeft te lezen. Voor het schrijven
  // controleren we altijd of de gekozen rij werkelijk leeg is.
  let rijenCache = null; // { rows, dataStartRow, at }
  const CACHE_MS = 10 * 60 * 1000;

  function zetRijenCache(rows, dataStartRow) {
    rijenCache = rows?.length ? { rows, dataStartRow, at: Date.now() } : null;
  }

  function vergeetRijenCache() {
    rijenCache = null;
  }

  function versRijenCache(dataStartRow) {
    if (!rijenCache || rijenCache.dataStartRow !== dataStartRow) return null;
    return Date.now() - rijenCache.at < CACHE_MS ? rijenCache : null;
  }

  /** Eén rij lezen om te controleren of hij nog vrij is (goedkoop). */
  async function rijIsLeeg(drivePath, token, sessionId, excelRow) {
    try {
      const data = await excelFetch(
        drivePath,
        token,
        wsPath(`/range(address='A${excelRow}:I${excelRow}')`),
        {},
        sessionId
      );
      const vals = data.values?.[0] || [];
      return rowIsEmptySlot(vals[COL.DATUM], vals);
    } catch (_) {
      return false; // niet zeker → veilige route (hele tabel lezen)
    }
  }

  /** Rij is nu bezet: in de cache markeren zodat de volgende opslag doorschuift. */
  function markeerRijGevuld(excelRow, dataStartRow) {
    if (!rijenCache) return;
    const rij = rijenCache.rows.find((r) => r.index === excelRow - dataStartRow);
    if (!rij) {
      vergeetRijenCache();
      return;
    }
    const vals = [...(rij.values || [])];
    vals[COL.DATUM] = "bezet";
    rij.values = vals;
  }

  async function fetchAllTableRows(drivePath, token, sessionId) {
    const rows = [];
    let path = `/tables('${encodeSheet(TABLE)}')/rows`;
    let seq = 0;
    while (path) {
      const data = await excelFetch(drivePath, token, path, {}, sessionId);
      for (const r of data.value || []) {
        const vals = r.values?.[0];
        if (!vals) continue;
        const idx = typeof r.index === "number" ? r.index : seq;
        rows.push({ index: idx, values: vals });
        seq = idx + 1;
      }
      const next = data["@odata.nextLink"];
      path = next || null;
    }
    rows.sort((a, b) => a.index - b.index);
    return rows;
  }

  async function readUsedRangeValues(drivePath, token, sessionId, dataStartRow) {
    const data = await excelFetch(
      drivePath,
      token,
      wsPath("/usedRange(valuesOnly=true)"),
      {},
      sessionId
    );
    const address = data.address || data.text || "";
    const match = address.match(/:I(\d+)/i);
    const endRow = match ? parseInt(match[1], 10) : dataStartRow + 500;
    if (endRow < dataStartRow) return [];
    const range = wsPath(`/range(address='A${dataStartRow}:I${endRow}')`);
    const block = await excelFetch(drivePath, token, range, {}, sessionId);
    return block.values || [];
  }

  function parseGraphDate(v) {
    if (v == null || v === "") return null;
    if (typeof v === "number") {
      const epoch = new Date(Date.UTC(1899, 11, 30));
      return new Date(epoch.getTime() + v * 86400000);
    }
    if (typeof v === "string") {
      if (v.trim().toLowerCase() === "totaal") return null;
      const d = new Date(v.slice(0, 10) + "T12:00:00");
      return Number.isNaN(d.getTime()) ? null : d;
    }
    return null;
  }

  function cellIsTotaal(v) {
    return typeof v === "string" && v.trim().toLowerCase() === "totaal";
  }

  function rowIsEmptySlot(datumVal, values) {
    if (cellIsTotaal(datumVal)) return false;
    if (datumVal != null && datumVal !== "" && String(datumVal).trim()) return false;
    if (values) {
      for (let c = COL.OG; c <= COL.TARIEF; c++) {
        const v = values[c];
        if (v == null || v === "") continue;
        if (typeof v === "number" && v === 0) continue;
        if (typeof v === "string" && !v.trim()) continue;
        return false;
      }
    }
    return true;
  }

  function rowToEntry(values, excelRow) {
    if (!values || values.length < 9) return null;
    const dt = parseGraphDate(values[COL.DATUM]);
    if (!dt) return null;
    const uren = Number(values[COL.UREN]) || 0;
    const tarief = Number(values[COL.TARIEF]) || 0;
    return {
      datum: dt,
      datumStr: UrenExcel.formatDateIso(dt),
      week: values[COL.WEEK],
      jaar: values[COL.JAAR] || dt.getFullYear(),
      opdrachtgever: String(values[COL.OG] || "").trim(),
      project: String(values[COL.PROJ] || "").trim(),
      werkzaamheden: String(values[COL.WERK] || "").trim(),
      locatie: String(values[COL.LOC] || "").trim(),
      uren,
      tarief,
      bedrag: uren * tarief,
      row_index: excelRow,
    };
  }

  async function readAllEntries(drivePath, token) {
    const layout = await getTableLayout(drivePath, token);
    const { dataStartRow } = layout;

    let tableRows = [];
    try {
      tableRows = await fetchAllTableRows(drivePath, token);
    } catch (_) {
      tableRows = [];
    }

    const entries = [];
    if (tableRows.length) {
      zetRijenCache(tableRows, dataStartRow);
      for (const tr of tableRows) {
        const excelRow = tableIndexToExcelRow(tr.index, dataStartRow);
        const entry = rowToEntry(tr.values, excelRow);
        if (entry) entries.push(entry);
      }
      return entries;
    }

    const values = await readUsedRangeValues(drivePath, token, null, dataStartRow);
    for (let i = 0; i < values.length; i++) {
      const entry = rowToEntry(values[i], dataStartRow + i);
      if (entry) entries.push(entry);
    }
    return entries;
  }

  function findAddRow(tableRows, flatValues, dataStartRow) {
    if (tableRows?.length) {
      let lastIdx = tableRows.length - 1;
      while (lastIdx >= 0) {
        const v = tableRows[lastIdx].values;
        if (v && v.some((c) => c != null && c !== "")) break;
        lastIdx--;
      }
      const totaalAtEnd =
        lastIdx >= 0 && cellIsTotaal(tableRows[lastIdx].values?.[COL.DATUM]);
      const searchEnd = totaalAtEnd ? lastIdx - 1 : lastIdx;
      for (let i = 0; i <= searchEnd; i++) {
        const d = tableRows[i].values?.[COL.DATUM];
        if (rowIsEmptySlot(d, tableRows[i].values)) {
          return {
            excelRow: tableIndexToExcelRow(tableRows[i].index, dataStartRow),
            needInsert: false,
            insertAtIndex: null,
          };
        }
      }
      const insertAtIndex = totaalAtEnd
        ? tableRows[lastIdx].index
        : null;
      const nextIndex = tableRows[lastIdx].index + 1;
      return {
        excelRow: tableIndexToExcelRow(
          insertAtIndex != null ? insertAtIndex : nextIndex,
          dataStartRow
        ),
        needInsert: true,
        insertAtIndex,
      };
    }

    const values = flatValues || [];
    let endIdx = values.length - 1;
    while (endIdx >= 0) {
      const row = values[endIdx];
      if (row && row.some((c) => c != null && c !== "")) break;
      endIdx--;
    }
    const totaalAtEnd = endIdx >= 0 && cellIsTotaal(values[endIdx]?.[COL.DATUM]);
    const searchEnd = totaalAtEnd ? endIdx - 1 : endIdx;
    for (let i = 0; i <= searchEnd; i++) {
      if (rowIsEmptySlot(values[i]?.[COL.DATUM], values[i])) {
        return {
          excelRow: dataStartRow + i,
          needInsert: false,
          insertAtIndex: null,
        };
      }
    }
    const insertIdx = totaalAtEnd ? endIdx : endIdx + 1;
    return {
      excelRow: dataStartRow + insertIdx,
      needInsert: true,
      insertAtIndex: totaalAtEnd ? endIdx : null,
    };
  }

  function fieldsToDataRow(fields) {
    return [
      fields.opdrachtgever || "",
      fields.project || "",
      fields.werkzaamheden || "",
      fields.locatie || "",
      Number(fields.uren) || 0,
      Number(fields.tarief) || 0,
    ];
  }

  async function patchRange(drivePath, token, sessionId, address, values) {
    await excelFetch(
      drivePath,
      token,
      wsPath(`/range(address='${address}')`),
      { method: "PATCH", body: JSON.stringify({ values }) },
      sessionId
    );
  }

  async function addEntry(drivePath, token, sessionId, fields) {
    const layout = await getTableLayout(drivePath, token, sessionId);
    // Snelle route: de rijen van de laatste verversing wijzen het eerste vrije
    // slot aan. Alleen als die rij echt nog leeg is schrijven we daar; anders
    // lezen we alsnog de hele tabel.
    let plek = null;
    const cache = versRijenCache(layout.dataStartRow);
    if (cache) {
      const kandidaat = findAddRow(cache.rows, null, cache.dataStartRow);
      if (!kandidaat.needInsert && (await rijIsLeeg(drivePath, token, sessionId, kandidaat.excelRow))) {
        plek = kandidaat;
      } else {
        vergeetRijenCache();
      }
    }
    if (!plek) {
      const tableRows = await fetchAllTableRows(drivePath, token, sessionId);
      zetRijenCache(tableRows, layout.dataStartRow);
      plek = findAddRow(tableRows, null, layout.dataStartRow);
    }
    const { excelRow, needInsert, insertAtIndex } = plek;

    if (needInsert) {
      const payload = {
        values: [
          [
            fields.datumStr,
            null,
            null,
            fields.opdrachtgever || "",
            fields.project || "",
            fields.werkzaamheden || "",
            fields.locatie || "",
            Number(fields.uren) || 0,
            Number(fields.tarief) || 0,
          ],
        ],
      };
      if (insertAtIndex != null) payload.index = insertAtIndex;
      await excelFetch(
        drivePath,
        token,
        `/tables('${encodeSheet(TABLE)}')/rows/add`,
        { method: "POST", body: JSON.stringify(payload) },
        sessionId
      );
      vergeetRijenCache(); // rij-indexen zijn verschoven
    } else {
      await patchRange(drivePath, token, sessionId, `A${excelRow}`, [[fields.datumStr]]);
      await patchRange(
        drivePath,
        token,
        sessionId,
        `D${excelRow}:I${excelRow}`,
        [fieldsToDataRow(fields)]
      );
      markeerRijGevuld(excelRow, layout.dataStartRow);
    }
  }

  async function updateEntry(drivePath, token, sessionId, rowIndex, fields) {
    const layout = await getTableLayout(drivePath, token, sessionId);
    if (rowIndex < layout.dataStartRow) {
      throw new Error("Regel niet meer gevonden in Excel (ververs lijsten).");
    }
    await patchRange(drivePath, token, sessionId, `A${rowIndex}`, [[fields.datumStr]]);
    await patchRange(
      drivePath,
      token,
      sessionId,
      `D${rowIndex}:I${rowIndex}`,
      [fieldsToDataRow(fields)]
    );
  }

  async function deleteEntry(drivePath, token, sessionId, rowIndex) {
    vergeetRijenCache(); // die rij komt vrij; opnieuw lezen bij de volgende opslag
    const layout = await getTableLayout(drivePath, token, sessionId);
    if (rowIndex < layout.dataStartRow) {
      throw new Error("Regel niet meer gevonden in Excel (ververs lijsten).");
    }
    // Leegschrijven i.p.v. rij DELETE — zelfde PATCH-pad als toevoegen/bewerken, geen vergrendeling.
    await patchRange(drivePath, token, sessionId, `A${rowIndex}:I${rowIndex}`, [
      ["", "", "", "", "", "", "", "", ""],
    ]);
  }

  /**
   * Vaste prijs vastzetten: kolom Tarief van alle regels van één project in één keer.
   * Eerst lezen we het blok opnieuw, zodat alleen rijen die nu nog bij dit project
   * horen een nieuw tarief krijgen; de rest van kolom I schrijven we ongewijzigd terug
   * (als formule, dus ook een eventuele formule blijft staan). Eén lees- en één schrijfactie.
   */
  async function zetTarief(drivePath, token, sessionId, project, tarief, rowIndexes) {
    const layout = await getTableLayout(drivePath, token, sessionId);
    const rijen = (rowIndexes || []).filter((r) => r >= layout.dataStartRow);
    if (!rijen.length) return 0;
    const van = Math.min(...rijen);
    const tot = Math.max(...rijen);
    const blok = await excelFetch(
      drivePath,
      token,
      wsPath(`/range(address='E${van}:I${tot}')`),
      {},
      sessionId
    );
    const doel = String(project || "").trim().toLowerCase();
    const gewenst = new Set(rijen);
    let aantal = 0;
    const kolomI = (blok.formulas || []).map((rij, i) => {
      const excelRow = van + i;
      const proj = String(blok.values?.[i]?.[0] ?? "").trim().toLowerCase();
      if (gewenst.has(excelRow) && proj === doel) {
        aantal++;
        return [Number(tarief) || 0];
      }
      return [rij[4] ?? ""];
    });
    if (!aantal) throw new Error("De regels van dit project zijn verschoven — ververs en probeer opnieuw.");
    await excelFetch(
      drivePath,
      token,
      wsPath(`/range(address='I${van}:I${tot}')`),
      { method: "PATCH", body: JSON.stringify({ formulas: kolomI }) },
      sessionId
    );
    return aantal;
  }

  global.UrenGraphExcel = {
    readAllEntries,
    zetTarief,
    addEntry,
    updateEntry,
    deleteEntry,
    withSession,
  };
})(window);
