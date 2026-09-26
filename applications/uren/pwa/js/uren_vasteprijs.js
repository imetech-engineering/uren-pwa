/**
 * Vaste prijs: het uurtarief volgt uit de offerte, niet uit de urenregels.
 *
 * Het bedrag voor uren (materiaal telt niet mee) komt uit:
 *   1. Ureninschattingen, kolom "Offerte uren (€)": vastgelegd (bij afronden of door jou gekozen);
 *   2. anders automatisch: de offertes die projectdoc aan dit projectnummer koppelt, plus
 *      meerwerkfacturen (facturen van het project die niet bij een offerte horen).
 * "-" in kolom Offerte betekent: bewust geen vaste prijs, niets automatisch.
 *
 * Tarief:
 *   - lopend project:  bedrag ÷ max(ureninschatting, gemaakte uren)
 *     (omzet naar voortgang; loop je uit, dan zakt het tarief mee)
 *   - afgerond:        bedrag ÷ gemaakte uren
 * Regie-projecten houden altijd het tarief uit de regels.
 *
 * De analyse rekent live met dit tarief; de regels in Excel krijgen het pas als je
 * het bij afronden vastzet.
 */
(function (global) {
  const sleutel = (p) => String(p || "").trim().toLowerCase();
  const rond2 = (n) => Math.round(n * 100) / 100;
  const projectNr = (p) => (/^\s*(\d{4})\b/.exec(p || "") || [])[1] || null;
  const UIT = "-";

  /* ---------------------------------------------- bronnen: offertes en meerwerk */

  /** Offertes en meerwerkfacturen van de bridge als één soort "bron" met regels. */
  function bronnen(geld) {
    const uit = [];
    for (const o of geld?.offertes || []) {
      uit.push({
        soort: "offerte",
        nummer: o.nummer,
        datum: o.datum || "",
        klant: o.klant || "",
        titel: o.onderwerp || "",
        zoek: `${o.onderwerp || ""} ${o.referentie || ""}`,
        totaal: Number(o.totaalExcl) || 0,
        regels: (o.regels || []).map((r) => ({ omschrijving: r.omschrijving, bedrag: Number(r.bedrag) || 0 })),
        regie: !!o.regie,
        status: o.status || "",
        projectNummer: o.projectNummer || null,
      });
    }
    for (const f of geld?.meerwerk || []) {
      uit.push({
        soort: "factuur",
        nummer: f.nummer,
        datum: f.datum || "",
        klant: f.klant || "",
        titel: f.omschrijving || "",
        zoek: f.omschrijving || "",
        totaal: Number(f.netto) || 0,
        regels: [{ omschrijving: f.omschrijving || "Meerwerk", bedrag: Number(f.netto) || 0 }],
        regie: false,
        status: "",
        projectNummer: f.projectNummer || null,
      });
    }
    return uit;
  }

  /* ---------------------------------------------- regels: uren of materiaal */

  // Altijd kosten, ook als er een arbeidswoord in staat ("licentie software", "verzendkosten montageset").
  const KOSTEN =
    /\b(licenties?|verzend\w*|verzending|transport\w*|vracht\w*|inkoop\w*|aanschaf\w*|doorbelast\w*)\b/i;
  // Arbeid: dan telt de regel als uren ("montage onderdelen" = uren).
  const ARBEID =
    /\b(uren|uur|arbeid|engineering|ontwerp\w*|ontwikkel\w*|advies|onderzoek|programmer\w*|software|firmware|test\w*|montage|installatie|inbedrijfstelling|\w*begeleiding|projectmanagement|documentatie|tekenwerk|werk(zaamheden)?|meerwerk)\b/i;
  // Duidelijk materiaal. Twijfel → wel aangevinkt; je ziet het en vinkt zelf af.
  const MATERIAAL =
    /\b(materi(aa|a)l\w*|onderdel\w*|componenten|component|hardware|pcb\w*|printplat\w*|behuizing\w*|kabel\w*|voeding\w*|sensor\w*|motor\w*|stuks?|3d[- ]?print\w*|bom|bill of materials|levering\w*)\b/i;

  /** Voorstel: telt deze regel als uren? */
  function isUrenRegel(omschrijving) {
    const t = String(omschrijving || "");
    if (KOSTEN.test(t)) return false;
    if (ARBEID.test(t)) return true;
    return !MATERIAAL.test(t);
  }

  /** Welke regels tellen als uren: jouw eerdere keuze, anders het voorstel op omschrijving. */
  function urenVinkjes(bron, vinkjes) {
    const bewaard = vinkjes?.[bron.nummer];
    if (Array.isArray(bewaard)) return bewaard;
    return bron.regels.map((r, i) => (isUrenRegel(r.omschrijving) ? i : -1)).filter((i) => i >= 0);
  }

  function urenBedrag(gekozen, vinkjes) {
    let som = 0;
    for (const b of gekozen) {
      const v = new Set(urenVinkjes(b, vinkjes));
      b.regels.forEach((r, i) => {
        if (v.has(i)) som += r.bedrag;
      });
    }
    return rond2(som);
  }

  /**
   * Automatisch voorstel voor een project met projectnummer: de offertes van dat nummer
   * (niet vervangen of vervallen) plus het meerwerk. Zonder offerte, of met een regie-offerte,
   * geen voorstel: dan is het geen vaste prijs.
   */
  function autoVoorstel(project, lijst, vinkjes) {
    const nr = projectNr(project);
    if (!nr) return null;
    const offertes = lijst.filter(
      (b) => b.soort === "offerte" && b.projectNummer === nr && !["vervangen", "vervallen"].includes(b.status)
    );
    if (!offertes.length || offertes.some((b) => b.regie)) return null;
    const meerwerk = lijst.filter((b) => b.soort === "factuur" && b.projectNummer === nr);
    const gekozen = [...offertes, ...meerwerk];
    const bedrag = urenBedrag(gekozen, vinkjes);
    if (!(bedrag > 0)) return null;
    return { nummers: gekozen.map((b) => b.nummer), bedrag, bronnen: gekozen };
  }

  /**
   * Het bedrag waarmee gerekend wordt: vastgelegd in Excel wint, dan automatisch.
   * → { bron: "excel"|"auto", bedrag, offerte } of null.
   */
  function effectief(est, auto) {
    if (!est || est.status === "Regie") return null;
    if (String(est.offerte || "").trim() === UIT) return null;
    if ((Number(est.offerteUren) || 0) > 0) {
      return { bron: "excel", bedrag: Number(est.offerteUren), offerte: est.offerte || "" };
    }
    if (auto) return { bron: "auto", bedrag: auto.bedrag, offerte: auto.nummers.join(", "), auto };
    return null;
  }

  /* ---------------------------------------------- tarief */

  /** Gemaakte uren per project uit de urenregels (zelfde koppeling als de SUMIFS in Excel). */
  function urenPerProject(entries) {
    const uit = new Map();
    for (const e of entries || []) {
      const k = sleutel(e.project);
      if (!k) continue;
      uit.set(k, (uit.get(k) || 0) + (Number(e.uren) || 0));
    }
    return uit;
  }

  /** Tarief voor een bedrag, of null als er (nog) niets te rekenen valt. est: { ureninschatting, status }. */
  function tariefVoor(bedrag, est, gemaakt) {
    if (!(bedrag > 0) || !est || est.status === "Regie") return null;
    const geschat = Number(est.ureninschatting) || 0;
    const noemer = est.status === "Afgerond" ? gemaakt : Math.max(geschat, gemaakt);
    return noemer > 0 ? rond2(bedrag / noemer) : null;
  }

  /**
   * Map project(sleutel) → { tarief, bedrag, bron, offerte, gemaakt, geschat, est } voor alle
   * vaste-prijsprojecten. geld = antwoord van de bridge (mag leeg zijn), vinkjes = jouw keuzes per nummer.
   */
  function tarieven(estimates, entries, geld = null, vinkjes = null) {
    const uren = urenPerProject(entries);
    const lijst = bronnen(geld);
    const uit = new Map();
    for (const est of estimates || []) {
      const e = effectief(est, lijst.length ? autoVoorstel(est.project, lijst, vinkjes) : null);
      if (!e) continue;
      const k = sleutel(est.project);
      const gemaakt = uren.get(k) || 0;
      const tarief = tariefVoor(e.bedrag, est, gemaakt);
      if (tarief == null) continue;
      uit.set(k, { ...e, tarief, gemaakt, geschat: Number(est.ureninschatting) || 0, est });
    }
    return uit;
  }

  /**
   * Kopie van de regels met het offertetarief ingevuld (voor analyse en grafieken).
   * tariefRegel bewaart wat er in Excel staat.
   */
  function pasToe(entries, map) {
    if (!map || !map.size) return entries;
    return (entries || []).map((e) => {
      const v = map.get(sleutel(e.project));
      if (!v) return e;
      const uren = Number(e.uren) || 0;
      return { ...e, tarief: v.tarief, bedrag: uren * v.tarief, tariefRegel: e.tarief, vastePrijs: true };
    });
  }

  /** Regels van dit project waarvan het tarief in Excel afwijkt van het offertetarief. */
  function afwijkend(entries, project, tarief) {
    const k = sleutel(project);
    return (entries || []).filter(
      (e) => sleutel(e.project) === k && Math.abs((Number(e.tarief) || 0) - tarief) >= 0.005
    );
  }

  global.UrenVastePrijs = {
    UIT,
    projectNr,
    bronnen,
    urenVinkjes,
    urenBedrag,
    autoVoorstel,
    effectief,
    urenPerProject,
    tariefVoor,
    tarieven,
    pasToe,
    afwijkend,
    isUrenRegel,
    sleutel,
  };
})(window);
