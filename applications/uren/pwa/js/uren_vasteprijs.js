/**
 * Vaste prijs: het uurtarief volgt uit de offerte, niet uit de urenregels.
 *
 * Per project staat in Ureninschattingen welk bedrag van de offerte voor uren is
 * (materiaal telt niet mee). Het tarief is dan:
 *   - lopend project:  offertebedrag ÷ max(ureninschatting, gemaakte uren)
 *     (omzet naar voortgang; loop je uit, dan zakt het tarief mee)
 *   - afgerond:        offertebedrag ÷ gemaakte uren
 * Regie-projecten houden altijd het tarief uit de regels.
 *
 * De analyse rekent live met dit tarief; de regels in Excel krijgen het pas als je
 * het bij afronden vastzet.
 */
(function (global) {
  const sleutel = (p) => String(p || "").trim().toLowerCase();
  const rond2 = (n) => Math.round(n * 100) / 100;

  /** Is dit project op vaste prijs (offertebedrag bekend en geen regie)? */
  function isVastePrijs(est) {
    return !!est && (Number(est.offerteUren) || 0) > 0 && est.status !== "Regie";
  }

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

  /** Tarief voor één project, of null als er (nog) niets te rekenen valt. */
  function tariefVoor(est, gemaakt) {
    if (!isVastePrijs(est)) return null;
    const bedrag = Number(est.offerteUren);
    const geschat = Number(est.ureninschatting) || 0;
    const noemer = est.status === "Afgerond" ? gemaakt : Math.max(geschat, gemaakt);
    return noemer > 0 ? rond2(bedrag / noemer) : null;
  }

  /** Map project(sleutel) → { tarief, bedrag, gemaakt, geschat, est } voor alle vaste-prijsprojecten. */
  function tarieven(estimates, entries) {
    const uren = urenPerProject(entries);
    const uit = new Map();
    for (const est of estimates || []) {
      if (!isVastePrijs(est)) continue;
      const k = sleutel(est.project);
      const gemaakt = uren.get(k) || 0;
      const tarief = tariefVoor(est, gemaakt);
      if (tarief == null) continue;
      uit.set(k, {
        tarief,
        bedrag: Number(est.offerteUren),
        gemaakt,
        geschat: Number(est.ureninschatting) || 0,
        est,
      });
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

  /* ---------------------------------------------- offerteregels: uren of materiaal */

  // Altijd kosten, ook als er een arbeidswoord in staat ("licentie software", "verzendkosten montageset").
  const KOSTEN =
    /\b(licenties?|verzend\w*|verzending|transport\w*|vracht\w*|inkoop\w*|aanschaf\w*|doorbelast\w*)\b/i;
  // Arbeid: dan telt de regel als uren ("montage onderdelen" = uren).
  const ARBEID =
    /\b(uren|uur|arbeid|engineering|ontwerp\w*|ontwikkel\w*|advies|onderzoek|programmer\w*|software|firmware|test\w*|montage|installatie|inbedrijfstelling|\w*begeleiding|projectmanagement|documentatie|tekenwerk|werk(zaamheden)?)\b/i;
  // Duidelijk materiaal. Twijfel → wel aangevinkt; je ziet het en vinkt zelf af.
  const MATERIAAL =
    /\b(materi(aa|a)l\w*|onderdel\w*|componenten|component|hardware|pcb\w*|printplat\w*|behuizing\w*|kabel\w*|voeding\w*|sensor\w*|motor\w*|stuks?|3d[- ]?print\w*|bom|bill of materials|levering\w*)\b/i;

  /** Voorstel: telt deze offerteregel als uren? */
  function isUrenRegel(omschrijving) {
    const t = String(omschrijving || "");
    if (KOSTEN.test(t)) return false;
    if (ARBEID.test(t)) return true;
    return !MATERIAAL.test(t);
  }

  global.UrenVastePrijs = {
    isVastePrijs,
    urenPerProject,
    tariefVoor,
    tarieven,
    pasToe,
    afwijkend,
    isUrenRegel,
    sleutel,
  };
})(window);
