// ════════════════════════════════════════════════════════════════════
//  PRIME RBM · DTC ASSISTED — une slide par mois : le MEILLEUR des 7 RBM.
//
//  RÈGLE (tranchée par le propriétaire le 2026-09-03) :
//   • critère = VOLUME POS DU MOIS, et rien d'autre. C'est déjà le critère
//     de la prime TSA (slide « Prime & Rang » = VOLUME_XAF) ; le RBM est
//     primé sur le même barème que ses TSA, pas sur un autre.
//   • UN SEUL primé par mois — le n°1 des 7. Les six autres sont classés,
//     pas récompensés.
//   • juin est HORS calcul : le programme n'a démarré que le 15/06, le mois
//     n'est pas plein. Seuls les mois PLEINS comptent (juillet, août, …).
//
//  🔴 LE MOIS SE COMPTE EN JOURS, PAS EN SEMAINES. Corrigé le 2026-09-03 après
//  remarque du propriétaire. La première version sommait les semaines POS qui
//  chevauchent le mois : la fenêtre réelle de juillet faisait alors 42 jours
//  (26/06 → 06/08) et gonflait le Nord Ouest de 19,6 M à 25,9 M. La source est
//  désormais `pos_daily.csv` (= `lka_client_mtn.pos_sales`, une ligne par POS et
//  par jour). Voir `lib/pos_jour.js` pour la mesure complète.
// ════════════════════════════════════════════════════════════════════
const path = require("path");
const T = require("./theme");
const { posDuMois } = require("./pos_jour");
const {
  INK, INK2, YEL, YELDK, YELSOFT, MUTE, MUTE2, ROW, GREEN, AMBER, REDX,
  fHead, fBody, MX, CW, REGIONS, PROGRAMME_START,
  readCSV, num, fmt, MOIS, toTitle, light, chrome, header, table,
} = T;

// Un mois plein : ses 7 RBM classés par volume POS décroissant.
// `dir` = dossier des CSV (classement.csv, tsa_ref.csv).
function classementRBM(dir, yy, mm) {
  const tsaRef = readCSV(path.join(dir, "tsa_ref.csv"));
  const P = posDuMois(dir, yy, mm, PROGRAMME_START);

  // région → RBM : l'annuaire TSA fait foi (les ventes peuvent porter des trous).
  const rbmVotes = {}, objDaily = {};
  REGIONS.forEach(r => { rbmVotes[r] = {}; objDaily[r] = 0; });
  tsaRef.forEach(r => {
    const reg = (r.REGION || "").trim().toUpperCase();
    if (!REGIONS.includes(reg)) return;
    const nom = (r.RBM || "").trim();
    if (nom) rbmVotes[reg][nom] = (rbmVotes[reg][nom] || 0) + 1;
    objDaily[reg] += num(r.OBJECTIF_DAILY);
  });
  const rbmDe = reg => (Object.entries(rbmVotes[reg]).sort((a, b) => b[1] - a[1])[0] || ["—"])[0];

  // meilleur TSA de chaque région, sur le cumul du mois compté au jour
  const meilleurTSA = {}, nbTSA = {};
  REGIONS.forEach(r => { meilleurTSA[r] = null; nbTSA[r] = 0; });
  P.volByTSA.forEach(t => {
    nbTSA[t.REGION]++;
    if (!meilleurTSA[t.REGION] || t.VOLUME_XAF > meilleurTSA[t.REGION].VOLUME_XAF) meilleurTSA[t.REGION] = t;
  });

  const lignes = REGIONS.map(reg => {
    // ⚠️ cible = journalier × jours RÉELLEMENT présents dans pos_sales, pas jours
    // du calendrier : sinon 30 j de réalisé sont jugés sur 31 j de cible.
    const objectif = objDaily[reg] * P.jours;
    return {
      region: reg, rbm: rbmDe(reg), volume: P.volByReg[reg], objectif,
      taux: objectif ? P.volByReg[reg] / objectif * 100 : 0,
      posActifs: P.posActifs[reg], nbTSA: nbTSA[reg],
      topTSA: meilleurTSA[reg] ? { nom: meilleurTSA[reg].TSA, v: meilleurTSA[reg].VOLUME_XAF } : null,
    };
    // 🔴 Départage EXPLICITE : sans lui, deux passages sur le même fichier
    // peuvent élire deux RBM différents (même faute que l'élection des BA).
  }).sort((a, b) => b.volume - a.volume || b.taux - a.taux || a.region.localeCompare(b.region));

  lignes.forEach((l, i) => l.rang = i + 1);
  return {
    yy, mm, jours: P.jours, posStart: P.premier || P.debut, end: P.dernier || P.fin,
    finCalendaire: P.fin, lignes,
    label: `${MOIS[mm - 1].toUpperCase()} ${yy}`,
    total: P.volTotal,
    objectifTotal: lignes.reduce((s, l) => s + l.objectif, 0),
  };
}

const dmy = d => `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")}/${d.getFullYear()}`;

// Dessine la slide. `page` = numéro affiché dans le chrome.
function slidePrimeRBM(pres, M, page) {
  const s = pres.addSlide(); light(s);
  chrome(pres, s, page, "Reconnaissance · Prime RBM mensuelle");
  header(pres, s, "Prime RBM", `Le meilleur des 7 RBM · ${M.label} · ${dmy(M.posStart)} au ${dmy(M.end)}`);

  const g = M.lignes[0], second = M.lignes[1] || { volume: 0, region: "—" };
  const avance = second.volume ? (g.volume / second.volume - 1) * 100 : 0;

  // ── Bandeau lauréat ──
  const by = 1.52, bh = 1.32;
  s.addShape(pres.shapes.RECTANGLE, { x: MX, y: by, w: CW, h: bh, fill: { color: YELSOFT }, line: { type: "none" } });
  s.addShape(pres.shapes.RECTANGLE, { x: MX, y: by, w: 0.12, h: bh, fill: { color: YELDK }, line: { type: "none" } });
  s.addText("RBM DU MOIS", { x: MX + 0.34, y: by + 0.14, w: 4.2, h: 0.26, margin: 0, fontFace: fBody, fontSize: 9, color: MUTE, charSpacing: 3 });
  s.addText(toTitle(g.rbm), { x: MX + 0.3, y: by + 0.36, w: 6.2, h: 0.55, margin: 0, fontFace: fHead, bold: true, fontSize: 30, color: INK, valign: "middle" });
  s.addText(`${g.region} · ${g.nbTSA} TSA · ${fmt(g.posActifs)} POS actifs`, {
    x: MX + 0.32, y: by + 0.92, w: 6.2, h: 0.28, margin: 0, fontFace: fBody, fontSize: 11, color: INK2,
  });

  const stats = [
    { l: "VOLUME POS DU MOIS", v: fmt(g.volume), u: "XAF" },
    { l: "TAUX D'ATTEINTE", v: g.taux.toFixed(1) + "%", u: `cible ${fmt(g.objectif)}` },
    { l: "AVANCE SUR LE 2e", v: "+" + avance.toFixed(0) + "%", u: toTitle(second.region) },
  ];
  const sw = 1.9, sx0 = MX + CW - 0.28 - stats.length * sw;
  stats.forEach((k, i) => {
    const x = sx0 + i * sw;
    s.addText(k.l, { x, y: by + 0.24, w: sw - 0.12, h: 0.24, margin: 0, fontFace: fBody, fontSize: 8, color: MUTE, charSpacing: 1.2 });
    s.addText(k.v, { x, y: by + 0.46, w: sw - 0.12, h: 0.46, margin: 0, fontFace: fHead, bold: true, fontSize: 20, color: YELDK, valign: "middle" });
    s.addText(k.u, { x, y: by + 0.92, w: sw - 0.12, h: 0.24, margin: 0, fontFace: fBody, fontSize: 8.5, color: MUTE2 });
  });

  // ── Règle du barème ──
  const ry = by + bh + 0.16;
  s.addShape(pres.shapes.RECTANGLE, { x: MX, y: ry, w: CW, h: 0.38, fill: { color: ROW }, line: { type: "none" } });
  s.addShape(pres.shapes.RECTANGLE, { x: MX, y: ry, w: 0.08, h: 0.38, fill: { color: YEL }, line: { type: "none" } });
  s.addText([
    { text: "Critère : ", options: { fontFace: fHead, bold: true, color: INK } },
    { text: `volume POS du mois — un seul primé sur les 7. Mois PLEINS uniquement (juin exclu : programme démarré le 15/06). Compté au JOUR : ${M.jours} jours de ventes, du ${dmy(M.posStart)} au ${dmy(M.end)}.`, options: { fontFace: fBody, color: INK2 } },
  ], { x: MX + 0.22, y: ry, w: CW - 0.4, h: 0.38, margin: 0, valign: "middle", fontSize: 10 });

  // ── Classement des 7 ──
  const rows = M.lignes.map(l => [
    { text: String(l.rang), bold: l.rang === 1, color: l.rang === 1 ? YELDK : INK2 },
    { text: toTitle(l.rbm), bold: l.rang === 1, color: l.rang === 1 ? INK : INK2 },
    l.region,
    fmt(l.volume),
    { text: l.taux.toFixed(1) + "%", color: l.taux >= 100 ? GREEN : l.taux >= 30 ? AMBER : REDX, bold: true },
    fmt(l.posActifs),
    l.topTSA ? toTitle(l.topTSA.nom) : "—",
    { text: l.rang === 1 ? "PRIMÉ" : "—", bold: l.rang === 1, color: l.rang === 1 ? GREEN : MUTE2 },
  ]);
  const tblY = ry + 0.58;
  table(pres, s, MX, tblY, CW,
    ["#", "RBM", "RÉGION", "VOLUME POS (XAF)", "ATTEINTE", "POS ACTIFS", "MEILLEUR TSA", "PRIME"], rows, {
      colWidths: [CW * 0.05, CW * 0.21, CW * 0.12, CW * 0.16, CW * 0.10, CW * 0.10, CW * 0.18, CW * 0.08],
      rowH: 0.40, fs: 10, hfs: 9,
    });

  s.addText(
    `Total programme ${M.label} : ${fmt(M.total)} XAF sur une cible de ${fmt(M.objectifTotal)} XAF ` +
    `(${(M.objectifTotal ? M.total / M.objectifTotal * 100 : 0).toFixed(1)} %). ` +
    `Cible régionale = somme des OBJECTIF_DAILY des TSA de la région × ${M.jours} jours de ventes réellement présents dans pos_sales.`,
    { x: MX, y: tblY + 0.40 * 8 + 0.16, w: CW, h: 0.3, margin: 0, fontFace: fBody, fontSize: 9, color: MUTE });
  return s;
}

module.exports = { classementRBM, slidePrimeRBM };
