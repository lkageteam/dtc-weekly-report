// ════════════════════════════════════════════════════════════════════
//  PROGRESS REPORT · DTC ASSISTED · MTN BÉNIN · Contribution LKA
//  Rapport MENSUEL — même deck que le rapport hebdomadaire, mais tous les
//  agrégats sont calculés sur un MOIS calendaire (borné au démarrage réel
//  du programme) au lieu d'une semaine.
//
//  Lancement : node scripts/monthly_report_dtc_assisted.js [YYYY-MM]
//   (sans argument → mois du dernier Timestamp trouvé dans form.csv)
//
//  Exemple : node scripts/monthly_report_dtc_assisted.js 2026-06
//   → fenêtre réelle = 15 au 30 juin 2026 (le programme n'a démarré que le 15).
//   Les mois suivants (2026-07, 2026-08, …) couvrent le mois calendaire complet.
//
//  SLIDES (8) — miroir de l'hebdo, hors « Retours terrain » :
//   1 Couverture · 2 Recrutement · 3 Activations POS · 4 Prime & Rang (mois)
//   5 Activations BA · 6 Month-over-month · 7 Analyse BA · 8 Clôture
//   La slide « Retours terrain » de l'hebdo n'existe pas ici : elle vient de la
//   colonne « problème principal » de l'export Google Form, qui n'est PAS mirrorée
//   dans la table MySQL `dtc_ba_activations` (seule source des données mensuelles).
//
//  RÈGLE (confirmée avec l'utilisateur, cf. mémoire [[progress-report-dtc-assisted]]) :
//   objectif du mois = objectif JOURNALIER (même taux que l'hebdo) × nombre de jours
//   de la fenêtre — jours ouvrés (hors dimanche) pour le BA, tous les jours pour le POS.
//   Réalisé BA = form filtré par Timestamp sur la fenêtre.
//   Réalisé POS = somme des "Semaine N" du classement dont la période chevauche la fenêtre
//   (le classement n'a pas de date par ligne, seulement une semaine POS ven→jeu) — une
//   semaine à cheval sur deux mois compte donc dans les deux.
//
//  Données dans INPUT_DIR (défaut ./inputs), mêmes fichiers que le rapport hebdo :
//   classement.csv (CLASSEMENT_TSA) · tsa_ref.csv (TSA_REF) · form.csv (Form BA / MySQL)
// ════════════════════════════════════════════════════════════════════
const pptxgen = require("pptxgenjs");
const { withPPTXEmbedFonts } = require("pptx-embed-fonts/pptxgenjs");
const fs = require("fs");
const path = require("path");
const T = require("./lib/theme");
const {
  BG, INK, INK2, YEL, YELDK, YELSOFT, MUTE, MUTE2, LINE, ROW, GREEN, AMBER, REDX,
  fHead, fBody, HAS_LOGO, LOGO, W, H, MX, CW, REGIONS, CONFIG, PROGRAMME_START, BA_START,
  readCSV, num, parseTS, MOIS, fmt, toTitle, light, chrome, header, table, embedFonts,
} = T;

const IN = process.env.INPUT_DIR ? path.resolve(process.env.INPUT_DIR) : path.join(__dirname, "..", "inputs");
const MS = 86400000;

// ── Paramètre mois : "YYYY-MM". Sans argument → mois du dernier Timestamp du form. ──
function latestMonthFromForm() {
  try {
    const form = readCSV(path.join(IN, "form.csv"));
    const H0 = Object.keys(form[0] || {});
    const cTS = H0.find(h => /Timestamp|Horodat/i.test(h));
    let maxD = null;
    form.forEach(r => { const d = parseTS(r[cTS]); if (d && (!maxD || d > maxD)) maxD = d; });
    if (!maxD) return null;
    return `${maxD.getFullYear()}-${String(maxD.getMonth() + 1).padStart(2, "0")}`;
  } catch { return null; }
}
const _arg = (process.argv[2] || "").trim();
const YM = /^\d{4}-\d{2}$/.test(_arg) ? _arg : (latestMonthFromForm() || `${new Date().getFullYear()}-${String(new Date().getMonth() + 1).padStart(2, "0")}`);
const [YY, MM] = YM.split("-").map(Number);

// ── Fenêtres d'un mois : bornées au démarrage réel du programme ──
// Deux ancres décalées de 3 jours (comme le rapport hebdo) : POS démarre le 12 juin,
// BA le 15 juin. Chaque flux a donc SA PROPRE fenêtre, clampée à SON propre départ —
// ne jamais utiliser une fenêtre unique pour les deux (le POS perdrait 3 jours de cible/réalisé).
function monthWindow(yy, mm) {
  const first = new Date(yy, mm - 1, 1);
  const end = new Date(yy, mm, 0);                    // dernier jour du mois
  const posStart = first < PROGRAMME_START ? PROGRAMME_START : first;
  const baStart = first < BA_START ? BA_START : first;
  let workDays = 0;
  for (let t = baStart.getTime(); t <= end.getTime(); t += MS) if (new Date(t).getDay() !== 0) workDays++;
  return {
    yy, mm, end, endTS: new Date(end.getTime() + MS - 1),
    posStart, baStart,
    posDays: Math.round((end - posStart) / MS) + 1,
    baDays: Math.round((end - baStart) / MS) + 1,
    workDays,
    label: `${MOIS[mm - 1].toUpperCase()} ${yy}`,
    short: MOIS[mm - 1].charAt(0).toUpperCase() + MOIS[mm - 1].slice(1),
    partial: first < BA_START,                        // mois tronqué par le démarrage du programme
  };
}
const dmy = d => `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")}/${d.getFullYear()}`;

const WIN = monthWindow(YY, MM);
const MONTH_LABEL = WIN.label;
const WINDOW_LABEL = `${dmy(WIN.baStart)} au ${dmy(WIN.end)}`;      // fenêtre BA
const POS_WINDOW_LABEL = `${dmy(WIN.posStart)} au ${dmy(WIN.end)}`; // fenêtre POS
const REPORT_DATE = dmy(new Date(WIN.end.getTime() + MS));
console.log(`📅 Rapport mensuel : ${MONTH_LABEL} · BA ${WINDOW_LABEL} (${WIN.baDays}j, ${WIN.workDays}j ouvrés) · POS ${POS_WINDOW_LABEL} (${WIN.posDays}j)`);

// ════════════════════════════════════════════════════════════════════
//  DONNÉES
// ════════════════════════════════════════════════════════════════════
const tsaRef = readCSV(path.join(IN, "tsa_ref.csv"));
const posDaily = {}; REGIONS.forEach(r => posDaily[r] = 0);
tsaRef.forEach(r => { const k = (r.REGION || "").trim().toUpperCase(); if (posDaily[k] !== undefined) posDaily[k] += num(r.OBJECTIF_DAILY); });
const posDailyTotal = REGIONS.reduce((s, r) => s + posDaily[r], 0);

const classement = readCSV(path.join(IN, "classement.csv"));
const weekNum = w => parseInt(String(w).replace(/\D/g, "")) || 0;
// période POS d'une semaine N (ven→jeu, ancrée PROGRAMME_START) — même formule que l'hebdo
function posWeekWindow(N) {
  const ps = new Date(PROGRAMME_START.getTime() + (N - 1) * 7 * MS);
  return [ps, new Date(ps.getTime() + 6 * MS)];
}
const allWeeks = [...new Set(classement.map(r => weekNum(r.SEMAINE)))].filter(Boolean).sort((a, b) => a - b);

const form = readCSV(path.join(IN, "form.csv"));
const H0 = Object.keys(form[0] || {});
const col = re => H0.find(h => re.test(h)) || "";
const cReg = col(/Region/i), cMont = col(/Montant/i), cTS = col(/Timestamp/i);
const cNumBA = col(/num.*ro.*du.*ba|numero_ba/i); // identifiant stable (téléphone), pas l'email
// pré-parse une seule fois : le form live fait ~50 k lignes et on le rebalaye par mois
const formRows = form.map(r => ({
  reg: (r[cReg] || "").trim().toUpperCase(), d: parseTS(r[cTS]),
  mont: num(r[cMont]), ba: (r[cNumBA] || "").trim().toUpperCase(),
})).filter(r => r.d && REGIONS.includes(r.reg));

const effTotal = Object.values(CONFIG.effectif).reduce((a, b) => a + b, 0);
const baDailyByReg = {}; REGIONS.forEach(r => { baDailyByReg[r] = Math.round((CONFIG.effectif[r] / effTotal) * CONFIG.baDailyGlobal); });
const dailyTarget = CONFIG.baDailyGlobal;

// — Agrégats d'un mois (POS + BA), réutilisés pour le mois rendu ET pour la série mois/mois —
function computeMonth(win) {
  // POS : semaines du classement dont la période chevauche la fenêtre POS
  const includedWeeks = allWeeks.filter(N => { const [ps, pe] = posWeekWindow(N); return ps <= win.end && pe >= win.posStart; });
  const volByReg = {}; REGIONS.forEach(r => volByReg[r] = 0);
  const volByTSA = new Map();
  classement.forEach(r => {
    if (!includedWeeks.includes(weekNum(r.SEMAINE))) return;
    const reg = (r.REGION || "").trim().toUpperCase(); if (volByReg[reg] === undefined) return;
    const v = num(r.VOLUME_XAF);
    volByReg[reg] += v;
    const key = String(r.TSA || r.TSA_FULL_NAME || "").trim();
    const cur = volByTSA.get(key) || { TSA: r.TSA_FULL_NAME || r.TSA, REGION: reg, RBM: r.RBM, VOLUME_XAF: 0 };
    cur.VOLUME_XAF += v; volByTSA.set(key, cur);
  });
  const volTotal = REGIONS.reduce((s, r) => s + volByReg[r], 0);
  // n°1 de chaque région sur le CUMUL du mois (le rang hebdo du classement ne s'applique pas ici)
  const primeTSA = REGIONS.map(reg =>
    [...volByTSA.values()].filter(t => t.REGION === reg).sort((a, b) => b.VOLUME_XAF - a.VOLUME_XAF)[0]
  ).filter(Boolean).sort((a, b) => b.VOLUME_XAF - a.VOLUME_XAF);
  const posObjectif = Math.round(posDailyTotal) * win.posDays;

  // BA : activations + montant par région, série jour-par-jour, diagnostic transactions/actifs
  const baAgg = {}; REGIONS.forEach(r => baAgg[r] = { act: 0, mont: 0 });
  const dayMap = {};
  const actifs = new Set();
  let nbTx = 0, baMont = 0;
  formRows.forEach(r => {
    if (r.d < win.baStart || r.d > win.endTS) return;
    baAgg[r.reg].act++; baAgg[r.reg].mont += r.mont;
    nbTx++; baMont += r.mont; if (r.ba) actifs.add(r.ba);
    const key = `${r.d.getFullYear()}-${String(r.d.getMonth() + 1).padStart(2, "0")}-${String(r.d.getDate()).padStart(2, "0")}`;
    dayMap[key] = (dayMap[key] || 0) + r.mont;
  });
  const baObjectif = dailyTarget * win.workDays;

  return { win, includedWeeks, volByReg, volTotal, primeTSA, posObjectif, baAgg, dayMap, baObjectif, nbTx, baMont, actifs: actifs.size };
}

const D = computeMonth(WIN);

// — Série MOIS PAR MOIS : du 1er mois du programme jusqu'au mois rendu —
const monthsSeries = [];
for (let y = BA_START.getFullYear(), m = BA_START.getMonth() + 1; y < YY || (y === YY && m <= MM); m === 12 ? (m = 1, y++) : m++) {
  const w = monthWindow(y, m);
  monthsSeries.push(y === YY && m === MM ? D : computeMonth(w));
}

// — rythme HEBDOMADAIRE (une barre par semaine BA du mois, clippée à la fenêtre) —
// chaque semaine BA (lun→dim, ancrée BA_START) qui chevauche la fenêtre ; objectif = daily × jours
// ouvrés (hors dimanche) réellement présents dans la portion visible de la semaine.
function baWeeksInWindow(win) {
  const weeks = [];
  for (let N = 1; N <= 60; N++) {
    const ws = new Date(BA_START.getTime() + (N - 1) * 7 * MS);
    if (ws > win.end) break;
    const we = new Date(ws.getTime() + 6 * MS);
    const cs = ws < win.baStart ? win.baStart : ws, ce = we > win.end ? win.end : we;
    if (ce < cs) continue;
    let wd = 0; for (let t = cs.getTime(); t <= ce.getTime(); t += MS) if (new Date(t).getDay() !== 0) wd++;
    weeks.push({ N, start: cs, end: new Date(ce.getTime() + MS - 1), workDays: wd });
  }
  return weeks;
}
const weekSeries = baWeeksInWindow(WIN).map(w => {
  let mont = 0;
  Object.keys(D.dayMap).forEach(k => {
    const [y, m, d] = k.split("-").map(Number); const dt = new Date(y, m - 1, d);
    if (dt >= w.start && dt <= w.end) mont += D.dayMap[k];
  });
  const target = dailyTarget * w.workDays;
  return { label: "Semaine " + w.N, mont, pct: target ? mont / target * 100 : 0 };
});

// ════════════════════════════════════════════════════════════════════
//  PPTX SETUP
// ════════════════════════════════════════════════════════════════════
const EnhancedPPTXGenJS = withPPTXEmbedFonts(pptxgen);
const pres = new EnhancedPPTXGenJS();
pres.layout = "LAYOUT_WIDE";
pres.title = "Progress Report DTC Assisted — Vue mensuelle — MTN Bénin";
pres.author = "Contribution LKA";
pres.company = "MTN Bénin";
const TOTAL = 8;
const tauxColor = p => p >= 90 ? GREEN : p >= 50 ? AMBER : REDX;

// ════════════════════════════════════════════════════════════════════
//  SLIDE 1 · COUVERTURE
// ════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide(); light(s);
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 0.09, fill: { color: YEL }, line: { type: "none" } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 4.95, w: 4.5, h: 2.55, fill: { color: YEL }, line: { type: "none" } });
  s.addShape(pres.shapes.RECTANGLE, { x: 4.5, y: 6.55, w: 0.95, h: 0.95, fill: { color: ROW }, line: { type: "none" } });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.45, y: 6.55, w: 0.95, h: 0.95, fill: { color: YELSOFT }, line: { type: "none" } });

  if (HAS_LOGO) s.addImage({ path: LOGO, x: MX, y: 0.55, w: 1.0, h: 0.93 });

  s.addText("MTN BÉNIN  ·  REPORTING TERRAIN", {
    x: MX, y: 1.95, w: 9, h: 0.3, margin: 0, fontFace: fBody, fontSize: 13, color: YELDK, charSpacing: 3,
  });
  s.addText("RAPPORT", { x: MX - 0.04, y: 2.3, w: 11, h: 1.1, margin: 0, fontFace: fHead, bold: true, fontSize: 70, color: INK });
  s.addText("MENSUEL", { x: MX - 0.04, y: 3.32, w: 12, h: 1.1, margin: 0, fontFace: fHead, bold: true, fontSize: 70, color: INK });
  s.addText("DTC ASSISTED", { x: MX, y: 4.42, w: 11, h: 0.55, margin: 0, fontFace: fHead, bold: true, fontSize: 26, color: YELDK, charSpacing: 1 });

  s.addText("Contribution LKA", { x: 0.5, y: 5.55, w: 3.8, h: 0.45, margin: 0, fontFace: fHead, bold: true, fontSize: 22, color: "000000" });
  s.addText(`${MONTH_LABEL} · ${WINDOW_LABEL}`, { x: 0.5, y: 6.08, w: 3.8, h: 0.7, margin: 0, fontFace: fBody, fontSize: 12.5, color: "1A1A1A", lineSpacingMultiple: 1.05 });

  s.addText("OBJECTIF", { x: 9.6, y: 5.0, w: 3.13, h: 0.3, margin: 0, fontFace: fBody, fontSize: 11, color: MUTE, charSpacing: 3, align: "right" });
  s.addText("60%", { x: 9.6, y: 5.22, w: 3.13, h: 1.0, margin: 0, fontFace: fHead, bold: true, fontSize: 64, color: YELDK, align: "right" });
  s.addText("DTC EOY 2026", { x: 9.6, y: 6.32, w: 3.13, h: 0.3, margin: 0, fontFace: fBody, fontSize: 12, color: INK, charSpacing: 2, align: "right" });

  s.addText([
    { text: MONTH_LABEL + "\n", options: { fontFace: fHead, bold: true, fontSize: 14, color: INK, breakLine: true } },
    { text: "Édité le " + REPORT_DATE, options: { fontFace: fBody, fontSize: 11, color: MUTE } },
  ], { x: 8.73, y: 0.62, w: 4, h: 0.8, margin: 0, align: "right", valign: "top", lineSpacingMultiple: 1.1 });
}

// ════════════════════════════════════════════════════════════════════
//  SLIDE 2 · ÉTAT DU RECRUTEMENT
// ════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide(); light(s);
  chrome(pres, s, 2, "État d'avancement · Recrutement des BA");
  header(pres, s, "État du recrutement", "Renforcement de la force de vente terrain pour l'opération DTC Assisted");

  s.addText(`Après avoir défini la stratégie et les objectifs, voici l'état d'avancement du recrutement des Brand Ambassadors (BA), qui constituent la force de soutien des TSA, à la date du ${REPORT_DATE}.`, {
    x: MX, y: 1.5, w: CW, h: 0.5, margin: 0, fontFace: fBody, fontSize: 12, color: INK2, lineSpacingMultiple: 1.1,
  });

  // recrutement à 100 % partout : recrutés = effectif (donnée figée, comme l'hebdo)
  const rows = []; let tAR = 0, tR = 0;
  REGIONS.forEach(reg => {
    const ar = CONFIG.effectif[reg] || 0, rec = ar;
    tAR += ar; tR += rec;
    const p = ar ? Math.round(rec / ar * 100) : 0;
    rows.push([reg, fmt(ar), fmt(rec), { text: p + "%", color: tauxColor(p), bold: true }]);
  });
  const gp = tAR ? Math.round(tR / tAR * 100) : 0;
  const tx = MX, tw = 7.0, ty = 2.2;
  table(pres, s, tx, ty, tw, ["RÉGION", "À RECRUTER", "RECRUTÉS", "TAUX"], rows, {
    colWidths: [tw * 0.34, tw * 0.24, tw * 0.22, tw * 0.20], rowH: 0.40,
    totalRow: ["TOTAL", fmt(tAR), fmt(tR), gp + "%"],
  });

  const rx = MX + tw + 0.45, rw = W - MX - rx;
  s.addShape(pres.shapes.RECTANGLE, { x: rx, y: ty, w: rw, h: 1.9, fill: { color: YEL }, line: { type: "none" } });
  s.addText("TAUX GLOBAL DE RECRUTEMENT", { x: rx + 0.25, y: ty + 0.22, w: rw - 0.5, h: 0.3, margin: 0, fontFace: fHead, bold: true, fontSize: 11, color: "000000", charSpacing: 1 });
  s.addText(gp + "%", { x: rx + 0.18, y: ty + 0.5, w: rw - 0.36, h: 1.25, margin: 0, fontFace: fHead, bold: true, fontSize: 76, color: "000000" });

  const insights = [
    { c: GREEN, t: `${REGIONS.length} régions sur ${REGIONS.length} à 100 % de leur objectif.` },
    { c: YEL, t: `${gp} % global · ${fmt(tR)} BA recrutés sur ${fmt(tAR)} prévus.` },
    { c: YELDK, t: `${fmt(D.actifs)} BA distincts ont réalisé au moins une activation en ${MOIS[MM - 1]}.` },
  ];
  let iy = ty + 2.15;
  insights.forEach(it => {
    s.addShape(pres.shapes.RECTANGLE, { x: rx, y: iy, w: 0.1, h: 0.6, fill: { color: it.c }, line: { type: "none" } });
    s.addText(it.t, { x: rx + 0.25, y: iy, w: rw - 0.3, h: 0.6, margin: 0, fontFace: fBody, fontSize: 11, color: INK, valign: "middle", lineSpacingMultiple: 1.0 });
    iy += 0.72;
  });
}

// ════════════════════════════════════════════════════════════════════
//  SLIDE 3 · ACTIVATIONS POS — VUE MENSUELLE
// ════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide(); light(s);
  chrome(pres, s, 3, "Vue mensuelle · Performance base globale TSA");
  header(pres, s, "Activations POS — Performance par région", `Volume des points de vente (base globale TSA) · ${MONTH_LABEL} · ${POS_WINDOW_LABEL}`);

  s.addText(`Performances globales des points de vente sur ${POS_WINDOW_LABEL} (${WIN.posDays} jours) — réalisation rapportée aux cibles mensuelles par région.`, {
    x: MX, y: 1.5, w: CW, h: 0.4, margin: 0, fontFace: fBody, fontSize: 12, color: INK2,
  });

  const rows = []; let tPos = 0, tD = 0, tM = 0, tReal = 0;
  REGIONS.forEach(reg => {
    const daily = Math.round(posDaily[reg]);
    const monthlyTarget = daily * WIN.posDays;
    const real = D.volByReg[reg] || 0;
    const uniquePos = (CONFIG.pos[reg] || {}).uniquePos || 0;
    const taux = monthlyTarget ? real / monthlyTarget * 100 : 0;
    tPos += uniquePos; tD += daily; tM += monthlyTarget; tReal += real;
    rows.push([reg, fmt(uniquePos), fmt(daily), fmt(monthlyTarget), fmt(real), { text: taux.toFixed(1) + "%", color: taux >= 100 ? GREEN : taux >= 50 ? AMBER : REDX, bold: true }]);
  });
  const tt = tM ? tReal / tM * 100 : 0;
  const ty = 2.0, tw = CW;
  table(pres, s, MX, ty, tw, ["RÉGION", "UNIQUE POS", "TARGET DAILY", "TARGET MOIS", "RÉALISATION", "TAUX"], rows, {
    colWidths: [tw * 0.18, tw * 0.13, tw * 0.18, tw * 0.18, tw * 0.21, tw * 0.12],
    rowH: 0.46, fs: 10.5, hfs: 9.5,
    totalRow: ["TOTAL", fmt(tPos), fmt(tD), fmt(tM), fmt(tReal), tt.toFixed(1) + "%"],
  });

  s.addText(`Cible mensuelle totale : ${fmt(tM)} FCFA · Réalisé : ${fmt(tReal)} FCFA · Taux de réalisation global : ${tt.toFixed(1)} %.`, {
    x: MX, y: 6.42, w: CW, h: 0.3, margin: 0, fontFace: fBody, italic: true, fontSize: 10.5, color: MUTE,
  });
  s.addText(`Réalisé = cumul des ${D.includedWeeks.map(n => "S" + n).join(", ")} du classement (semaines POS ven→jeu chevauchant la fenêtre ; une semaine à cheval compte dans les deux mois).`, {
    x: MX, y: 6.72, w: CW, h: 0.3, margin: 0, fontFace: fBody, italic: true, fontSize: 9, color: MUTE2,
  });
}

// ════════════════════════════════════════════════════════════════════
//  SLIDE 4 · PRIME & RANG — CLASSEMENT TSA DU MOIS
// ════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide(); light(s);
  chrome(pres, s, 4, "Reconnaissance · Classement TSA mensuel");
  header(pres, s, "Prime & Rang", `Classement mensuel — le n°1 de chaque région · ${MONTH_LABEL}`);

  const champion = D.primeTSA[0] || {};
  const taux = D.posObjectif ? D.volTotal / D.posObjectif * 100 : 0;
  const kpis = [
    { l: "OBJECTIF DU MOIS", v: fmt(D.posObjectif), c: INK },
    { l: "VOLUME TOTAL", v: fmt(D.volTotal), c: INK },
    { l: "TAUX DE RÉALISATION", v: taux.toFixed(2) + "%", c: YELDK },
    { l: "TSA DU MOIS", v: champion.TSA ? toTitle(champion.TSA) : "—", c: GREEN, small: true },
  ];
  const ky = 1.55, kwid = (CW - 0.3 * 3) / 4, kh = 1.05;
  kpis.forEach((k, i) => {
    const x = MX + i * (kwid + 0.3);
    s.addShape(pres.shapes.RECTANGLE, { x, y: ky, w: kwid, h: kh, fill: { color: ROW }, line: { type: "none" } });
    s.addShape(pres.shapes.RECTANGLE, { x, y: ky, w: 0.08, h: kh, fill: { color: YEL }, line: { type: "none" } });
    s.addText(k.l, { x: x + 0.2, y: ky + 0.13, w: kwid - 0.36, h: 0.25, margin: 0, fontFace: fBody, fontSize: 8.5, color: MUTE, charSpacing: 1.5 });
    s.addText(k.v, { x: x + 0.18, y: ky + 0.36, w: kwid - 0.32, h: 0.6, margin: 0, fontFace: fHead, bold: true, fontSize: k.small ? 14 : 26, color: k.c, valign: "middle" });
  });

  // Seuil de prime mensualisé : seuil hebdo × nb de semaines POS retenues pour le mois
  const nbWeeks = D.includedWeeks.length || 1;
  const threshold = CONFIG.primeThreshold * nbWeeks;
  let nbElig = 0;
  const rows = D.primeTSA.map(r => {
    const elig = r.VOLUME_XAF >= threshold; if (elig) nbElig++;
    return [r.REGION, toTitle(r.TSA), r.RBM ? toTitle(r.RBM) : "—", fmt(r.VOLUME_XAF),
      { text: elig ? "OUI" : "NON", color: elig ? GREEN : REDX, bold: true }];
  });
  const elgOK = nbElig > 0, bnY = 2.62;
  s.addShape(pres.shapes.RECTANGLE, { x: MX, y: bnY, w: CW, h: 0.4, fill: { color: elgOK ? "E9F6EC" : YELSOFT }, line: { type: "none" } });
  s.addShape(pres.shapes.RECTANGLE, { x: MX, y: bnY, w: 0.08, h: 0.4, fill: { color: elgOK ? GREEN : YELDK }, line: { type: "none" } });
  s.addText([
    { text: "Prime TSA (mois) : ", options: { fontFace: fHead, bold: true, color: INK } },
    { text: `n°1 régional sur le cumul du mois ET volume ≥ ${fmt(threshold)} FCFA (${fmt(CONFIG.primeThreshold)} × ${nbWeeks} sem.) — `, options: { fontFace: fBody, color: INK2 } },
    { text: elgOK ? `${nbElig} lauréat${nbElig > 1 ? "s" : ""} ce mois-ci.` : `aucun éligible ce mois-ci (meilleur volume : ${fmt(champion.VOLUME_XAF || 0)} FCFA).`, options: { fontFace: fHead, bold: true, color: elgOK ? GREEN : REDX } },
  ], { x: MX + 0.22, y: bnY, w: CW - 0.4, h: 0.4, margin: 0, valign: "middle", fontSize: 10.5 });

  table(pres, s, MX, 3.18, CW, ["RÉGION", "TSA (N°1 RÉGIONAL)", "RBM", "VOLUME (XAF)", "PRIME"], rows, {
    colWidths: [CW * 0.16, CW * 0.30, CW * 0.28, CW * 0.16, CW * 0.10],
    rowH: 0.42, fs: 10, hfs: 9.5,
  });
}

// ════════════════════════════════════════════════════════════════════
//  SLIDE 5 · ACTIVATIONS BA — VUE MENSUELLE
// ════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide(); light(s);
  chrome(pres, s, 5, "Vue mensuelle · Activations BA");
  header(pres, s, "Activations BA — Force de soutien aux TSA", `Programme BA en appui des TSA · ${MONTH_LABEL} · ${WINDOW_LABEL}`);

  const rows = []; let tEff = 0, tAct = 0, tTar = 0, tMon = 0;
  REGIONS.forEach(reg => {
    const eff = CONFIG.effectif[reg] || 0;
    const o = D.baAgg[reg]; const act = o.act, mont = o.mont;
    const target = baDailyByReg[reg] * WIN.workDays;
    const p = target ? mont / target * 100 : 0;
    tEff += eff; tAct += act; tTar += target; tMon += mont;
    rows.push([reg, fmt(eff), fmt(act), fmt(target), fmt(mont), { text: p.toFixed(1) + "%", color: p >= 100 ? GREEN : p >= 50 ? AMBER : REDX, bold: true }]);
  });
  const tp = tTar ? tMon / tTar * 100 : 0;
  const tx = MX, tw = 7.35, ty = 2.1;
  table(pres, s, tx, ty, tw, ["RÉGION", "EFFECTIF", "ACTIVATIONS", "TARGET", "MONTANT", "TAUX"], rows, {
    colWidths: [tw * 0.22, tw * 0.15, tw * 0.20, tw * 0.17, tw * 0.16, tw * 0.10],
    rowH: 0.40, fs: 9.5, hfs: 8.5,
    totalRow: ["TOTAL", fmt(tEff), fmt(tAct), fmt(tTar), fmt(tMon), tp.toFixed(1) + "%"],
  });

  // colonne droite : callouts + rythme HEBDOMADAIRE (une barre par semaine BA du mois)
  const rx = MX + tw + 0.4, rw = W - MX - rx;
  const cwid = (rw - 0.2) / 2;
  s.addShape(pres.shapes.RECTANGLE, { x: rx, y: ty, w: cwid, h: 1.05, fill: { color: YEL }, line: { type: "none" } });
  s.addText("ACTIVATIONS", { x: rx + 0.16, y: ty + 0.13, w: cwid - 0.3, h: 0.24, margin: 0, fontFace: fHead, bold: true, fontSize: 9.5, color: "000000", charSpacing: 1.5 });
  s.addText(fmt(tAct), { x: rx + 0.14, y: ty + 0.36, w: cwid - 0.28, h: 0.62, margin: 0, fontFace: fHead, bold: true, fontSize: 34, color: "000000", valign: "middle" });
  const mx2 = rx + cwid + 0.2;
  s.addShape(pres.shapes.RECTANGLE, { x: mx2, y: ty, w: cwid, h: 1.05, fill: { color: BG }, line: { color: LINE, width: 1 } });
  s.addText("MONTANT (FCFA)", { x: mx2 + 0.16, y: ty + 0.13, w: cwid - 0.3, h: 0.24, margin: 0, fontFace: fBody, fontSize: 9, color: MUTE, charSpacing: 1.2 });
  // le cumul mensuel a 2 chiffres de plus que l'hebdo → réduire pour tenir sur une ligne
  const monTxt = fmt(tMon);
  s.addText(monTxt, { x: mx2 + 0.14, y: ty + 0.36, w: cwid - 0.28, h: 0.62, margin: 0, fontFace: fHead, bold: true, fontSize: monTxt.length > 9 ? 17 : 24, color: INK, valign: "middle" });

  const py = ty + 1.3;
  s.addText("RYTHME HEBDOMADAIRE", { x: rx, y: py, w: rw - 1.4, h: 0.26, margin: 0, fontFace: fHead, bold: true, fontSize: 11, color: INK, charSpacing: 1 });
  s.addText("% vs objectif/sem", { x: rx + rw - 1.6, y: py + 0.02, w: 1.6, h: 0.24, margin: 0, fontFace: fBody, fontSize: 8.5, color: MUTE, align: "right" });
  const maxPct = Math.max(100, ...weekSeries.map(d => d.pct));
  const scaleMax = Math.ceil(maxPct / 20) * 20;
  const lblW = 0.95, valW = 0.62, gap = 0.1;
  const trackX = rx + lblW + gap, trackW = rw - lblW - valW - gap * 2;
  const availH = 4.55 - (py + 0.4 - ty);
  const rowH2 = Math.max(0.32, Math.min(0.55, availH / Math.max(1, weekSeries.length)));
  const by = py + 0.4;
  weekSeries.forEach((d, i) => {
    const y = by + i * rowH2;
    const pr = Math.round(d.pct);
    const c = pr >= 90 ? GREEN : pr >= 50 ? AMBER : REDX;
    s.addText(d.label, { x: rx, y, w: lblW, h: rowH2 - 0.05, margin: 0, fontFace: fBody, fontSize: 9, color: MUTE, valign: "middle" });
    s.addShape(pres.shapes.RECTANGLE, { x: trackX, y: y + rowH2 * 0.15, w: trackW, h: rowH2 * 0.7, fill: { color: ROW }, line: { type: "none" } });
    const fillW = Math.max(0.02, Math.min(1, d.pct / scaleMax) * trackW);
    s.addShape(pres.shapes.RECTANGLE, { x: trackX, y: y + rowH2 * 0.15, w: fillW, h: rowH2 * 0.7, fill: { color: c }, line: { type: "none" } });
    s.addText(pr + "%", { x: rx + rw - valW, y, w: valW, h: rowH2 - 0.05, margin: 0, fontFace: fHead, bold: true, fontSize: 9.5, color: c, align: "right", valign: "middle" });
  });
  const objX = trackX + (100 / scaleMax) * trackW;
  const barsBottom = by + weekSeries.length * rowH2;
  s.addShape(pres.shapes.RECTANGLE, { x: objX, y: by - 0.02, w: 0.014, h: Math.max(0, barsBottom - by - 0.05), fill: { color: INK }, line: { type: "none" } });
  s.addText("objectif", { x: objX - 0.6, y: barsBottom, w: 1.2, h: 0.2, margin: 0, fontFace: fBody, fontSize: 7.5, color: MUTE, align: "center" });
}

// ════════════════════════════════════════════════════════════════════
//  SLIDE 6 · MONTH-OVER-MONTH — RÉALISÉ vs OBJECTIF, PAR MOIS
// ════════════════════════════════════════════════════════════════════
// Pendant mensuel de la slide week-over-week de l'hebdo, mais dessiné nativement
// plutôt que repris de gen_wow_chart.py : ce graphique-là superpose une échelle de
// variance % et une échelle absolue centrées sur 0, ce qui fait coïncider la barre
// de variance avec la ligne d'objectif dès que l'objectif domine l'échelle — illisible
// avec 2 ou 3 mois. Ici chaque mois a SON objectif (nb de jours différent, juin tronqué),
// donc on compare des TAUX de réalisation, seule grandeur comparable entre mois.
{
  const s = pres.addSlide(); light(s);
  chrome(pres, s, 6, "Tendances · Évolution mois/mois");
  header(pres, s, "Month-over-month — POS & BA", `Réalisé vs objectif du mois · jusqu'à ${MONTH_LABEL}`);

  const M = n => (n / 1e6).toFixed(n >= 1e7 ? 0 : 1).replace(".", ",") + " M";
  const panels = [
    { title: "POS — Volume réalisé", real: m => m.volTotal, obj: m => m.posObjectif, unit: m => `${m.win.posDays} j` },
    { title: "BA — Montant réalisé", real: m => m.baMont, obj: m => m.baObjectif, unit: m => `${m.win.workDays} j ouvrés` },
  ];
  const pw = (CW - 0.6) / 2, py0 = 1.7;
  panels.forEach((P, pi) => {
    const px = MX + pi * (pw + 0.6);
    s.addText(P.title, { x: px, y: py0, w: pw, h: 0.3, margin: 0, fontFace: fHead, bold: true, fontSize: 13, color: INK });
    s.addShape(pres.shapes.RECTANGLE, { x: px, y: py0 + 0.34, w: pw, h: 0.014, fill: { color: LINE }, line: { type: "none" } });

    const rowH = Math.min(1.35, 4.3 / Math.max(1, monthsSeries.length));
    monthsSeries.forEach((m, i) => {
      const real = P.real(m), obj = P.obj(m), pct = obj ? real / obj * 100 : 0;
      const c = tauxColor(pct);
      const y = py0 + 0.55 + i * rowH;
      // libellé du mois + fenêtre
      s.addText(m.win.short + (m.win.partial ? " *" : ""), { x: px, y, w: pw * 0.42, h: 0.26, margin: 0, fontFace: fHead, bold: true, fontSize: 11.5, color: INK });
      s.addText(`objectif ${M(obj)} · ${P.unit(m)}`, { x: px + pw * 0.42, y: y + 0.02, w: pw * 0.58, h: 0.24, margin: 0, fontFace: fBody, fontSize: 8.5, color: MUTE2, align: "right" });
      // piste = objectif du mois ; remplissage = réalisé (échelle propre au mois → taux comparables)
      const barY = y + 0.32, barH = 0.34;
      s.addShape(pres.shapes.RECTANGLE, { x: px, y: barY, w: pw, h: barH, fill: { color: ROW }, line: { type: "none" } });
      s.addShape(pres.shapes.RECTANGLE, { x: px, y: barY, w: Math.max(0.03, Math.min(1, pct / 100) * pw), h: barH, fill: { color: c }, line: { type: "none" } });
      s.addText(`${M(real)}   ${pct.toFixed(1)} %`, { x: px + 0.1, y: barY, w: pw - 0.2, h: barH, margin: 0, fontFace: fHead, bold: true, fontSize: 10, color: pct >= 45 ? "FFFFFF" : INK, valign: "middle" });
      // variation m/m (le taux, pas le volume brut : seul comparable quand les mois n'ont pas la même durée)
      if (i > 0) {
        const pp = P.obj(monthsSeries[i - 1]) ? P.real(monthsSeries[i - 1]) / P.obj(monthsSeries[i - 1]) * 100 : 0;
        const d = pct - pp, up = d >= 0;
        s.addText(`${up ? "▲" : "▼"} ${up ? "+" : ""}${d.toFixed(1)} pt vs ${monthsSeries[i - 1].win.short.toLowerCase()}`, {
          x: px, y: barY + barH + 0.04, w: pw, h: 0.24, margin: 0, fontFace: fHead, bold: true, fontSize: 9, color: up ? GREEN : REDX,
        });
      }
    });
  });

  const partial = monthsSeries.filter(m => m.win.partial).map(m => m.win.short);
  s.addText(`Chaque mois est rapporté à SON propre objectif (objectif journalier × nombre de jours du mois) : les mois n'ayant pas la même durée, seuls les taux de réalisation sont comparables — pas les volumes bruts.${partial.length ? ` * ${partial.join(", ")} : fenêtre partielle, le programme a démarré le ${dmy(BA_START)}.` : ""}`, {
    x: MX, y: 6.4, w: CW, h: 0.6, margin: 0, fontFace: fBody, italic: true, fontSize: 9, color: MUTE2, lineSpacingMultiple: 1.05,
  });
}

// ════════════════════════════════════════════════════════════════════
//  SLIDE 7 · ANALYSE BA — TRANSACTIONS, VALEUR & PERSONNES ACTIVES
// ════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide(); light(s);
  chrome(pres, s, 7, "Analyse · Activité BA mois par mois");
  header(pres, s, "Analyse BA — Transactions & activité", `Nombre de transactions, valeur et personnes actives par mois · jusqu'à ${MONTH_LABEL}`);

  const upliftCell = (curr, prev) => {
    if (prev == null || !prev) return { text: fmt(curr), color: INK };
    const d = (curr - prev) / prev * 100;
    const up = d >= 0;
    return { text: `${fmt(curr)}  ${up ? "▲" : "▼"} ${up ? "+" : ""}${d.toFixed(1)}%`, color: up ? GREEN : REDX, bold: true };
  };
  const rows = monthsSeries.map((m, i) => {
    const p = i > 0 ? monthsSeries[i - 1] : null;
    return [
      m.win.short + (m.win.partial ? " *" : ""),
      upliftCell(m.nbTx, p && p.nbTx),
      upliftCell(m.baMont, p && p.baMont),
      upliftCell(m.actifs, p && p.actifs),
    ];
  });
  const ty = 2.0;
  table(pres, s, MX, ty, CW, ["MOIS", "NB TRANSACTIONS", "MONTANT (FCFA)", "#PERSONNES ACTIVES"], rows, {
    colWidths: [CW * 0.14, CW * 0.30, CW * 0.30, CW * 0.26], rowH: 0.5, fs: 11, hfs: 10,
  });

  s.addText([
    { text: "Personnes actives = numéros de téléphone BA distincts ayant réalisé au moins une activation dans le mois (identifiant stable, indépendant de l'email).", options: { breakLine: true } },
    { text: monthsSeries.some(m => m.win.partial) ? `* fenêtre partielle : le programme a démarré le ${dmy(BA_START)}.` : "" },
  ], {
    x: MX, y: ty + monthsSeries.length * 0.5 + 0.6, w: CW, h: 0.6, margin: 0,
    fontFace: fBody, italic: true, fontSize: 10, color: MUTE, lineSpacingMultiple: 1.1,
  });
}

// ════════════════════════════════════════════════════════════════════
//  SLIDE 8 · CLÔTURE
// ════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide(); light(s);
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 0.09, fill: { color: YEL }, line: { type: "none" } });
  s.addShape(pres.shapes.RECTANGLE, { x: 9.55, y: 0, w: 3.78, h: H, fill: { color: YEL }, line: { type: "none" } });
  s.addShape(pres.shapes.RECTANGLE, { x: 9.55, y: 0, w: 0.95, h: 0.95, fill: { color: ROW }, line: { type: "none" } });

  if (HAS_LOGO) s.addImage({ path: LOGO, x: MX, y: 0.85, w: 1.0, h: 0.93 });

  s.addText("MERCI", { x: MX - 0.06, y: 2.5, w: 9, h: 1.5, margin: 0, fontFace: fHead, bold: true, fontSize: 92, color: INK });
  s.addText("Ensemble vers 60 % DTC EOY 2026", { x: MX, y: 4.05, w: 8.5, h: 0.4, margin: 0, fontFace: fBody, fontSize: 16, color: YELDK, charSpacing: 1 });

  s.addShape(pres.shapes.RECTANGLE, { x: MX, y: 5.05, w: 0.24, h: 1.0, fill: { color: YEL }, line: { type: "none" } });
  s.addText([
    { text: "Prochaines étapes\n", options: { fontFace: fHead, bold: true, fontSize: 13, color: INK, breakLine: true } },
    { text: "Maintenir le rythme jour-par-jour · Soutenir le week-end · Convertir le Top des POS", options: { fontFace: fBody, fontSize: 11.5, color: MUTE } },
  ], { x: MX + 0.42, y: 5.05, w: 8.5, h: 1.1, margin: 0, valign: "top", lineSpacingMultiple: 1.2 });

  s.addText("RAPPORT MENSUEL · DTC ASSISTED", { x: 10.0, y: 2.3, w: 5.0, h: 0.4, margin: 0, fontFace: fHead, bold: true, fontSize: 13, color: "000000", rotate: 90, align: "center" });
  s.addText(`${MONTH_LABEL} · ${WINDOW_LABEL}`, { x: MX, y: 6.55, w: 8.5, h: 0.3, margin: 0, fontFace: fBody, fontSize: 10, color: MUTE2, charSpacing: 1 });
}

// ── EMBED FONTS & WRITE ─────────────────────────────────────────────
(async () => {
  await embedFonts(pres);
  const outDir = process.env.OUT_DIR ? path.resolve(process.env.OUT_DIR) : path.join(__dirname, "..", "outputs");
  fs.mkdirSync(outDir, { recursive: true });
  const tag = YM.replace("-", "");
  let out = path.join(outDir, `Progress_Report_DTC_Assisted_LKA_Mensuel_${tag}.pptx`);
  try {
    await pres.writeFile({ fileName: out });
  } catch (e) {
    if (e.code !== "EBUSY" && e.code !== "EPERM") throw e;
    out = path.join(outDir, `Progress_Report_DTC_Assisted_LKA_Mensuel_${tag}_${Date.now().toString().slice(-4)}.pptx`);
    await pres.writeFile({ fileName: out });
    console.warn("⚠️ Fichier principal verrouillé (ouvert ?) — écrit sous un nom alternatif.");
  }
  console.log(`✅ Done · ${TOTAL} slides · mois=${MONTH_LABEL} · ${out}`);
})();
