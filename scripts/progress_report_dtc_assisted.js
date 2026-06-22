// ════════════════════════════════════════════════════════════════════
//  PROGRESS REPORT · DTC ASSISTED · MTN BÉNIN · Contribution LKA
//  Rapport hebdomadaire — généré depuis les onglets Google Sheets de la chaîne DTC.
//  Lancement : node scripts/progress_report_dtc_assisted.js [semaine]
//   (sans argument → dernière SEMAINE présente dans classement.csv)
//
//  Données dans INPUT_DIR (défaut ./inputs ; en CI = onglets Sheets téléchargés) :
//   • classement.csv  (onglet CLASSEMENT_TSA)  → Prime & Rang + réalisation POS
//   • tsa_ref.csv      (onglet TSA_REF, OBJECTIF_DAILY) → cibles POS par région (12M/jour)
//   • form.csv         (réponses du Form BA)    → Activations BA par région
// ════════════════════════════════════════════════════════════════════
const pptxgen = require("pptxgenjs");
const { withPPTXEmbedFonts } = require("pptx-embed-fonts/pptxgenjs");
const fs = require("fs");
const path = require("path");

// Dossier des données. En CI, le workflow télécharge les onglets Sheets ici.
const IN = process.env.INPUT_DIR ? path.resolve(process.env.INPUT_DIR) : path.join(__dirname, "..", "inputs");

// ╔══════════════════════════════════════════════════════════════════╗
// ║  ⚙️  CONFIGURATION                                                 ║
// ║  Le rapport est un CYCLE par semaine. On choisit la semaine en    ║
// ║  paramètre :   node "Scripts/progress_report_dtc_assisted.js" 1   ║
// ╚══════════════════════════════════════════════════════════════════╝

// — Un cycle par semaine. Ajouter une entrée chaque semaine. —
//   ba  : plage de dates des activations BA (filtre Timestamp du formulaire) [début, fin] AAAA-MM-JJ
//   pos : libellé de la période POS (le volume vient de la tranche "Semaine N" du classement)
const CYCLES = {
  "Semaine 1": {
    label:      "Semaine 1 · 15 au 21 juin 2026",
    reportDate: "22/06/2026",
    ba:  ["2026-06-15", "2026-06-21"],
    pos: "12 au 18 juin 2026",
  },
  // "Semaine 2": {
  //   label: "Semaine 2 · 22 au 28 juin 2026", reportDate: "29/06/2026",
  //   ba: ["2026-06-22", "2026-06-28"], pos: "19 au 25 juin 2026",
  // },
};
const DEFAULT_WEEK = "Semaine 1";
const PROGRAMME_START = new Date(2026, 5, 12); // 12 juin 2026 — début du programme DTC

// Dernière "SEMAINE" présente dans le classement (source canonique = chaîne MySQL→Sheets).
function latestWeekFromClassement() {
  try {
    const txt = fs.readFileSync(path.join(IN, "classement.csv"), "utf8").replace(/^﻿/, "");
    const weeks = [...new Set(txt.split(/\r?\n/).slice(1).map(l => (l.split(",")[1] || "").trim()).filter(w => /semaine/i.test(w)))];
    return weeks.sort((a, b) => (parseInt(b.replace(/\D/g, "")) || 0) - (parseInt(a.replace(/\D/g, "")) || 0))[0] || null;
  } catch { return null; }
}
// Dérive un cycle (dates) pour une semaine absente de CYCLES : 7 jours à partir du début programme.
function deriveCycle(week) {
  const N = parseInt(String(week).replace(/\D/g, "")) || 1;
  const s = new Date(PROGRAMME_START.getTime() + (N - 1) * 7 * 86400000);
  const e = new Date(s.getTime() + 6 * 86400000);
  const iso = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  const dmy = d => `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")}/${d.getFullYear()}`;
  const MOIS = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre"];
  const human = d => `${d.getDate()} ${MOIS[d.getMonth()]} ${d.getFullYear()}`;
  return { label: `Semaine ${N} · ${human(s)} au ${human(e)}`, reportDate: dmy(new Date(e.getTime() + 86400000)),
    ba: [iso(s), iso(e)], pos: `${human(s)} au ${human(e)}` };
}

// Paramètre semaine : accepte 1 · "1" · "Semaine 1". Sans argument → dernière semaine du classement.
const _arg = (process.argv[2] || "").trim();
const SEMAINE = _arg ? (/^\d+$/.test(_arg) ? "Semaine " + _arg : _arg) : (latestWeekFromClassement() || DEFAULT_WEEK);
const CY = CYCLES[SEMAINE] || deriveCycle(SEMAINE);

// — Valeurs structurelles (stables d'une semaine à l'autre) —
const CONFIG = {
  // Recrutement à 100 % partout (recrutés = objectif) — donnée figée
  effectif: { "ATLANTIQUE": 20, "COTONOU 1": 44, "COTONOU 2": 31, "NORD EST": 41, "NORD OUEST": 24, "SUD EST": 31, "SUD OUEST": 30 },
  baDailyGlobal: 1500000,             // objectif BA global / JOUR, réparti par % d'effectif régional
  baWorkDays: 6,                      // les BA travaillent 6 jours/semaine → cible hebdo = daily × 6
  pos: {                              // BASE GLOBALE TSA : UNIQUE POS figé ; daily/weekly recalculés depuis TSA_REF
    "ATLANTIQUE": { uniquePos:  91 },
    "COTONOU 1":  { uniquePos: 148 },
    "COTONOU 2":  { uniquePos: 158 },
    "NORD EST":   { uniquePos: 296 },
    "NORD OUEST": { uniquePos: 280 },
    "SUD EST":    { uniquePos: 520 },
    "SUD OUEST":  { uniquePos: 378 },
  },
  primeThreshold: 450000,             // prime TSA : être n°1 régional ET volume ≥ ce seuil
  noteActivations: " ",
  noteCloture:     "Maintenir le rythme jour-par-jour · Soutenir le week-end · Convertir le Top des POS",
};

// — Vue dérivée consommée par les slides —
const WEEKLY = {
  weekLabel: CY.label, reportDate: CY.reportDate, weekNo: SEMAINE,
  classementWeek: SEMAINE, baStart: CY.ba[0], baEnd: CY.ba[1], posPeriod: CY.pos,
  // recrutement [à recruter, recrutés] = 100 %
  recrutement: Object.fromEntries(Object.entries(CONFIG.effectif).map(([k, v]) => [k, [v, v]])),
  pos: CONFIG.pos,
  noteRecrutement: "", noteActivations: CONFIG.noteActivations, noteCloture: CONFIG.noteCloture,
};

// ── PALETTE (light premium MTN) ─────────────────────────────────────
const BG = "FFFFFF", INK = "1A1A1A", INK2 = "2B2B2B";
const YEL = "FFCC00", YELDK = "E0B400", YELSOFT = "FFF3C4";
const MUTE = "6B6B6B", MUTE2 = "9A9A9A", LINE = "E4E4E4", ROW = "F4F4F4";
const GREEN = "2E9E4F", AMBER = "E8932B", REDX = "D1453B", BLUE = "1F6FB2";

// ── FONTS ───────────────────────────────────────────────────────────
const fHead = "MTN Brighter Sans Bold";
const fBody = "MTN Brighter Sans";

const LOGO = path.join(__dirname, "..", "assets", "lka_logo.png");   // logo LKA Services
const HAS_LOGO = fs.existsSync(LOGO);

const W = 13.33, H = 7.5, MX = 0.7, CW = W - 2 * MX;
const REGIONS = ["ATLANTIQUE", "COTONOU 1", "COTONOU 2", "NORD EST", "NORD OUEST", "SUD EST", "SUD OUEST"];

// ════════════════════════════════════════════════════════════════════
//  CSV PARSING + AGGREGATION
// ════════════════════════════════════════════════════════════════════
function parseCSV(text) {
  text = text.replace(/^﻿/, "");
  const rows = []; let i = 0, f = "", row = [], q = false;
  while (i < text.length) {
    const c = text[i];
    if (q) { if (c === '"') { if (text[i + 1] === '"') { f += '"'; i++; } else q = false; } else f += c; }
    else { if (c === '"') q = true; else if (c === ",") { row.push(f); f = ""; }
      else if (c === "\n") { row.push(f); rows.push(row); row = []; f = ""; }
      else if (c !== "\r") f += c; }
    i++;
  }
  if (f.length || row.length) { row.push(f); rows.push(row); }
  return rows;
}
function readCSV(file) {
  const R = parseCSV(fs.readFileSync(file, "utf8"));
  const H = R[0].map(h => h.trim());
  return R.slice(1).filter(r => r.length > 1 && r.some(c => c !== "")).map(r => {
    const o = {}; H.forEach((h, j) => o[h] = (r[j] || "").trim()); return o;
  });
}
const num = v => +String(v).replace(/[^0-9.\-]/g, "") || 0;
const fmt = n => Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, " "); // thin space
function parseTS(s) {
  const m = String(s).match(/(\d+)\/(\d+)\/(\d+)\s+(\d+):(\d+):(\d+)/);
  return m ? new Date(+m[3], +m[1] - 1, +m[2], +m[4], +m[5], +m[6]) : null;
}

// ── helpers "jour" ──────────────────────────────────────────────────
const JOURS = ["Dim", "Lun", "Mar", "Mer", "Jeu", "Ven", "Sam"];
const dayLabel = k => { const [y, m, d] = k.split("-").map(Number); return JOURS[new Date(y, m - 1, d).getDay()] + " " + d; };

// ════════════════════════════════════════════════════════════════════
//  DONNÉES — deux modes :
//   • PROD : JSON déjà calculé par l'Apps Script (report.json / $REPORT_JSON)
//   • DEV  : calcul depuis les CSV locaux (inputs/*.csv)
//  Les slides ne lisent que les variables décompactées plus bas (D).
// ════════════════════════════════════════════════════════════════════
function computeFromCSV() {
  const classement = readCSV(path.join(IN, "classement.csv"));
  const allWeeks = [...new Set(classement.map(r => r.SEMAINE))];
  const weekNum = w => parseInt(String(w).replace(/\D/g, "")) || 0;
  const CWEEK = WEEKLY.classementWeek === "auto"
    ? allWeeks.sort((a, b) => weekNum(b) - weekNum(a))[0] : WEEKLY.classementWeek;
  const cw = classement.filter(r => r.SEMAINE === CWEEK);

  const volByReg = {}; REGIONS.forEach(r => volByReg[r] = 0);
  cw.forEach(r => { if (volByReg[r.REGION] !== undefined) volByReg[r.REGION] += num(r.VOLUME_XAF); });
  const volTotal = Object.values(volByReg).reduce((a, b) => a + b, 0);

  // n°1 de chaque région (RANG_REGIONAL = 1), trié par volume décroissant
  const primeTSA = REGIONS.map(reg => {
    const rows = cw.filter(r => r.REGION === reg);
    return rows.find(r => num(r.RANG_REGIONAL) === 1) || rows.sort((a, b) => num(b.VOLUME_XAF) - num(a.VOLUME_XAF))[0];
  }).filter(Boolean).sort((a, b) => num(b.VOLUME_XAF) - num(a.VOLUME_XAF));

  // TSA_REF → cibles POS (12M/jour) ; UNIQUE POS reste figé
  const tsaRef = readCSV(path.join(IN, "tsa_ref.csv"));
  const posDaily = {}; REGIONS.forEach(r => posDaily[r] = 0);
  tsaRef.forEach(r => { const k = (r.REGION || "").trim().toUpperCase(); if (posDaily[k] !== undefined) posDaily[k] += num(r.OBJECTIF_DAILY); });
  const objectifTSA = Object.values(posDaily).reduce((a, b) => a + b, 0) * 7;
  REGIONS.forEach(r => { CONFIG.pos[r].daily = Math.round(posDaily[r]); CONFIG.pos[r].weekly = Math.round(posDaily[r] * 7); });

  // FORM BA → activations + montant par région + série jour-par-jour
  const form = readCSV(path.join(IN, "form.csv"));
  const H0 = Object.keys(form[0] || {});
  const col = re => H0.find(h => re.test(h)) || "";
  const cReg = col(/Region/i), cMont = col(/Montant/i), cTS = col(/Timestamp/i);
  const dStart = WEEKLY.baStart ? new Date(WEEKLY.baStart + "T00:00:00") : null;
  const dEnd = WEEKLY.baEnd ? new Date(WEEKLY.baEnd + "T23:59:59") : null;
  const inWin = d => d && (!dStart || d >= dStart) && (!dEnd || d <= dEnd);
  const baAgg = {}; REGIONS.forEach(r => baAgg[r] = { act: 0, mont: 0 });
  const dayMap = {};
  form.forEach(r => {
    const reg = (r[cReg] || "").trim().toUpperCase(); if (!baAgg[reg]) return;
    const d = parseTS(r[cTS]); if (!inWin(d)) return;
    baAgg[reg].act++; baAgg[reg].mont += num(r[cMont]);
    const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
    dayMap[key] = dayMap[key] || { mont: 0 }; dayMap[key].mont += num(r[cMont]);
  });
  // Objectif BA : 1,5M/jour réparti par % d'effectif régional ; cible semaine = daily × baWorkDays (6)
  const effTotal = Object.values(CONFIG.effectif).reduce((a, b) => a + b, 0);
  const baDailyByReg = {}; REGIONS.forEach(r => { baDailyByReg[r] = Math.round((CONFIG.effectif[r] / effTotal) * CONFIG.baDailyGlobal); });
  const dailyTarget = CONFIG.baDailyGlobal;
  const daySeries = Object.keys(dayMap).sort().map(k => ({ label: dayLabel(k), mont: dayMap[k].mont, pct: dailyTarget ? dayMap[k].mont / dailyTarget * 100 : 0 }));

  return { CWEEK, volByReg, volTotal, primeTSA, champion: primeTSA[0] || {}, objectifTSA, baAgg, baDailyByReg, baWorkDays: CONFIG.baWorkDays, daySeries, dailyTarget };
}

// JSON-contrat (envoyé par l'Apps Script) → mêmes variables internes que computeFromCSV
function applyJson(j) {
  WEEKLY.weekLabel = j.label; WEEKLY.weekNo = j.weekNo || j.semaine; WEEKLY.reportDate = j.reportDate; WEEKLY.posPeriod = j.posPeriod;
  WEEKLY.recrutement = j.recrutement;
  WEEKLY.noteActivations = (j.notes && j.notes.activations) || ""; WEEKLY.noteCloture = (j.notes && j.notes.cloture) || "";
  CONFIG.primeThreshold = j.primeThreshold;
  const volByReg = {}, baAgg = {}, baDailyByReg = {};
  REGIONS.forEach(r => {
    const p = (j.pos && j.pos[r]) || {}; CONFIG.pos[r] = { uniquePos: p.uniquePos || 0, daily: p.daily || 0, weekly: p.weekly || 0 };
    volByReg[r] = p.real || 0;
    const b = (j.ba && j.ba.rows && j.ba.rows[r]) || {};
    CONFIG.effectif[r] = b.effectif || 0; baAgg[r] = { act: b.activations || 0, mont: b.montant || 0 }; baDailyByReg[r] = b.dailyTarget || 0;
  });
  const primeTSA = (j.prime || []).map(p => ({ REGION: p.region, TSA: p.tsa, RBM: p.rbm, VOLUME_XAF: p.volume, NB_POS_ACTIFS: p.posActifs, TAUX_ACTIVATION: p.tauxAct }));
  return {
    CWEEK: j.semaine, volByReg, volTotal: Object.values(volByReg).reduce((a, b) => a + b, 0),
    primeTSA, champion: primeTSA[0] || {}, objectifTSA: j.objectifTSA,
    baAgg, baDailyByReg, baWorkDays: (j.ba && j.ba.workDays) || 6,
    daySeries: (j.ba && j.ba.days) || [], dailyTarget: (j.ba && j.ba.dailyObjectif) || 0,
  };
}

// Variables internes → JSON-contrat (sert de fixture de test + spec pour l'Apps Script)
function buildReportJson(D) {
  const pos = {}, rows = {};
  REGIONS.forEach(r => {
    pos[r] = { uniquePos: CONFIG.pos[r].uniquePos, daily: CONFIG.pos[r].daily, weekly: CONFIG.pos[r].weekly, real: D.volByReg[r] };
    rows[r] = { effectif: CONFIG.effectif[r], activations: D.baAgg[r].act, montant: D.baAgg[r].mont, dailyTarget: D.baDailyByReg[r] };
  });
  return {
    semaine: D.CWEEK, label: WEEKLY.weekLabel, weekNo: WEEKLY.weekNo, reportDate: WEEKLY.reportDate, posPeriod: WEEKLY.posPeriod,
    recrutement: WEEKLY.recrutement, primeThreshold: CONFIG.primeThreshold, objectifTSA: D.objectifTSA, pos,
    prime: D.primeTSA.map(p => ({ region: p.REGION, tsa: p.TSA, rbm: p.RBM, volume: num(p.VOLUME_XAF), posActifs: num(p.NB_POS_ACTIFS), tauxAct: num(p.TAUX_ACTIVATION) })),
    ba: { workDays: D.baWorkDays, dailyObjectif: D.dailyTarget, rows, days: D.daySeries.map(d => ({ label: d.label, mont: d.mont, pct: +d.pct.toFixed(1) })) },
    notes: { activations: WEEKLY.noteActivations, cloture: WEEKLY.noteCloture },
  };
}

const REPORT_JSON = process.env.REPORT_JSON ? path.resolve(process.env.REPORT_JSON) : path.join(IN, "report.json");
const DUMP = process.argv.includes("--dump");
const useJson = !DUMP && fs.existsSync(REPORT_JSON);
console.log(useJson ? `🧩 Rendu depuis ${REPORT_JSON}` : "🧮 Calcul depuis les CSV (mode dev)");
const D = useJson ? applyJson(JSON.parse(fs.readFileSync(REPORT_JSON, "utf8"))) : computeFromCSV();
if (DUMP) {
  fs.mkdirSync(IN, { recursive: true });
  fs.writeFileSync(path.join(IN, "report.json"), JSON.stringify(buildReportJson(D), null, 2), "utf8");
  console.log("📤 report.json écrit (JSON-contrat) → " + path.join(IN, "report.json"));
  process.exit(0);
}
const { CWEEK, volByReg, volTotal, primeTSA, champion, objectifTSA, baAgg, baDailyByReg, baWorkDays, daySeries, dailyTarget } = D;
console.log(`📅 Semaine rendue : ${CWEEK}`);

// ════════════════════════════════════════════════════════════════════
//  PPTX SETUP
// ════════════════════════════════════════════════════════════════════
const EnhancedPPTXGenJS = withPPTXEmbedFonts(pptxgen);
const pres = new EnhancedPPTXGenJS();
pres.layout = "LAYOUT_WIDE";
pres.title = "Progress Report DTC Assisted — MTN Bénin";
pres.author = "Contribution LKA";
pres.company = "MTN Bénin";

const TOTAL = 6;

// ── HELPERS ─────────────────────────────────────────────────────────
function light(slide) { slide.background = { color: BG }; }

function chrome(slide, page, section) {
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 0.09, fill: { color: YEL }, line: { type: "none" } });
  if (section) slide.addText(section.toUpperCase(), {
    x: MX, y: H - 0.46, w: 8.5, h: 0.26, align: "left", margin: 0,
    fontFace: fBody, fontSize: 8.5, charSpacing: 3, color: MUTE2,
  });
  if (HAS_LOGO) slide.addImage({ path: LOGO, x: W - MX - 1.18, y: H - 0.62, w: 0.5, h: 0.46 });
  slide.addShape(pres.shapes.RECTANGLE, { x: W - MX - 0.42, y: H - 0.58, w: 0.42, h: 0.42, fill: { color: YEL }, line: { type: "none" } });
  slide.addText(String(page).padStart(2, "0"), {
    x: W - MX - 0.42, y: H - 0.58, w: 0.42, h: 0.42, align: "center", valign: "middle",
    margin: 0, fontFace: fHead, bold: true, fontSize: 13, color: "000000",
  });
}

function header(slide, title, accent) {
  slide.addShape(pres.shapes.RECTANGLE, { x: MX, y: 0.5, w: 0.24, h: 0.5, fill: { color: YEL }, line: { type: "none" } });
  slide.addText(title.toUpperCase(), {
    x: MX + 0.4, y: 0.44, w: CW - 0.4, h: 0.62, align: "left", valign: "middle", margin: 0,
    fontFace: fHead, bold: true, fontSize: 28, color: INK, charSpacing: 0.5,
  });
  if (accent) slide.addText(accent.toUpperCase(), {
    x: MX + 0.4, y: 1.06, w: CW - 0.4, h: 0.3, align: "left", margin: 0,
    fontFace: fBody, fontSize: 11, color: MUTE, charSpacing: 2,
  });
}

// light premium table
function table(slide, x, y, w, headers, rows, opts = {}) {
  const hdrRow = headers.map((h, j) => ({
    text: h, options: {
      bold: true, color: "000000", fill: { color: YEL }, fontSize: opts.hfs || 10,
      fontFace: fHead, align: j === 0 ? "left" : "center", valign: "middle", margin: [0.03, 0.06, 0.03, 0.08],
    },
  }));
  const body = [hdrRow];
  rows.forEach((r, i) => body.push(r.map((cell, j) => {
    const o = typeof cell === "object" ? cell : { text: cell };
    return { text: o.text, options: {
      bold: !!o.bold, color: o.color || (j === 0 ? INK : INK2),
      fill: { color: i % 2 === 0 ? BG : ROW },
      fontSize: opts.fs || 10, fontFace: o.bold ? fHead : fBody,
      align: j === 0 ? "left" : "center", valign: "middle", margin: [0.03, 0.06, 0.03, 0.08],
    } };
  })));
  if (opts.totalRow) body.push(opts.totalRow.map((cell, j) => {
    const o = typeof cell === "object" ? cell : { text: cell };
    return { text: o.text, options: {
      bold: true, color: o.color || "000000", fill: { color: YEL }, fontSize: opts.fs || 10,
      fontFace: fHead, align: j === 0 ? "left" : "center", valign: "middle", margin: [0.03, 0.06, 0.03, 0.08],
    } };
  }));
  slide.addTable(body, {
    x, y, w, colW: opts.colWidths, rowH: opts.rowH || 0.34,
    border: { type: "solid", pt: 0.5, color: LINE }, autoPage: false,
  });
}

const tauxColor = p => p >= 90 ? GREEN : p >= 50 ? AMBER : REDX;

// ════════════════════════════════════════════════════════════════════
//  SLIDE 1 · COUVERTURE (page de garde)
// ════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide(); light(s);
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 0.09, fill: { color: YEL }, line: { type: "none" } });
  // grand bloc jaune en bas à gauche (motif signature)
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 4.95, w: 4.5, h: 2.55, fill: { color: YEL }, line: { type: "none" } });
  s.addShape(pres.shapes.RECTANGLE, { x: 4.5, y: 6.55, w: 0.95, h: 0.95, fill: { color: ROW }, line: { type: "none" } });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.45, y: 6.55, w: 0.95, h: 0.95, fill: { color: YELSOFT }, line: { type: "none" } });

  if (HAS_LOGO) s.addImage({ path: LOGO, x: MX, y: 0.55, w: 1.0, h: 0.93 });

  s.addText("MTN BÉNIN  ·  REPORTING TERRAIN", {
    x: MX, y: 1.95, w: 9, h: 0.3, margin: 0, fontFace: fBody, fontSize: 13, color: YELDK, charSpacing: 3,
  });
  s.addText("RAPPORT", { x: MX - 0.04, y: 2.3, w: 11, h: 1.1, margin: 0, fontFace: fHead, bold: true, fontSize: 70, color: INK });
  s.addText("HEBDOMADAIRE", { x: MX - 0.04, y: 3.32, w: 12, h: 1.1, margin: 0, fontFace: fHead, bold: true, fontSize: 70, color: INK });
  s.addText("DTC ASSISTED", { x: MX, y: 4.42, w: 11, h: 0.55, margin: 0, fontFace: fHead, bold: true, fontSize: 26, color: YELDK, charSpacing: 1 });

  // bloc jaune : sous-titre
  s.addText("Contribution LKA", { x: 0.5, y: 5.55, w: 3.8, h: 0.45, margin: 0, fontFace: fHead, bold: true, fontSize: 22, color: "000000" });
  s.addText(WEEKLY.weekLabel, { x: 0.5, y: 6.08, w: 3.8, h: 0.7, margin: 0, fontFace: fBody, fontSize: 12.5, color: "1A1A1A", lineSpacingMultiple: 1.05 });

  // colonne meta droite
  s.addText("OBJECTIF", { x: 9.6, y: 5.0, w: 3.13, h: 0.3, margin: 0, fontFace: fBody, fontSize: 11, color: MUTE, charSpacing: 3, align: "right" });
  s.addText("60%", { x: 9.6, y: 5.22, w: 3.13, h: 1.0, margin: 0, fontFace: fHead, bold: true, fontSize: 64, color: YELDK, align: "right" });
  s.addText("DTC EOY 2026", { x: 9.6, y: 6.32, w: 3.13, h: 0.3, margin: 0, fontFace: fBody, fontSize: 12, color: INK, charSpacing: 2, align: "right" });

  // date d'édition haut droite
  s.addText([
    { text: WEEKLY.weekNo + "\n", options: { fontFace: fHead, bold: true, fontSize: 14, color: INK, breakLine: true } },
    { text: "Édité le " + WEEKLY.reportDate, options: { fontFace: fBody, fontSize: 11, color: MUTE } },
  ], { x: 8.73, y: 0.62, w: 4, h: 0.8, margin: 0, align: "right", valign: "top", lineSpacingMultiple: 1.1 });
}

// ════════════════════════════════════════════════════════════════════
//  SLIDE 2 · ÉTAT DU RECRUTEMENT
// ════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide(); light(s);
  chrome(s, 2, "État d'avancement · Recrutement des BA");
  header(s, "État du recrutement", "Renforcement de la force de vente terrain pour l'opération DTC Assisted");

  s.addText(`Après avoir défini la stratégie et les objectifs, voici l'état d'avancement du recrutement des Brand Ambassadors (BA), qui constituent la force de soutien des TSA, à la date du ${WEEKLY.reportDate}.`, {
    x: MX, y: 1.5, w: CW, h: 0.5, margin: 0, fontFace: fBody, fontSize: 12, color: INK2, lineSpacingMultiple: 1.1,
  });

  const rows = []; let tAR = 0, tR = 0;
  REGIONS.forEach(reg => {
    const [ar, rec] = WEEKLY.recrutement[reg] || [0, 0];
    tAR += ar; tR += rec;
    const p = ar ? Math.round(rec / ar * 100) : 0;
    rows.push([reg, fmt(ar), fmt(rec), { text: p + "%", color: tauxColor(p), bold: true }]);
  });
  const gp = tAR ? Math.round(tR / tAR * 100) : 0;
  const tx = MX, tw = 7.0, ty = 2.2;
  table(s, tx, ty, tw, ["RÉGION", "À RECRUTER", "RECRUTÉS", "TAUX"], rows, {
    colWidths: [tw * 0.34, tw * 0.24, tw * 0.22, tw * 0.20], rowH: 0.40,
    totalRow: ["TOTAL", fmt(tAR), fmt(tR), gp + "%"],
  });

  // KPI droite
  const rx = MX + tw + 0.45, rw = W - MX - rx;
  s.addShape(pres.shapes.RECTANGLE, { x: rx, y: ty, w: rw, h: 1.9, fill: { color: YEL }, line: { type: "none" } });
  s.addText("TAUX GLOBAL DE RECRUTEMENT", { x: rx + 0.25, y: ty + 0.22, w: rw - 0.5, h: 0.3, margin: 0, fontFace: fHead, bold: true, fontSize: 11, color: "000000", charSpacing: 1 });
  s.addText(gp + "%", { x: rx + 0.18, y: ty + 0.5, w: rw - 0.36, h: 1.25, margin: 0, fontFace: fHead, bold: true, fontSize: 76, color: "000000" });

  const nbDone = REGIONS.filter(r => { const [a, b] = WEEKLY.recrutement[r] || [0, 0]; return a && b >= a; }).length;
  const lagging = REGIONS.filter(r => { const [a, b] = WEEKLY.recrutement[r] || [0, 0]; return a && b / a < 0.5; });
  const insights = [
    { c: GREEN, t: `${nbDone} région${nbDone > 1 ? "s" : ""} sur ${REGIONS.length} à 100 % de leur objectif.` },
  ];
  if (lagging.length) insights.push({ c: REDX, t: `${lagging.join(", ")} en retard — déploiement à accélérer.` });
  insights.push({ c: YEL, t: `${gp} % global · ${fmt(tR)} BA recrutés sur ${fmt(tAR)} prévus.` });
  if (WEEKLY.noteRecrutement) insights.push({ c: YEL, t: WEEKLY.noteRecrutement });

  let iy = ty + 2.15;
  insights.forEach(it => {
    s.addShape(pres.shapes.RECTANGLE, { x: rx, y: iy, w: 0.1, h: 0.6, fill: { color: it.c }, line: { type: "none" } });
    s.addText(it.t, { x: rx + 0.25, y: iy, w: rw - 0.3, h: 0.6, margin: 0, fontFace: fBody, fontSize: 11, color: INK, valign: "middle", lineSpacingMultiple: 1.0 });
    iy += 0.72;
  });
}

// ════════════════════════════════════════════════════════════════════
//  SLIDE 5 · ACTIVATIONS BA — PROGRAMME FORCE DE SOUTIEN AUX TSA
// ════════════════════════════════════════════════════════════════════
const slideBA = () => {
  const s = pres.addSlide(); light(s);
  chrome(s, 5, "Force de soutien · Activations BA");
  header(s, "Activations BA — Force de soutien aux TSA", `Programme BA en appui des TSA · ${WEEKLY.weekLabel}`);

  const rows = []; let tEff = 0, tAct = 0, tTar = 0, tMon = 0;
  REGIONS.forEach(reg => {
    const eff = CONFIG.effectif[reg] || 0;                       // effectif = recrutement (figé)
    const o = baAgg[reg]; const act = o.act, mont = o.mont;
    const target = baDailyByReg[reg] * baWorkDays;             // cible hebdo = daily région × 6 jours ouvrés
    const p = target ? mont / target * 100 : 0;
    tEff += eff; tAct += act; tTar += target; tMon += mont;
    rows.push([reg, fmt(eff), fmt(act), fmt(target), fmt(mont), { text: p.toFixed(1) + "%", color: p >= 50 ? GREEN : p >= 20 ? AMBER : REDX, bold: true }]);
  });
  const tp = tTar ? tMon / tTar * 100 : 0;
  const tx = MX, tw = 7.35, ty = 2.1;
  table(s, tx, ty, tw, ["RÉGION", "EFFECTIF", "ACTIVATIONS", "TARGET", "MONTANT", "TAUX"], rows, {
    colWidths: [tw * 0.22, tw * 0.15, tw * 0.20, tw * 0.17, tw * 0.16, tw * 0.10],
    rowH: 0.40, fs: 9.5, hfs: 8.5,
    totalRow: ["TOTAL", fmt(tEff), fmt(tAct), fmt(tTar), fmt(tMon), tp.toFixed(1) + "%"],
  });

  // ── colonne droite : 2 callouts compacts + rythme jour-par-jour ──
  const rx = MX + tw + 0.4, rw = W - MX - rx;
  const cwid = (rw - 0.2) / 2;
  // callout ACTIVATIONS (jaune)
  s.addShape(pres.shapes.RECTANGLE, { x: rx, y: ty, w: cwid, h: 1.05, fill: { color: YEL }, line: { type: "none" } });
  s.addText("ACTIVATIONS", { x: rx + 0.16, y: ty + 0.13, w: cwid - 0.3, h: 0.24, margin: 0, fontFace: fHead, bold: true, fontSize: 9.5, color: "000000", charSpacing: 1.5 });
  s.addText(fmt(tAct), { x: rx + 0.14, y: ty + 0.36, w: cwid - 0.28, h: 0.62, margin: 0, fontFace: fHead, bold: true, fontSize: 34, color: "000000", valign: "middle" });
  // callout MONTANT (bordé)
  const mx2 = rx + cwid + 0.2;
  s.addShape(pres.shapes.RECTANGLE, { x: mx2, y: ty, w: cwid, h: 1.05, fill: { color: BG }, line: { color: LINE, width: 1 } });
  s.addText("MONTANT (FCFA)", { x: mx2 + 0.16, y: ty + 0.13, w: cwid - 0.3, h: 0.24, margin: 0, fontFace: fBody, fontSize: 9, color: MUTE, charSpacing: 1.2 });
  s.addText(fmt(tMon), { x: mx2 + 0.14, y: ty + 0.36, w: cwid - 0.28, h: 0.62, margin: 0, fontFace: fHead, bold: true, fontSize: 24, color: INK, valign: "middle" });

  // panneau rythme journalier
  const py = ty + 1.3;
  s.addText("RYTHME JOURNALIER", { x: rx, y: py, w: rw - 1.4, h: 0.26, margin: 0, fontFace: fHead, bold: true, fontSize: 11, color: INK, charSpacing: 1 });
  s.addText("% vs objectif/j", { x: rx + rw - 1.5, y: py + 0.02, w: 1.5, h: 0.24, margin: 0, fontFace: fBody, fontSize: 8.5, color: MUTE, align: "right" });
  // échelle : 100 % = objectif journalier (15 000 × effectif). On élargit si dépassement.
  const maxPct = Math.max(100, ...daySeries.map(d => d.pct));
  const scaleMax = Math.ceil(maxPct / 20) * 20;
  const lblW = 0.66, valW = 0.62, gap = 0.08;
  const trackX = rx + lblW + gap, trackW = rw - lblW - valW - gap * 2;
  const rowH = 0.305, by = py + 0.4;
  daySeries.forEach((d, i) => {
    const y = by + i * rowH;
    const pr = Math.round(d.pct);
    const c = pr >= 90 ? GREEN : pr >= 50 ? AMBER : REDX;
    s.addText(d.label, { x: rx, y, w: lblW, h: rowH - 0.05, margin: 0, fontFace: fBody, fontSize: 8.5, color: MUTE, valign: "middle" });
    // piste
    s.addShape(pres.shapes.RECTANGLE, { x: trackX, y: y + 0.045, w: trackW, h: rowH - 0.14, fill: { color: ROW }, line: { type: "none" } });
    // remplissage
    const fillW = Math.max(0.02, Math.min(1, d.pct / scaleMax) * trackW);
    s.addShape(pres.shapes.RECTANGLE, { x: trackX, y: y + 0.045, w: fillW, h: rowH - 0.14, fill: { color: c }, line: { type: "none" } });
    // valeur %
    s.addText(pr + "%", { x: rx + rw - valW, y, w: valW, h: rowH - 0.05, margin: 0, fontFace: fHead, bold: true, fontSize: 8.5, color: c, align: "right", valign: "middle" });
  });
  // repère objectif (ligne verticale à 100 %)
  const objX = trackX + (100 / scaleMax) * trackW;
  const barsBottom = by + daySeries.length * rowH;
  s.addShape(pres.shapes.RECTANGLE, { x: objX, y: by - 0.02, w: 0.014, h: barsBottom - by - 0.05, fill: { color: INK }, line: { type: "none" } });
  s.addText("objectif", { x: objX - 0.6, y: barsBottom, w: 1.2, h: 0.2, margin: 0, fontFace: fBody, fontSize: 7.5, color: MUTE, align: "center" });

  if (WEEKLY.noteActivations) s.addText(WEEKLY.noteActivations, {
    x: MX, y: 6.45, w: tw, h: 0.5, margin: 0, fontFace: fBody, italic: true, fontSize: 9.5, color: REDX, lineSpacingMultiple: 1.05, valign: "top",
  });
};

// ════════════════════════════════════════════════════════════════════
//  SLIDE 3 · ACTIVATIONS POS — PERFORMANCE BASE GLOBALE TSA
// ════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide(); light(s);
  chrome(s, 3, "État d'avancement · Performance base globale TSA");
  header(s, "Activations POS — Performance par région", `Volume des points de vente (base globale TSA) · ${WEEKLY.posPeriod}`);

  s.addText(`Performances globales des points de vente sur la période du ${WEEKLY.posPeriod} — réalisation rapportée aux cibles hebdomadaires par région.`, {
    x: MX, y: 1.5, w: CW, h: 0.4, margin: 0, fontFace: fBody, fontSize: 12, color: INK2,
  });

  const rows = []; let tPos = 0, tD = 0, tW = 0, tReal = 0;
  REGIONS.forEach(reg => {
    const p = WEEKLY.pos[reg] || { uniquePos: 0, daily: 0, weekly: 0 };
    const real = volByReg[reg] || 0;
    const taux = p.weekly ? real / p.weekly * 100 : 0;
    tPos += p.uniquePos; tD += p.daily; tW += p.weekly; tReal += real;
    rows.push([reg, fmt(p.uniquePos), fmt(p.daily), fmt(p.weekly), fmt(real), { text: taux.toFixed(1) + "%", color: taux >= 5 ? GREEN : taux >= 2.5 ? AMBER : REDX, bold: true }]);
  });
  const tt = tW ? tReal / tW * 100 : 0;
  const ty = 2.0, rowH = 0.46, tw = CW;
  table(s, MX, ty, tw, ["RÉGION", "UNIQUE POS", "TARGET DAILY", "TARGET WEEKLY", "RÉALISATION", "TAUX"], rows, {
    colWidths: [tw * 0.18, tw * 0.13, tw * 0.18, tw * 0.18, tw * 0.21, tw * 0.12],
    rowH, fs: 10.5, hfs: 9.5,
    totalRow: ["TOTAL", fmt(tPos), fmt(tD), fmt(tW), fmt(tReal), tt.toFixed(1) + "%"],
  });

  // note de lecture (le TOTAL de la table reprend déjà les chiffres de synthèse)
  s.addText(`Cible hebdomadaire totale : ${fmt(tW)} FCFA · Réalisé : ${fmt(tReal)} FCFA · Taux de réalisation global : ${tt.toFixed(1)} %.`, {
    x: MX, y: 6.55, w: CW, h: 0.35, margin: 0, fontFace: fBody, italic: true, fontSize: 10.5, color: MUTE,
  });
}

// ════════════════════════════════════════════════════════════════════
//  SLIDE 4 · PRIME & RANG — CLASSEMENT TSA HEBDOMADAIRE
// ════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide(); light(s);
  chrome(s, 4, "Reconnaissance · Classement TSA hebdomadaire");
  header(s, "Prime & Rang", `Classement hebdomadaire — le n°1 de chaque région · ${CWEEK}`);

  const taux = objectifTSA ? volTotal / objectifTSA * 100 : 0;
  // KPI cards (rangée)
  const kpis = [
    { l: "OBJECTIF SEMAINE", v: fmt(objectifTSA), c: INK },
    { l: "VOLUME TOTAL", v: fmt(volTotal), c: INK },
    { l: "TAUX DE RÉALISATION", v: taux.toFixed(2) + "%", c: YELDK },
    { l: "TSA DE LA SEMAINE", v: champion.TSA ? toTitle(champion.TSA) : "—", c: GREEN, small: true },
  ];
  const ky = 1.55, kwid = (CW - 0.3 * 3) / 4, kh = 1.05;
  kpis.forEach((k, i) => {
    const x = MX + i * (kwid + 0.3);
    s.addShape(pres.shapes.RECTANGLE, { x, y: ky, w: kwid, h: kh, fill: { color: ROW }, line: { type: "none" } });
    s.addShape(pres.shapes.RECTANGLE, { x, y: ky, w: 0.08, h: kh, fill: { color: YEL }, line: { type: "none" } });
    s.addText(k.l, { x: x + 0.2, y: ky + 0.13, w: kwid - 0.36, h: 0.25, margin: 0, fontFace: fBody, fontSize: 8.5, color: MUTE, charSpacing: 1.5 });
    s.addText(k.v, { x: x + 0.18, y: ky + 0.36, w: kwid - 0.32, h: 0.6, margin: 0, fontFace: fHead, bold: true, fontSize: k.small ? 14 : 26, color: k.c, valign: "middle" });
  });

  // les 7 lauréats (n°1 régional) triés par volume — prime TSA si volume ≥ seuil
  let nbElig = 0;
  const rows = primeTSA.map(r => {
    const elig = num(r.VOLUME_XAF) >= CONFIG.primeThreshold; if (elig) nbElig++;
    return [r.REGION, toTitle(r.TSA), r.RBM ? toTitle(r.RBM) : "—",
      fmt(r.VOLUME_XAF), fmt(r.NB_POS_ACTIFS), num(r.TAUX_ACTIVATION).toFixed(1) + "%",
      { text: elig ? "OUI" : "NON", color: elig ? GREEN : REDX, bold: true }];
  });
  // bandeau règle + verdict d'éligibilité
  const elgOK = nbElig > 0, bnY = 2.62;
  s.addShape(pres.shapes.RECTANGLE, { x: MX, y: bnY, w: CW, h: 0.4, fill: { color: elgOK ? "E9F6EC" : YELSOFT }, line: { type: "none" } });
  s.addShape(pres.shapes.RECTANGLE, { x: MX, y: bnY, w: 0.08, h: 0.4, fill: { color: elgOK ? GREEN : YELDK }, line: { type: "none" } });
  s.addText([
    { text: "Prime TSA : ", options: { fontFace: fHead, bold: true, color: INK } },
    { text: `n°1 régional ET volume ≥ ${fmt(CONFIG.primeThreshold)} FCFA — `, options: { fontFace: fBody, color: INK2 } },
    { text: elgOK ? `${nbElig} lauréat${nbElig > 1 ? "s" : ""} cette semaine.` : `aucun éligible cette semaine (meilleur volume : ${fmt(num(champion.VOLUME_XAF))} FCFA).`, options: { fontFace: fHead, bold: true, color: elgOK ? GREEN : REDX } },
  ], { x: MX + 0.22, y: bnY, w: CW - 0.4, h: 0.4, margin: 0, valign: "middle", fontSize: 10.5 });

  const ty = 3.18, tw = CW;
  table(s, MX, ty, tw, ["RÉGION", "TSA (N°1 RÉGIONAL)", "RBM", "VOLUME (XAF)", "POS ACTIFS", "TAUX ACT.", "PRIME"], rows, {
    colWidths: [tw * 0.15, tw * 0.25, tw * 0.21, tw * 0.13, tw * 0.10, tw * 0.09, tw * 0.07],
    rowH: 0.40, fs: 9.5, hfs: 9,
  });
}

// SLIDE 5 · ACTIVATIONS BA — rendue ici (après Prime & Rang)
slideBA();

// ════════════════════════════════════════════════════════════════════
//  SLIDE 6 · CLÔTURE (page de clôture)
// ════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide(); light(s);
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 0.09, fill: { color: YEL }, line: { type: "none" } });
  // grand bloc jaune à droite
  s.addShape(pres.shapes.RECTANGLE, { x: 9.55, y: 0, w: 3.78, h: H, fill: { color: YEL }, line: { type: "none" } });
  s.addShape(pres.shapes.RECTANGLE, { x: 9.55, y: 0, w: 0.95, h: 0.95, fill: { color: ROW }, line: { type: "none" } });

  if (HAS_LOGO) s.addImage({ path: LOGO, x: MX, y: 0.85, w: 1.0, h: 0.93 });

  s.addText("MERCI", { x: MX - 0.06, y: 2.5, w: 9, h: 1.5, margin: 0, fontFace: fHead, bold: true, fontSize: 92, color: INK });
  s.addText("Ensemble vers 60 % DTC EOY 2026", { x: MX, y: 4.05, w: 8.5, h: 0.4, margin: 0, fontFace: fBody, fontSize: 16, color: YELDK, charSpacing: 1 });

  s.addShape(pres.shapes.RECTANGLE, { x: MX, y: 5.05, w: 0.24, h: 1.0, fill: { color: YEL }, line: { type: "none" } });
  s.addText([
    { text: "Prochaines étapes\n", options: { fontFace: fHead, bold: true, fontSize: 13, color: INK, breakLine: true } },
    { text: WEEKLY.noteCloture, options: { fontFace: fBody, fontSize: 11.5, color: MUTE } },
  ], { x: MX + 0.42, y: 5.05, w: 8.5, h: 1.1, margin: 0, valign: "top", lineSpacingMultiple: 1.2 });

  s.addText("RAPPORT HEBDOMADAIRE · DTC ASSISTED", { x: 10.0, y: 2.3, w: 5.0, h: 0.4, margin: 0, fontFace: fHead, bold: true, fontSize: 13, color: "000000", rotate: 90, align: "center" });
  s.addText(WEEKLY.weekLabel, { x: MX, y: 6.55, w: 8.5, h: 0.3, margin: 0, fontFace: fBody, fontSize: 10, color: MUTE2, charSpacing: 1 });
}

// ── util : Title Case pour les noms en MAJUSCULES ───────────────────
function toTitle(s) {
  return String(s).toLowerCase().replace(/\b([a-zàâäéèêëïîôöùûüç])/g, c => c.toUpperCase()).replace(/\b-([a-z])/g, (m, c) => "-" + c.toUpperCase());
}

// ── EMBED FONTS & WRITE ─────────────────────────────────────────────
async function embedFonts() {
  try {
    const reg = fs.readFileSync(path.join(__dirname, "..", "fonts", "MTNBrighterSans-Regular.ttf"));
    const bold = fs.readFileSync(path.join(__dirname, "..", "fonts", "MTNBrighterSans-Bold.ttf"));
    await pres.addFont({ fontFace: fBody, fontFile: reg, fontType: "ttf" });
    await pres.addFont({ fontFace: fHead, fontFile: bold, fontType: "ttf" });
    console.log("✅ Fonts embedded");
  } catch (e) { console.warn("⚠️ Font embed failed:", e.message); }
}

(async () => {
  await embedFonts();
  const outDir = process.env.OUT_DIR ? path.resolve(process.env.OUT_DIR) : path.join(__dirname, "..", "outputs");
  fs.mkdirSync(outDir, { recursive: true });
  const tag = String(CWEEK).replace(/[^0-9A-Za-z]+/g, "");
  let out = path.join(outDir, `Progress_Report_DTC_Assisted_LKA_${tag}.pptx`);
  try {
    await pres.writeFile({ fileName: out });
  } catch (e) {
    if (e.code !== "EBUSY" && e.code !== "EPERM") throw e;
    out = path.join(outDir, `Progress_Report_DTC_Assisted_LKA_${tag}_${Date.now().toString().slice(-4)}.pptx`);
    await pres.writeFile({ fileName: out });
    console.warn("⚠️ Fichier principal verrouillé (ouvert ?) — écrit sous un nom alternatif.");
  }
  console.log(`✅ Done · 6 slides · classement=${CWEEK} · ${out}`);
})();
