// ════════════════════════════════════════════════════════════════════
//  PROGRESS REPORT · DTC ASSISTED · MTN BÉNIN · Contribution LKA
//  Rapport MENSUEL (POS + BA uniquement) — même style que le rapport hebdo,
//  mais agrégé sur un mois calendaire (borné au démarrage du programme, 15 juin 2026).
//
//  Lancement : node scripts/monthly_report_dtc_assisted.js [YYYY-MM]
//   (sans argument → mois du dernier Timestamp trouvé dans form.csv)
//
//  Exemple : node scripts/monthly_report_dtc_assisted.js 2026-06
//   → fenêtre réelle = 15 au 30 juin 2026 (le programme n'a démarré que le 15).
//   Les mois suivants (2026-07, 2026-08, …) couvrent le mois calendaire complet.
//
//  RÈGLE (confirmée avec l'utilisateur, cf. mémoire [[progress-report-dtc-assisted]]) :
//   objectif du mois = objectif JOURNALIER (même taux que l'hebdo) × nombre de jours
//   de la fenêtre — jours ouvrés (hors dimanche) pour le BA, tous les jours pour le POS.
//   Réalisé BA = form filtré par Timestamp sur la fenêtre.
//   Réalisé POS = somme des "Semaine N" du classement dont la période chevauche la fenêtre
//   (le classement n'a pas de date par ligne, seulement une semaine POS ven→jeu).
//
//  Données dans INPUT_DIR (défaut ./inputs), mêmes fichiers que le rapport hebdo :
//   classement.csv (CLASSEMENT_TSA) · tsa_ref.csv (TSA_REF) · form.csv (Form BA)
// ════════════════════════════════════════════════════════════════════
const pptxgen = require("pptxgenjs");
const { withPPTXEmbedFonts } = require("pptx-embed-fonts/pptxgenjs");
const fs = require("fs");
const path = require("path");
const T = require("./lib/theme");
const {
  BG, INK, INK2, YEL, YELDK, MUTE, MUTE2, LINE, ROW, GREEN, AMBER, REDX,
  fHead, fBody, HAS_LOGO, LOGO, W, H, MX, CW, REGIONS, CONFIG, PROGRAMME_START, BA_START,
  readCSV, num, parseTS, MOIS, dayLabel, fmt, light, chrome, header, table, embedFonts,
} = T;

const IN = process.env.INPUT_DIR ? path.resolve(process.env.INPUT_DIR) : path.join(__dirname, "..", "inputs");

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

// ── Fenêtre du mois : bornée au démarrage réel du programme (15 juin 2026) ──
const monthFirst = new Date(YY, MM - 1, 1);
const monthLast = new Date(YY, MM, 0); // dernier jour du mois
// Deux ancres décalées de 3 jours (comme le rapport hebdo) : POS démarre le 12 juin,
// BA le 15 juin. Chaque flux a donc SA PROPRE fenêtre, clampée à SON propre départ —
// ne jamais utiliser une fenêtre unique pour les deux (le POS perdrait 3 jours de cible/réalisé).
const winEnd = monthLast;
const winEndTS = new Date(winEnd.getTime() + 86400000 - 1); // fin de journée incluse
const posWinStart = monthFirst < PROGRAMME_START ? PROGRAMME_START : monthFirst;
const winStart = monthFirst < BA_START ? BA_START : monthFirst; // fenêtre BA
const nbCalendarDaysPOS = Math.round((winEnd - posWinStart) / 86400000) + 1;
const nbCalendarDays = Math.round((winEnd - winStart) / 86400000) + 1; // (BA, non utilisé pour la cible mais gardé pour affichage)
let nbWorkDays = 0;
for (let t = winStart.getTime(); t <= winEnd.getTime(); t += 86400000) if (new Date(t).getDay() !== 0) nbWorkDays++;
const dmy = d => `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")}/${d.getFullYear()}`;
const MONTH_LABEL = `${MOIS[MM - 1].toUpperCase()} ${YY}`;
const WINDOW_LABEL = `${dmy(winStart)} au ${dmy(winEnd)}`; // fenêtre BA (affichée sur la slide BA)
const POS_WINDOW_LABEL = `${dmy(posWinStart)} au ${dmy(winEnd)}`; // fenêtre POS (affichée sur la slide POS)
console.log(`📅 Rapport mensuel : ${MONTH_LABEL} · BA ${WINDOW_LABEL} (${nbCalendarDays}j, ${nbWorkDays}j ouvrés) · POS ${POS_WINDOW_LABEL} (${nbCalendarDaysPOS}j)`);

// ════════════════════════════════════════════════════════════════════
//  DONNÉES
// ════════════════════════════════════════════════════════════════════
// — POS : cibles journalières (TSA_REF) + réalisé (classement, par semaines chevauchant la fenêtre) —
const tsaRef = readCSV(path.join(IN, "tsa_ref.csv"));
const posDaily = {}; REGIONS.forEach(r => posDaily[r] = 0);
tsaRef.forEach(r => { const k = (r.REGION || "").trim().toUpperCase(); if (posDaily[k] !== undefined) posDaily[k] += num(r.OBJECTIF_DAILY); });

const classement = readCSV(path.join(IN, "classement.csv"));
const weekNum = w => parseInt(String(w).replace(/\D/g, "")) || 0;
// période POS d'une semaine N (ven→jeu, ancrée PROGRAMME_START) — même formule que le rapport hebdo
function posWeekWindow(N) {
  const MS = 86400000;
  const ps = new Date(PROGRAMME_START.getTime() + (N - 1) * 7 * MS);
  const pe = new Date(ps.getTime() + 6 * MS);
  return [ps, pe];
}
const allWeeks = [...new Set(classement.map(r => weekNum(r.SEMAINE)))].filter(Boolean);
const includedWeeks = allWeeks.filter(N => { const [ps, pe] = posWeekWindow(N); return ps <= winEnd && pe >= posWinStart; });
const volByReg = {}; REGIONS.forEach(r => volByReg[r] = 0);
classement.forEach(r => { if (includedWeeks.includes(weekNum(r.SEMAINE)) && volByReg[r.REGION] !== undefined) volByReg[r.REGION] += num(r.VOLUME_XAF); });
const volTotal = Object.values(volByReg).reduce((a, b) => a + b, 0);

// — BA : activations + montant par région (form filtré sur la fenêtre) + série jour-par-jour —
const form = readCSV(path.join(IN, "form.csv"));
const H0 = Object.keys(form[0] || {});
const col = re => H0.find(h => re.test(h)) || "";
const cReg = col(/Region/i), cMont = col(/Montant/i), cTS = col(/Timestamp/i);
const baAgg = {}; REGIONS.forEach(r => baAgg[r] = { act: 0, mont: 0 });
const dayMap = {};
form.forEach(r => {
  const reg = (r[cReg] || "").trim().toUpperCase(); if (!baAgg[reg]) return;
  const d = parseTS(r[cTS]); if (!d || d < winStart || d > winEndTS) return;
  baAgg[reg].act++; baAgg[reg].mont += num(r[cMont]);
  const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  dayMap[key] = (dayMap[key] || 0) + num(r[cMont]);
});
const effTotal = Object.values(CONFIG.effectif).reduce((a, b) => a + b, 0);
const baDailyByReg = {}; REGIONS.forEach(r => { baDailyByReg[r] = Math.round((CONFIG.effectif[r] / effTotal) * CONFIG.baDailyGlobal); });
const dailyTarget = CONFIG.baDailyGlobal;

// — rythme HEBDOMADAIRE (une barre par semaine BA du mois, clippée à la fenêtre) —
// chaque semaine BA (lun→dim, ancrée BA_START) qui chevauche la fenêtre du mois ; objectif = daily × jours ouvrés (hors dimanche) réellement dans la portion visible.
function baWeeksInWindow() {
  const weeks = []; let N = 1;
  while (true) {
    const ws = new Date(BA_START.getTime() + (N - 1) * 7 * 86400000);
    if (ws > winEnd) break;
    const we = new Date(ws.getTime() + 6 * 86400000);
    const cs = ws < winStart ? winStart : ws, ce = we > winEnd ? winEnd : we;
    if (ce >= cs) {
      let wd = 0; for (let t = cs.getTime(); t <= ce.getTime(); t += 86400000) if (new Date(t).getDay() !== 0) wd++;
      weeks.push({ N, start: cs, end: new Date(ce.getTime() + 86400000 - 1), workDays: wd });
    }
    N++; if (N > 60) break;
  }
  return weeks;
}
const weekSeries = baWeeksInWindow().map(w => {
  let mont = 0;
  Object.keys(dayMap).forEach(k => { const [y, m, d] = k.split("-").map(Number); const dt = new Date(y, m - 1, d); if (dt >= w.start && dt <= w.end) mont += dayMap[k]; });
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
const TOTAL = 2;

// ════════════════════════════════════════════════════════════════════
//  SLIDE 1 · ACTIVATIONS POS — VUE MENSUELLE
// ════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide(); light(s);
  chrome(pres, s, 1, "Vue mensuelle · Performance base globale TSA");
  header(pres, s, "Activations POS — Performance par région", `Volume des points de vente (base globale TSA) · ${MONTH_LABEL} · ${POS_WINDOW_LABEL}`);

  s.addText(`Performances globales des points de vente sur ${POS_WINDOW_LABEL} (${nbCalendarDaysPOS} jours) — réalisation rapportée aux cibles mensuelles par région.`, {
    x: MX, y: 1.5, w: CW, h: 0.4, margin: 0, fontFace: fBody, fontSize: 12, color: INK2,
  });

  const rows = []; let tPos = 0, tD = 0, tM = 0, tReal = 0;
  REGIONS.forEach(reg => {
    const daily = Math.round(posDaily[reg]);
    const monthlyTarget = daily * nbCalendarDaysPOS;
    const real = volByReg[reg] || 0;
    const uniquePos = (CONFIG.pos[reg] || {}).uniquePos || 0;
    const taux = monthlyTarget ? real / monthlyTarget * 100 : 0;
    tPos += uniquePos; tD += daily; tM += monthlyTarget; tReal += real;
    rows.push([reg, fmt(uniquePos), fmt(daily), fmt(monthlyTarget), fmt(real), { text: taux.toFixed(1) + "%", color: taux >= 100 ? GREEN : taux >= 50 ? AMBER : REDX, bold: true }]);
  });
  const tt = tM ? tReal / tM * 100 : 0;
  const ty = 2.0, rowH = 0.46, tw = CW;
  table(pres, s, MX, ty, tw, ["RÉGION", "UNIQUE POS", "TARGET DAILY", "TARGET MOIS", "RÉALISATION", "TAUX"], rows, {
    colWidths: [tw * 0.18, tw * 0.13, tw * 0.18, tw * 0.18, tw * 0.21, tw * 0.12],
    rowH, fs: 10.5, hfs: 9.5,
    totalRow: ["TOTAL", fmt(tPos), fmt(tD), fmt(tM), fmt(tReal), tt.toFixed(1) + "%"],
  });

  s.addText(`Cible mensuelle totale : ${fmt(tM)} FCFA · Réalisé : ${fmt(tReal)} FCFA · Taux de réalisation global : ${tt.toFixed(1)} %.`, {
    x: MX, y: 6.55, w: CW, h: 0.35, margin: 0, fontFace: fBody, italic: true, fontSize: 10.5, color: MUTE,
  });
}

// ════════════════════════════════════════════════════════════════════
//  SLIDE 2 · ACTIVATIONS BA — VUE MENSUELLE
// ════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide(); light(s);
  chrome(pres, s, 2, "Vue mensuelle · Activations BA");
  header(pres, s, "Activations BA — Force de soutien aux TSA", `Programme BA en appui des TSA · ${MONTH_LABEL} · ${WINDOW_LABEL}`);

  const rows = []; let tEff = 0, tAct = 0, tTar = 0, tMon = 0;
  REGIONS.forEach(reg => {
    const eff = CONFIG.effectif[reg] || 0;
    const o = baAgg[reg]; const act = o.act, mont = o.mont;
    const target = baDailyByReg[reg] * nbWorkDays;
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
  s.addText(fmt(tMon), { x: mx2 + 0.14, y: ty + 0.36, w: cwid - 0.28, h: 0.62, margin: 0, fontFace: fHead, bold: true, fontSize: 24, color: INK, valign: "middle" });

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
