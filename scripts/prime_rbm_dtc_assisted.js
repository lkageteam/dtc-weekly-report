// ════════════════════════════════════════════════════════════════════
//  PRIME RBM · DTC ASSISTED — deck autonome, une slide par mois PLEIN.
//
//  Lancement :
//    INPUT_DIR=./inputs_live node scripts/prime_rbm_dtc_assisted.js 2026-07 2026-08
//    (sans argument → 2026-07 et 2026-08)
//  → ./outputs/Prime_RBM_DTC_Assisted_LKA_<mois>.pptx
//
//  Les mêmes slides sont insérées dans le rapport mensuel
//  (monthly_report_dtc_assisted.js) : le calcul vit dans lib/prime_rbm.js
//  et n'est écrit qu'UNE fois.
// ════════════════════════════════════════════════════════════════════
const pptxgen = require("pptxgenjs");
const { withPPTXEmbedFonts } = require("pptx-embed-fonts/pptxgenjs");
const fs = require("fs");
const path = require("path");
const T = require("./lib/theme");
const { INK, INK2, YEL, MUTE, MUTE2, fHead, fBody, W, H, MX, CW, HAS_LOGO, LOGO, MOIS, fmt, toTitle, light, embedFonts } = T;
const { classementRBM, slidePrimeRBM } = require("./lib/prime_rbm");

const IN = process.env.INPUT_DIR ? path.resolve(process.env.INPUT_DIR) : path.join(__dirname, "..", "inputs");
const OUT = path.join(__dirname, "..", "outputs");

const args = process.argv.slice(2).filter(a => /^\d{4}-\d{2}$/.test(a));
const MOISCIBLES = args.length ? args : ["2026-07", "2026-08"];

(async () => {
  const EnhancedPPTXGenJS = withPPTXEmbedFonts(pptxgen);
  const pres = new EnhancedPPTXGenJS();
  pres.layout = "LAYOUT_WIDE";
  pres.title = "Prime RBM DTC Assisted — MTN Bénin";
  pres.author = "Contribution LKA";
  await embedFonts(pres);

  const mois = MOISCIBLES.map(ym => {
    const [yy, mm] = ym.split("-").map(Number);
    return classementRBM(IN, yy, mm);
  });

  // ── Couverture ──
  {
    const s = pres.addSlide(); light(s);
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 0.09, fill: { color: YEL }, line: { type: "none" } });
    s.addShape(pres.shapes.RECTANGLE, { x: MX, y: 2.2, w: 0.24, h: 1.5, fill: { color: YEL }, line: { type: "none" } });
    s.addText("PRIME RBM", { x: MX + 0.5, y: 2.15, w: CW, h: 0.9, margin: 0, fontFace: fHead, bold: true, fontSize: 46, color: INK });
    s.addText("DTC ASSISTED · MTN BÉNIN × LKA", { x: MX + 0.5, y: 3.02, w: CW, h: 0.4, margin: 0, fontFace: fBody, fontSize: 14, color: MUTE, charSpacing: 3 });
    s.addText(
      mois.map(m => `${m.label} — ${toTitle(m.lignes[0].rbm)} (${m.lignes[0].region}) · ${fmt(m.lignes[0].volume)} XAF`).join("\n"),
      { x: MX + 0.5, y: 3.6, w: CW - 0.5, h: 1.0, margin: 0, fontFace: fBody, fontSize: 13, color: INK2, lineSpacingMultiple: 1.4 });
    s.addText("Critère : volume POS du mois · un seul primé sur les 7 · mois pleins uniquement (juin exclu)", {
      x: MX + 0.5, y: 4.75, w: CW - 0.5, h: 0.3, margin: 0, fontFace: fBody, fontSize: 10, color: MUTE2 });
    if (HAS_LOGO) s.addImage({ path: LOGO, x: W - MX - 1.0, y: H - 0.9, w: 0.72, h: 0.66 });
  }

  mois.forEach((m, i) => slidePrimeRBM(pres, m, i + 2));

  fs.mkdirSync(OUT, { recursive: true });
  const suffixe = MOISCIBLES.map(x => x.replace("-", "")).join("_");
  const file = path.join(OUT, `Prime_RBM_DTC_Assisted_LKA_${suffixe}.pptx`);
  await pres.writeFile({ fileName: file });
  mois.forEach(m => console.log(`🏆 ${m.label} : ${m.lignes[0].rbm} (${m.lignes[0].region}) — ${fmt(m.lignes[0].volume)} XAF · ${m.lignes[0].taux.toFixed(1)}% · ${m.jours} jours`));
  console.log(`✅ ${file}`);
})();
