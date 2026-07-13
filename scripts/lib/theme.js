// ════════════════════════════════════════════════════════════════════
//  THEME & HELPERS PARTAGÉS — palette MTN, dimensions, CSV, table/chrome.
//  Utilisé par les scripts de rendu (hebdomadaire, mensuel, ...).
//  Ne PAS dupliquer ces constantes dans un nouveau script : require() ce fichier.
// ════════════════════════════════════════════════════════════════════
const fs = require("fs");
const path = require("path");

// ── PALETTE (light premium MTN) ─────────────────────────────────────
const BG = "FFFFFF", INK = "1A1A1A", INK2 = "2B2B2B";
const YEL = "FFCC00", YELDK = "E0B400", YELSOFT = "FFF3C4";
const MUTE = "6B6B6B", MUTE2 = "9A9A9A", LINE = "E4E4E4", ROW = "F4F4F4";
const GREEN = "2E9E4F", AMBER = "E8932B", REDX = "D1453B", BLUE = "1F6FB2";

// ── FONTS ───────────────────────────────────────────────────────────
const fHead = "MTN Brighter Sans Bold";
const fBody = "MTN Brighter Sans";

const ROOT = path.join(__dirname, "..", "..");
const LOGO = path.join(ROOT, "assets", "lka_logo.png");
const HAS_LOGO = fs.existsSync(LOGO);

const W = 13.33, H = 7.5, MX = 0.7, CW = W - 2 * MX;
const REGIONS = ["ATLANTIQUE", "COTONOU 1", "COTONOU 2", "NORD EST", "NORD OUEST", "SUD EST", "SUD OUEST"];

// — Valeurs structurelles (stables — mêmes que le rapport hebdomadaire) —
const CONFIG = {
  effectif: { "ATLANTIQUE": 20, "COTONOU 1": 44, "COTONOU 2": 31, "NORD EST": 41, "NORD OUEST": 24, "SUD EST": 31, "SUD OUEST": 30 },
  baDailyGlobal: 1500000,   // objectif BA global / JOUR, réparti par % d'effectif régional
  pos: {
    "ATLANTIQUE": { uniquePos:  91 },
    "COTONOU 1":  { uniquePos: 148 },
    "COTONOU 2":  { uniquePos: 158 },
    "NORD EST":   { uniquePos: 296 },
    "NORD OUEST": { uniquePos: 280 },
    "SUD EST":    { uniquePos: 520 },
    "SUD OUEST":  { uniquePos: 378 },
  },
  primeThreshold: 450000,
};

const PROGRAMME_START = new Date(2026, 5, 12); // 12 juin 2026 — ancre POS (ven→jeu)
const BA_START        = new Date(2026, 5, 15); // 15 juin 2026 — ancre BA (lun→dim), = début réel du programme

// ════════════════════════════════════════════════════════════════════
//  CSV PARSING
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
const fmt = n => Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, " ");
function parseTS(s) {
  const m = String(s).match(/(\d+)\/(\d+)\/(\d+)\s+(\d+):(\d+):(\d+)/);
  return m ? new Date(+m[3], +m[1] - 1, +m[2], +m[4], +m[5], +m[6]) : null;
}
const MOIS = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre"];
const JOURS = ["Dim", "Lun", "Mar", "Mer", "Jeu", "Ven", "Sam"];
const dayLabel = k => { const [y, m, d] = k.split("-").map(Number); return JOURS[new Date(y, m - 1, d).getDay()] + " " + d; };
const toTitle = s => String(s).toLowerCase().replace(/\b([a-zàâäéèêëïîôöùûüç])/g, c => c.toUpperCase()).replace(/\b-([a-z])/g, (m, c) => "-" + c.toUpperCase());

// ════════════════════════════════════════════════════════════════════
//  PPTX HELPERS — nécessitent l'instance `pres` du script appelant
// ════════════════════════════════════════════════════════════════════
function light(slide) { slide.background = { color: BG }; }

function chrome(pres, slide, page, section) {
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

function header(pres, slide, title, accent) {
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
function table(pres, slide, x, y, w, headers, rows, opts = {}) {
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

async function embedFonts(pres) {
  try {
    const reg = fs.readFileSync(path.join(ROOT, "fonts", "MTNBrighterSans-Regular.ttf"));
    const bold = fs.readFileSync(path.join(ROOT, "fonts", "MTNBrighterSans-Bold.ttf"));
    await pres.addFont({ fontFace: fBody, fontFile: reg, fontType: "ttf" });
    await pres.addFont({ fontFace: fHead, fontFile: bold, fontType: "ttf" });
    console.log("✅ Fonts embedded");
  } catch (e) { console.warn("⚠️ Font embed failed:", e.message); }
}

module.exports = {
  BG, INK, INK2, YEL, YELDK, YELSOFT, MUTE, MUTE2, LINE, ROW, GREEN, AMBER, REDX, BLUE,
  fHead, fBody, LOGO, HAS_LOGO, W, H, MX, CW, REGIONS, CONFIG, PROGRAMME_START, BA_START,
  parseCSV, readCSV, num, fmt, parseTS, MOIS, JOURS, dayLabel, toTitle,
  light, chrome, header, table, embedFonts,
};
