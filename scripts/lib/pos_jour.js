// ════════════════════════════════════════════════════════════════════
//  POS JOURNALIER — la seule source correcte d'un agrégat MENSUEL.
//
//  🔴 POURQUOI CE FICHIER EXISTE — MESURÉ LE 2026-09-03, PAS SUPPOSÉ.
//  `classement.csv` (vue `v_classement_tsa_semaine`) ne porte PAS de date :
//  chaque ligne porte une SEMAINE POS (ven→jeu, ancrée au 12/06). Agréger un
//  mois en retenant « les semaines qui chevauchent le mois » fait donc déborder
//  la fenêtre des deux côtés :
//
//     juillet 2026, Nord Ouest
//       par semaines (S3→S8, fenêtre RÉELLE 26/06 → 06/08, 42 jours) : 25 910 800
//       par jours    (01/07 → 31/07, 31 jours)                        : 19 557 450
//
//  Soit +32 % de volume attribué à juillet, dont une partie appartient à juin et
//  une partie à août. Le classement des 7 RBM en est changé : par semaines le
//  Nord Est passait 2ᵉ de juillet avec 20,7 M, par jours il est 2ᵉ avec 15,4 M —
//  et les rangs 3/4 (Sud Est / Sud Ouest) s'inversent.
//
//  Le jour EXISTE dans la base (`lka_client_mtn.pos_sales`, une ligne par POS et
//  par jour, 156 960 lignes du 12/06 au 30/08). On l'exploite. La semaine ne sert
//  plus qu'au rapport HEBDOMADAIRE, dont la fenêtre est justement une semaine.
//
//  ⚠️ `pos_sales` peut s'arrêter avant la fin du mois (au 03/09 elle s'arrête au
//  30/08 : le 31/08 n'est pas encore chargé). `jours` rend le nombre de jours
//  RÉELLEMENT présents — c'est lui qui doit multiplier la cible journalière, pas
//  le nombre de jours du calendrier. Sinon on compare 30 jours de réalisé à 31
//  jours de cible et on invente un retard.
// ════════════════════════════════════════════════════════════════════
const path = require("path");
const T = require("./theme");
const { REGIONS, readCSV, num } = T;

let _cache = null, _cacheDir = null;

function chargerPosDaily(dir) {
  if (_cache && _cacheDir === dir) return _cache;
  const brut = readCSV(path.join(dir, "pos_daily.csv"));
  _cache = brut.map(r => {
    const [y, m, d] = String(r.SALE_DATE).split("-").map(Number);
    return {
      d: new Date(y, m - 1, d),
      pos: String(r.POS_MSISDN || "").trim(),
      tsa: String(r.TSA || "").trim(),
      tsaNom: r.TSA_FULL_NAME || r.TSA,
      reg: (r.REGION || "").trim().toUpperCase(),
      rbm: (r.RBM || "").trim(),
      v: num(r.VOLUME_XAF),
    };
  }).filter(r => r.d && REGIONS.includes(r.reg));
  _cacheDir = dir;
  return _cache;
}

// Agrégats POS d'une fenêtre [debut, fin] (bornes incluses), calculés au JOUR.
function posSurFenetre(dir, debut, fin) {
  const lignes = chargerPosDaily(dir);
  const volByReg = {}, posByReg = {};
  REGIONS.forEach(r => { volByReg[r] = 0; posByReg[r] = new Set(); });
  const volByTSA = new Map();
  const jours = new Set();
  lignes.forEach(r => {
    if (r.d < debut || r.d > fin) return;
    jours.add(r.d.getTime());
    volByReg[r.reg] += r.v;
    if (r.v > 0) posByReg[r.reg].add(r.pos);
    const cur = volByTSA.get(r.tsa) || { TSA: r.tsaNom, CORPORATE: r.tsa, REGION: r.reg, RBM: r.rbm, VOLUME_XAF: 0, POS: new Set() };
    cur.VOLUME_XAF += r.v; if (r.v > 0) cur.POS.add(r.pos);
    volByTSA.set(r.tsa, cur);
  });
  volByTSA.forEach(t => { t.NB_POS_ACTIFS = t.POS.size; delete t.POS; });
  const posActifs = {}; REGIONS.forEach(r => posActifs[r] = posByReg[r].size);
  const joursTries = [...jours].sort((a, b) => a - b);
  return {
    volByReg, posActifs, volByTSA,
    volTotal: REGIONS.reduce((s, r) => s + volByReg[r], 0),
    jours: joursTries.length,
    premier: joursTries.length ? new Date(joursTries[0]) : null,
    dernier: joursTries.length ? new Date(joursTries[joursTries.length - 1]) : null,
  };
}

// Agrégats POS d'un mois calendaire, bornés au démarrage du programme.
function posDuMois(dir, yy, mm, posStart) {
  const first = new Date(yy, mm - 1, 1);
  const debut = posStart && first < posStart ? posStart : first;
  const fin = new Date(yy, mm, 0);
  return Object.assign(posSurFenetre(dir, debut, fin), { debut, fin });
}

module.exports = { chargerPosDaily, posSurFenetre, posDuMois };
