/**
 * Calcule le rapport hebdomadaire DTC et déclenche son RENDU (.pptx) sur GitHub Actions.
 * À ajouter au projet Apps Script de la chaîne (D:\LKA\DTC Pushed), lié au classeur
 * contenant les onglets CLASSEMENT_TSA et TSA_REF.
 *
 * → Appeler buildAndTriggerReport()  à la fin de syncSheetsFromMySQL()  (ou trigger hebdo).
 *
 * Pré-requis (une fois) :
 *   - Script Property  GH_TOKEN = PAT GitHub (fine-grained sur lkageteam/dtc-weekly-report,
 *     permission « Contents: Read and write » → suffit pour repository_dispatch ; sinon classic scope `repo`).
 *   - Le Form BA : classeur FORM_SS_ID, onglet d'id FORM_GID (déjà renseignés ci-dessous).
 *
 * Le payload envoyé (~5 Ko) contient UNIQUEMENT des valeurs déjà calculées. Aucune donnée brute,
 * aucun CSV : GitHub Actions ne fait QUE dessiner le .pptx.
 */
var GH_REPO     = 'lkageteam/dtc-weekly-report';
var FORM_SS_ID  = '18n9EGmt9VFW4N7zlj41Fv-4AoVRLSrp0ah1bRZwtAc8';
var FORM_GID    = 1776404976;

var RPT_REGIONS   = ['ATLANTIQUE', 'COTONOU 1', 'COTONOU 2', 'NORD EST', 'NORD OUEST', 'SUD EST', 'SUD OUEST'];
var RPT_EFFECTIF  = { 'ATLANTIQUE': 20, 'COTONOU 1': 44, 'COTONOU 2': 31, 'NORD EST': 41, 'NORD OUEST': 24, 'SUD EST': 31, 'SUD OUEST': 30 };
var RPT_UNIQUEPOS = { 'ATLANTIQUE': 91, 'COTONOU 1': 148, 'COTONOU 2': 158, 'NORD EST': 296, 'NORD OUEST': 280, 'SUD EST': 520, 'SUD OUEST': 378 };
var RPT_BA_DAILY_GLOBAL = 1500000;     // objectif BA global / jour
var RPT_BA_WORK_DAYS    = 6;           // les BA travaillent 6 jours/semaine → cible hebdo = daily × 6
var RPT_PRIME_THRESHOLD = 450000;      // prime TSA : n°1 régional ET volume ≥ ce seuil
var RPT_PROGRAMME_START = new Date(2026, 5, 12); // 12 juin 2026 — ancre POS (ven→jeu)
var RPT_BA_START        = new Date(2026, 5, 15); // 15 juin 2026 — ancre BA (lun→dim)
var RPT_NOTE_ACTIVATIONS = "Certains BA recrutés n'ont pas été activés immédiatement en raison de l'absence du dress code (non fourni par MTN) ; l'agence a eu recours à son fournisseur pour pallier ce challenge.";
var RPT_NOTE_CLOTURE     = "Maintenir le rythme jour-par-jour · Soutenir le week-end · Convertir le Top des POS";

function buildAndTriggerReport(semaineArg) {
  var report = buildReportData_(semaineArg);
  var token = PropertiesService.getScriptProperties().getProperty('GH_TOKEN');
  if (!token) throw new Error('GH_TOKEN manquant dans les Script Properties');
  var res = UrlFetchApp.fetch('https://api.github.com/repos/' + GH_REPO + '/dispatches', {
    method: 'post', contentType: 'application/json',
    headers: { Authorization: 'token ' + token, Accept: 'application/vnd.github+json', 'X-GitHub-Api-Version': '2022-11-28' },
    payload: JSON.stringify({ event_type: 'weekly-report', client_payload: { report: report } }),
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() >= 300) throw new Error('dispatch ' + res.getResponseCode() + ' ' + res.getContentText());
  Logger.log('✅ Rapport ' + report.semaine + ' déclenché sur GitHub');
}

/** Aperçu local du JSON (sans déclencher) — pratique pour vérifier les chiffres. */
function previewReport() { Logger.log(JSON.stringify(buildReportData_(), null, 2)); }

// ───────────────────────────── calcul ─────────────────────────────
function buildReportData_(semaineArg) {
  var bound = SpreadsheetApp.getActiveSpreadsheet(); // classeur de la chaîne (CLASSEMENT_TSA + TSA_REF)
  var classement = _rptRows_(_rptSheetByName_(bound, 'CLASSEMENT_TSA'));
  var tsaRef     = _rptRows_(_rptSheetByName_(bound, 'TSA_REF'));

  // semaine cible : argument, sinon dernière SEMAINE du classement
  var weeks = {};
  classement.forEach(function (r) { var w = String(r.SEMAINE || '').trim(); if (/semaine/i.test(w)) weeks[w] = true; });
  var weekList = Object.keys(weeks).sort(function (a, b) { return (parseInt(b.replace(/\D/g, '')) || 0) - (parseInt(a.replace(/\D/g, '')) || 0); });
  var semaine = semaineArg ? (/^\d+$/.test(String(semaineArg)) ? 'Semaine ' + semaineArg : String(semaineArg)) : (weekList[0] || 'Semaine 1');
  var cy = _rptCycle_(semaine);
  var cw = classement.filter(function (r) { return String(r.SEMAINE).trim() === semaine; });

  // POS : cibles journalières par région (somme OBJECTIF_DAILY de TSA_REF)
  var posDaily = {}; RPT_REGIONS.forEach(function (r) { posDaily[r] = 0; });
  tsaRef.forEach(function (r) { var k = String(r.REGION || '').trim().toUpperCase(); if (posDaily[k] !== undefined) posDaily[k] += _rptNum_(r.OBJECTIF_DAILY); });
  var objectifTSA = 0; RPT_REGIONS.forEach(function (r) { objectifTSA += posDaily[r]; }); objectifTSA *= 7;

  // volume réalisé par région (classement de la semaine)
  var volByReg = {}; RPT_REGIONS.forEach(function (r) { volByReg[r] = 0; });
  cw.forEach(function (r) { var k = String(r.REGION).trim().toUpperCase(); if (volByReg[k] !== undefined) volByReg[k] += _rptNum_(r.VOLUME_XAF); });

  // prime : n°1 régional (RANG_REGIONAL = 1) par région, trié par volume décroissant
  var prime = RPT_REGIONS.map(function (reg) {
    var rows = cw.filter(function (r) { return String(r.REGION).trim().toUpperCase() === reg; });
    var top = null;
    for (var i = 0; i < rows.length; i++) { if (_rptNum_(rows[i].RANG_REGIONAL) === 1) { top = rows[i]; break; } }
    if (!top) { rows.sort(function (a, b) { return _rptNum_(b.VOLUME_XAF) - _rptNum_(a.VOLUME_XAF); }); top = rows[0]; }
    return top ? { region: reg, tsa: String(top.TSA || ''), rbm: String(top.RBM || ''), volume: _rptNum_(top.VOLUME_XAF), posActifs: _rptNum_(top.NB_POS_ACTIFS), tauxAct: _rptNum_(top.TAUX_ACTIVATION) } : null;
  }).filter(function (x) { return x; }).sort(function (a, b) { return b.volume - a.volume; });

  // BA : activations + montant par région + série jour-par-jour (Form, filtré par dates)
  var formRows = _rptRows_(_rptSheetByGid_(SpreadsheetApp.openById(FORM_SS_ID), FORM_GID));
  var hk = formRows.length ? formRows[0] : {};
  var kReg = _rptKey_(hk, /Region/i), kMont = _rptKey_(hk, /Montant/i), kTs = _rptKey_(hk, /Timestamp|Horodat/i);
  var totalEff = 0; RPT_REGIONS.forEach(function (r) { totalEff += RPT_EFFECTIF[r]; });
  var winEnd = new Date(cy.baEnd.getTime() + 86400000 - 1); // inclut tout le dernier jour
  var baAgg = {}; RPT_REGIONS.forEach(function (r) { baAgg[r] = { act: 0, mont: 0 }; });
  var dayMap = {}, JOURS = ['Dim', 'Lun', 'Mar', 'Mer', 'Jeu', 'Ven', 'Sam'];
  formRows.forEach(function (r) {
    var reg = String(r[kReg] || '').trim().toUpperCase(); if (!baAgg[reg]) return;
    var d = r[kTs]; if (!(d instanceof Date)) d = _rptDate_(d); if (!d || d < cy.baStart || d > winEnd) return;
    var m = _rptNum_(r[kMont]);
    baAgg[reg].act++; baAgg[reg].mont += m;
    var key = d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
    dayMap[key] = (dayMap[key] || 0) + m;
  });
  var days = Object.keys(dayMap).sort().map(function (k) {
    var p = k.split('-'), dd = new Date(+p[0], +p[1] - 1, +p[2]);
    return { label: JOURS[dd.getDay()] + ' ' + dd.getDate(), mont: dayMap[k], pct: Math.round(dayMap[k] / RPT_BA_DAILY_GLOBAL * 1000) / 10 };
  });

  // assemblage
  var pos = {}, baRows = {}, recrutement = {};
  RPT_REGIONS.forEach(function (r) {
    pos[r] = { uniquePos: RPT_UNIQUEPOS[r], daily: Math.round(posDaily[r]), weekly: Math.round(posDaily[r] * 7), real: volByReg[r] };
    baRows[r] = { effectif: RPT_EFFECTIF[r], activations: baAgg[r].act, montant: baAgg[r].mont, dailyTarget: Math.round(RPT_EFFECTIF[r] / totalEff * RPT_BA_DAILY_GLOBAL) };
    recrutement[r] = [RPT_EFFECTIF[r], RPT_EFFECTIF[r]];
  });
  return {
    semaine: semaine, label: cy.label, weekNo: semaine, reportDate: cy.reportDate, posPeriod: cy.pos,
    recrutement: recrutement, primeThreshold: RPT_PRIME_THRESHOLD, objectifTSA: objectifTSA, pos: pos, prime: prime,
    ba: { workDays: RPT_BA_WORK_DAYS, dailyObjectif: RPT_BA_DAILY_GLOBAL, rows: baRows, days: days },
    notes: { activations: RPT_NOTE_ACTIVATIONS, cloture: RPT_NOTE_CLOTURE },
  };
}

// ───────────────────────────── helpers ─────────────────────────────
function _rptCycle_(semaine) {
  var N = parseInt(String(semaine).replace(/\D/g, ''), 10) || 1, MS = 86400000;
  var bs = new Date(RPT_BA_START.getTime() + (N - 1) * 7 * MS), be = new Date(bs.getTime() + 6 * MS);        // BA lun→dim
  var ps = new Date(RPT_PROGRAMME_START.getTime() + (N - 1) * 7 * MS), pe = new Date(ps.getTime() + 6 * MS); // POS ven→jeu
  var MOIS = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre'];
  function human(d) { return d.getDate() + ' ' + MOIS[d.getMonth()] + ' ' + d.getFullYear(); }
  function dmy(d) { return ('0' + d.getDate()).slice(-2) + '/' + ('0' + (d.getMonth() + 1)).slice(-2) + '/' + d.getFullYear(); }
  return { label: 'Semaine ' + N + ' · ' + human(bs) + ' au ' + human(be), reportDate: dmy(new Date(be.getTime() + MS)), baStart: bs, baEnd: be, pos: human(ps) + ' au ' + human(pe) };
}
function _rptSheetByName_(ss, name) { var s = ss.getSheetByName(name); if (!s) throw new Error('Onglet "' + name + '" introuvable'); return s; }
function _rptSheetByGid_(ss, gid) { var sh = ss.getSheets(); for (var i = 0; i < sh.length; i++) if (sh[i].getSheetId() === gid) return sh[i]; throw new Error('Onglet gid ' + gid + ' introuvable'); }
function _rptRows_(sheet) {
  var v = sheet.getDataRange().getValues(); if (v.length < 2) return [];
  var H = v[0].map(function (h) { return String(h).trim(); });
  return v.slice(1).map(function (r) { var o = {}; for (var i = 0; i < H.length; i++) o[H[i]] = r[i]; return o; });
}
function _rptNum_(x) { if (x === null || x === undefined || x === '') return 0; var n = parseFloat(String(x).replace(/[^0-9.\-]/g, '')); return isNaN(n) ? 0 : n; }
function _rptKey_(obj, re) { for (var k in obj) if (re.test(k)) return k; return null; }
function _rptDate_(s) { if (!s) return null; var m = String(s).match(/(\d+)\/(\d+)\/(\d+)\s+(\d+):(\d+):(\d+)/); if (m) return new Date(+m[3], +m[1] - 1, +m[2], +m[4], +m[5], +m[6]); var d = new Date(s); return isNaN(d) ? null : d; }
