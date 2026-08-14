/**
 * Calcule le rapport hebdomadaire DTC et déclenche son RENDU (.pptx) sur GitHub Actions.
 * À ajouter au projet Apps Script de la chaîne (D:\LKA\DTC Pushed), lié au classeur
 * contenant les onglets CLASSEMENT_TSA et TSA_REF.
 *
 * Mise en service (une fois) :
 *   1. Script Property  GH_TOKEN = PAT GitHub (fine-grained sur lkageteam/dtc-weekly-report,
 *      « Contents: Read and write » → suffit pour repository_dispatch ; sinon classic scope `repo`).
 *   2. Exécuter installWeeklyTriggers()  → pose le filet lundi 8h/9h/10h.
 *   3. Ajouter  maybeSendReport();  à la fin de syncSheetsFromMySQL()  → envoi dès que les données du lundi arrivent.
 *
 * Fonctions utiles :
 *   • previewReport()         → logge le JSON calculé sans rien envoyer (vérif des chiffres).
 *   • buildAndTriggerReport(n)→ force l'envoi d'une semaine (ex. "1"), ignore l'état.
 *   • maybeSendReport()       → envoi auto de la dernière semaine BOUCLÉE, 1 seule fois (état LAST_SENT_WEEK).
 *   • resetLastSentWeek()     → ré-autorise le renvoi.
 *
 * Le Form BA : classeur FORM_SS_ID, onglet d'id FORM_GID (déjà renseignés ci-dessous).
 * Le payload envoyé (~5 Ko) ne contient QUE des valeurs calculées — aucun CSV, aucune donnée brute.
 */
var GH_REPO     = 'lkageteam/dtc-weekly-report';
var FORM_SS_ID  = '18n9EGmt9VFW4N7zlj41Fv-4AoVRLSrp0ah1bRZwtAc8';
var FORM_GID    = 1776404976;

// ⚙️  CHOIX MANUEL DE LA SEMAINE (pour le bouton Run / previewReport / buildAndTriggerReport sans argument)
//   • mettre un nombre (ex. 2) → force cette semaine
//   • laisser null            → prend la dernière semaine présente dans le classement
//   (N'affecte PAS l'envoi automatique maybeSendReport, qui choisit toujours par date.)
var RPT_FORCE_WEEK = 2;

var RPT_REGIONS   = ['ATLANTIQUE', 'COTONOU 1', 'COTONOU 2', 'NORD EST', 'NORD OUEST', 'SUD EST', 'SUD OUEST'];
var RPT_EFFECTIF  = { 'ATLANTIQUE': 20, 'COTONOU 1': 44, 'COTONOU 2': 31, 'NORD EST': 41, 'NORD OUEST': 24, 'SUD EST': 31, 'SUD OUEST': 30 };
var RPT_UNIQUEPOS = { 'ATLANTIQUE': 91, 'COTONOU 1': 148, 'COTONOU 2': 158, 'NORD EST': 296, 'NORD OUEST': 280, 'SUD EST': 520, 'SUD OUEST': 378 };
var RPT_BA_DAILY_GLOBAL = 1500000;     // objectif BA global / jour
var RPT_BA_WORK_DAYS    = 6;           // les BA travaillent 6 jours/semaine → cible hebdo = daily × 6
var RPT_PRIME_THRESHOLD = 450000;      // prime TSA : n°1 régional ET volume ≥ ce seuil
var RPT_PROGRAMME_START = new Date(2026, 5, 12); // 12 juin 2026 — ancre POS (ven→jeu)
var RPT_BA_START        = new Date(2026, 5, 15); // 15 juin 2026 — ancre BA (lun→dim)
var RPT_NOTE_ACTIVATIONS = "";
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
function previewReport(semaineArg) { Logger.log(JSON.stringify(buildReportData_(semaineArg), null, 2)); }

// ───────────────────── DÉCLENCHEMENT AUTOMATIQUE ─────────────────────
// À appeler (a) à la fin de syncSheetsFromMySQL()  ET  (b) via un trigger
// hebdo (installWeeklyTriggers). L'état LAST_SENT_WEEK garantit 1 envoi/semaine.
//
// Logique : on envoie la dernière semaine BA BOUCLÉE (lun→dim). Comme le BA ferme
// le dimanche et le POS le jeudi, le lundi tout est complet → targetN devient la
// semaine qui vient de finir. On n'envoie que si ses données sont dans le classement.
function maybeSendReport() {
  var props = PropertiesService.getScriptProperties();
  var lastSent = parseInt(props.getProperty('LAST_SENT_WEEK') || '0', 10);
  var targetN = Math.floor((new Date() - RPT_BA_START) / (7 * 86400000)); // 0 avant le 1er lundi
  if (targetN < 1) { Logger.log('⏳ Aucune semaine BA bouclée pour l’instant'); return; }
  if (targetN <= lastSent) { Logger.log('↩︎ Semaine ' + targetN + ' déjà envoyée'); return; }
  if (!_rptWeekPresent_(targetN)) { Logger.log('⏳ Données Semaine ' + targetN + ' pas encore dans le classement'); return; }
  buildAndTriggerReport(targetN);
  props.setProperty('LAST_SENT_WEEK', String(targetN));
  Logger.log('✅ Rapport Semaine ' + targetN + ' déclenché (auto)');
}

/** Installe le filet de sécurité : lundi 8h, 9h, 10h. À exécuter UNE fois. */
function installWeeklyTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'maybeSendReport') ScriptApp.deleteTrigger(t); });
  [8, 9, 10].forEach(function (h) {
    ScriptApp.newTrigger('maybeSendReport').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(h).create();
  });
  Logger.log('✅ Triggers installés : lundi 8h, 9h, 10h → maybeSendReport');
}

/** Réinitialise l'état (force le renvoi de la semaine en cours au prochain run). */
function resetLastSentWeek() { PropertiesService.getScriptProperties().deleteProperty('LAST_SENT_WEEK'); Logger.log('🔄 LAST_SENT_WEEK réinitialisé'); }

function _rptWeekPresent_(N) {
  var v = _rptSheetByName_(SpreadsheetApp.getActiveSpreadsheet(), 'CLASSEMENT_TSA').getDataRange().getValues();
  if (!v.length) return false;
  var ci = v[0].map(function (x) { return String(x).trim(); }).indexOf('SEMAINE');
  if (ci < 0) return false;
  var target = 'Semaine ' + N;
  for (var i = 1; i < v.length; i++) if (String(v[i][ci]).trim() === target) return true;
  return false;
}

// ───────────────────────────── calcul ─────────────────────────────
function buildReportData_(semaineArg) {
  var bound = SpreadsheetApp.getActiveSpreadsheet(); // classeur de la chaîne (CLASSEMENT_TSA + TSA_REF)
  var classement = _rptRows_(_rptSheetByName_(bound, 'CLASSEMENT_TSA'));
  var tsaRef     = _rptRows_(_rptSheetByName_(bound, 'TSA_REF'));

  // semaine cible : argument si présent, sinon constante RPT_FORCE_WEEK (si définie), sinon dernière SEMAINE du classement
  var weeks = {};
  classement.forEach(function (r) { var w = String(r.SEMAINE || '').trim(); if (/semaine/i.test(w)) weeks[w] = true; });
  var weekList = Object.keys(weeks).sort(function (a, b) { return (parseInt(b.replace(/\D/g, '')) || 0) - (parseInt(a.replace(/\D/g, '')) || 0); });
  var targetWeek = semaineArg || RPT_FORCE_WEEK;
  var semaine = targetWeek ? (/^\d+$/.test(String(targetWeek)) ? 'Semaine ' + targetWeek : String(targetWeek)) : (weekList[0] || 'Semaine 1');
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
  var kNum = _rptKey_(hk, /num.*ro.*du.*ba|numero_ba/i); // identifiant stable (téléphone) pour compter les BA actifs
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

  // WEEK ON WEEK : réalisé par semaine 1..N (POS depuis classement, BA depuis form, fenêtres lun→dim)
  var N = parseInt(String(semaine).replace(/\D/g, ''), 10) || 1;
  var wow = { weeks: [], periodes: [], pos: [], ba: [], posObjectif: objectifTSA, baObjectif: RPT_BA_DAILY_GLOBAL * RPT_BA_WORK_DAYS };
  // Diagnostic BA : volume (nb transactions), valeur (montant) et #personnes actives (numéros distincts) par semaine
  var baDiag = { weeks: [], nbTx: [], montant: [], actifs: [] };
  for (var k = 1; k <= N; k++) {
    wow.weeks.push('S' + k);
    var wk = 'Semaine ' + k, vp = 0;
    classement.forEach(function (r) { if (String(r.SEMAINE).trim() === wk) vp += _rptNum_(r.VOLUME_XAF); });
    wow.pos.push(vp);
    var ws = new Date(RPT_BA_START.getTime() + (k - 1) * 7 * 86400000), we = new Date(ws.getTime() + 7 * 86400000 - 1), vb = 0, nbTx = 0;
    // Libellé de PÉRIODE affiché sur l'axe du graphique WoW à la place de « S8 ».
    // `we` porte 23:59:59.999 du dimanche : getDate() rend bien le jour de clôture.
    wow.periodes.push(_rptLibellePeriode_(ws, we));
    var actifsSet = {};
    formRows.forEach(function (r) {
      var reg = String(r[kReg] || '').trim().toUpperCase(); if (!RPT_EFFECTIF[reg]) return;
      var d = r[kTs]; if (!(d instanceof Date)) d = _rptDate_(d); if (!d || d < ws || d > we) return;
      vb += _rptNum_(r[kMont]); nbTx++;
      var numBA = String(r[kNum] || '').trim().toUpperCase(); if (numBA) actifsSet[numBA] = true;
    });
    wow.ba.push(vb);
    baDiag.weeks.push('S' + k); baDiag.nbTx.push(nbTx); baDiag.montant.push(vb); baDiag.actifs.push(Object.keys(actifsSet).length);
  }

  // RETOURS TERRAIN : répartition cumulée des problèmes signalés (toutes semaines)
  var kProb = _rptKey_(hk, /probl.*principal|principal.*rencontr/i);
  var retours = _rptRetours_(formRows, kProb);

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
    wow: wow,
    retours: retours,
    baDiag: baDiag,
    notes: { activations: RPT_NOTE_ACTIVATIONS, cloture: RPT_NOTE_CLOTURE },
  };
}

// ─── Retours Terrain : classification des problèmes (Form) ───────────
var RPT_RETOUR_LABELS = [
  "Lenteur du message flash / prompt / validation",
  "Le prompt de validation ne vient pas du tout",
  "Compte débité mais le prompt n'arrive pas chez le client",
  "Erreur de code PIN lors de la liaison/validation",
  "Client/abonné ne se présente pas avec son téléphone (ou refuse de le donner / de confirmer)",
  "Problème avec la SIM rattachée (ne fonctionne pas / échoue)",
  "Problème avec la SIM marchand (permutation, non éligible)",
  "Le numéro du client/abonné n'apparaît pas dans le message/SMS retour côté POS",
  "Demande relative au versement des commissions sur le compte commission",
];
var RPT_RETOUR_AUTRES = "Divers / Autres";
function _rptNormStr_(s) { return String(s == null ? '' : s).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').trim(); }
function _rptClassifyRetour_(raw) {
  var t = _rptNormStr_(raw);
  if (!t) return null;
  if (/^(non|aucun|ras|r\.a\.s|rien|pas de probleme|aucun probleme|n\/?a)\b/.test(t)) return null;
  for (var i = 0; i < RPT_RETOUR_LABELS.length; i++) if (_rptNormStr_(RPT_RETOUR_LABELS[i]) === t) return RPT_RETOUR_LABELS[i];
  if (/commission|versement/.test(t)) return RPT_RETOUR_LABELS[8];
  if (/marchand|permut|eligib/.test(t)) return RPT_RETOUR_LABELS[6];
  if (/sim/.test(t)) return RPT_RETOUR_LABELS[5];
  if (/\bpin\b|code pin/.test(t)) return RPT_RETOUR_LABELS[3];
  if (/numero.*(apparait|apparait pas)|sms.*retour|message.*retour|retour.*pos/.test(t)) return RPT_RETOUR_LABELS[7];
  if (/debit|defalqu|compte.*(debit|defalqu)|debite.*prompt/.test(t)) return RPT_RETOUR_LABELS[2];
  if (/(prompt|validation).*(ne vient|n.?arrive|pas du tout)|ne vient pas/.test(t)) return RPT_RETOUR_LABELS[1];
  if (/telephone|presente|refuse|confirmer|papier|mefian/.test(t)) return RPT_RETOUR_LABELS[4];
  if (/lenteur|lent|flash|prompt|validation/.test(t)) return RPT_RETOUR_LABELS[0];
  return RPT_RETOUR_AUTRES;
}
function _rptRetours_(formRows, kProb) {
  if (!kProb) return { total: 0, items: [] };
  var counts = {}, total = 0;
  formRows.forEach(function (r) {
    var lbl = _rptClassifyRetour_(r[kProb]); if (!lbl) return;
    counts[lbl] = (counts[lbl] || 0) + 1; total++;
  });
  var items = Object.keys(counts).map(function (label) {
    return { label: label, count: counts[label], pct: total ? Math.round(counts[label] / total * 1000) / 10 : 0 };
  }).sort(function (a, b) { return b.count - a.count; });
  return { total: total, items: items };
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
// Mois abrégés — ce libellé sert d'étiquette d'axe et doit tenir sur UNE ligne
// horizontale dans une colonne de graphique. Doit rester identique à `MOIS_AB` de
// `scripts/progress_report_dtc_assisted.js` et de `pipelines/dtc_weekly/dtc_weekly.py`
// (lka-unified) : les trois producteurs alimentent le MÊME graphique.
var RPT_MOIS_AB = ['janv', 'févr', 'mars', 'avr', 'mai', 'juin', 'juil', 'août', 'sept', 'oct', 'nov', 'déc'];

// « 3 – 9 août », ou « 31 juil – 6 août » quand la semaine chevauche deux mois.
// Le mois n'est répété que s'il change : l'écrire deux fois alourdit sans rien apprendre.
function _rptLibellePeriode_(debut, fin) {
  var m1 = RPT_MOIS_AB[debut.getMonth()], m2 = RPT_MOIS_AB[fin.getMonth()];
  return m1 === m2
    ? debut.getDate() + ' – ' + fin.getDate() + ' ' + m1
    : debut.getDate() + ' ' + m1 + ' – ' + fin.getDate() + ' ' + m2;
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
