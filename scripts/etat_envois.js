// ════════════════════════════════════════════════════════════════════
//  ÉTAT DES ENVOIS — la clé d'idempotence, au POINT DE LIVRAISON
//
//  Deux producteurs peuvent déclencher ce dépôt et ne se voient pas :
//   • `appscript/compute_and_trigger.gs` — se dédoublonne avec LAST_SENT_WEEK
//     (Script Properties du projet Apps Script) ;
//   • `lka-unified/pipelines/dtc_weekly/scripts/verifier_et_lancer.py` — se
//     dédoublonne avec `lka_client_mtn.dtc_weekly_envois` (MySQL).
//  Chacun est correct chez lui, et ensemble ils enverraient le rapport DEUX
//  FOIS le lundi (le filet du tronc part à 07:00 locale, les triggers Apps
//  Script à 8h/9h/10h). Le garde-fou ne peut donc pas vivre chez un producteur :
//  il vit ICI, là où les deux passent.
//
//  ⚠️ LA CLÉ EST UNE PÉRIODE (« Semaine 4 »), JAMAIS UN HORODATAGE. Une clé à
//  la minute n'entre en collision avec rien : elle donnerait l'illusion d'un
//  dédoublonnage sans en être un.
//
//  Usage :
//    node scripts/etat_envois.js check <semaine>   → exit 0 si JAMAIS envoyée,
//                                                     exit 1 si déjà envoyée
//    node scripts/etat_envois.js mark  <semaine>   → inscrit l'envoi
// ════════════════════════════════════════════════════════════════════
const fs = require("fs");
const path = require("path");

const FICHIER = path.join(__dirname, "..", "state", "envois_hebdo.json");

function lire() {
  try { return JSON.parse(fs.readFileSync(FICHIER, "utf8")); } catch { return []; }
}

function ecrire(envois) {
  fs.mkdirSync(path.dirname(FICHIER), { recursive: true });
  fs.writeFileSync(FICHIER, JSON.stringify(envois, null, 2) + "\n", "utf8");
}

const [, , commande, ...reste] = process.argv;
const semaine = reste.join(" ").trim();

if (!commande || !semaine) {
  console.error("usage : node scripts/etat_envois.js <check|mark> <semaine>");
  process.exit(2);
}

const envois = lire();
const deja = envois.find(e => e.semaine === semaine);

if (commande === "check") {
  if (deja) {
    console.log(`↩︎ « ${semaine} » déjà envoyée le ${deja.envoye_le} (${deja.run || "run inconnu"}) — pas de second mail.`);
    process.exit(1);
  }
  console.log(`✓ « ${semaine} » jamais envoyée — envoi autorisé.`);
  process.exit(0);
}

if (commande === "mark") {
  if (deja) { console.log(`↩︎ « ${semaine} » déjà inscrite — rien à faire.`); process.exit(0); }
  const run = process.env.GITHUB_SERVER_URL && process.env.GITHUB_REPOSITORY && process.env.GITHUB_RUN_ID
    ? `${process.env.GITHUB_SERVER_URL}/${process.env.GITHUB_REPOSITORY}/actions/runs/${process.env.GITHUB_RUN_ID}`
    : "";
  envois.push({
    semaine,
    envoye_le: new Date().toISOString(),
    // quel producteur a déclenché ce run — la seule trace qui permettra de
    // constater qu'un cycle réel est bien parti du tronc, condition de retrait
    // de l'ancien producteur (cf. l'en-tête de compute_and_trigger.gs)
    producteur: process.env.DTC_PRODUCTEUR || "inconnu",
    run,
  });
  ecrire(envois);
  console.log(`✅ « ${semaine} » inscrite dans state/envois_hebdo.json`);
  process.exit(0);
}

console.error(`commande inconnue : ${commande}`);
process.exit(2);
