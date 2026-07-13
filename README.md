# Rapport hebdomadaire DTC Assisted — MTN Bénin × LKA

Génère le rapport `.pptx` (9 slides). **Répartition des rôles :**

- **Apps Script** (chaîne DTC) = lit les données **et fait tous les calculs** (régions, rangs, targets, prime, série jour-par-jour). Il a déjà les données en main.
- **GitHub Actions / Node** = **dessine uniquement** le `.pptx` (pptxgenjs — la seule chose qu'Apps Script ne sait pas faire).

```
Gmail MTN ─► Apps Script ─► MySQL ─► Sheets (CLASSEMENT_TSA, TSA_REF)
                  │
                  │ buildAndTriggerReport() : calcule un PETIT JSON (~5 Ko)
                  ▼
        repository_dispatch (client_payload.report) ─► GitHub Actions
                                                            │
                            node scripts/… (moteur de rendu) ─► .pptx ─► artifact (+ mail)
```

Aucun CSV, aucune donnée brute n'est transmise : seules les **valeurs finales** transitent.

## Deux modes d'exécution du script Node

| Mode | Entrée | Usage |
|------|--------|-------|
| **PROD** | `inputs/report.json` (écrit depuis `client_payload`) | GitHub Actions — rend uniquement |
| **DEV** | `inputs/classement.csv`, `tsa_ref.csv`, `form.csv` | local — calcule puis rend (pour tester le rendu) |

```bash
npm install
# DEV : déposer les 3 CSV dans ./inputs puis
node scripts/progress_report_dtc_assisted.js          # dernière semaine auto
node scripts/progress_report_dtc_assisted.js 1        # forcer Semaine 1
node scripts/progress_report_dtc_assisted.js 1 --dump # écrit inputs/report.json (= le JSON-contrat)
# → ./outputs/Progress_Report_DTC_Assisted_LKA_SemaineN.pptx
```

`--dump` produit exactement le JSON que l'Apps Script doit envoyer (fixture + spec du contrat).

## Récupérer les données LIVE (MySQL) avant un rendu DEV

`inputs/*.csv` est un jeu de test figé (peut être stale). Pour des chiffres réels à jour, interroger directement la base MySQL de la chaîne DTC (celle qui alimente aussi les Sheets) :

```bash
python scripts/fetch_live_data.py                       # → ./inputs_live/{classement,tsa_ref,form}.csv
python scripts/fetch_live_data.py --out /autre/dossier
```

Nécessite `pymysql` (`pip install pymysql`). Credentials MySQL dans `D:\LKA\DTC Pushed\.env` (chaîne LKA×MTN — voir mémoire `dtc-data-pipeline`), avec retry intégré (la base est connue pour être flaky). Source directe : vue `v_classement_tsa_semaine` (Prime&Rang/POS), table `tsa_reference_mtn` (cibles POS), table `dtc_ba_activations` (activations BA — **déjà en MySQL**, inutile de passer par un export du Form Google). Puis pointer `INPUT_DIR` vers ce dossier pour n'importe quel script de rendu :

```bash
INPUT_DIR=./inputs_live node scripts/progress_report_dtc_assisted.js 3
INPUT_DIR=./inputs_live node scripts/monthly_report_dtc_assisted.js 2026-06
```

## Rapport MENSUEL (POS + BA)

```bash
node scripts/monthly_report_dtc_assisted.js          # mois du dernier Timestamp du form.csv
node scripts/monthly_report_dtc_assisted.js 2026-07  # forcer un mois (AAAA-MM)
# → ./outputs/Progress_Report_DTC_Assisted_LKA_Mensuel_AAAAMM.pptx
```

2 slides seulement (Activations POS + Activations BA), même style que l'hebdo, agrégées sur un **mois calendaire** borné au démarrage réel du programme (15 juin 2026 — donc juin 2026 = fenêtre 15→30 juin ; les mois suivants couvrent le mois complet). Objectif du mois = objectif **journalier** (même taux que l'hebdo, TSA_REF pour le POS / 1,5M réparti par effectif pour le BA) **× nombre de jours de la fenêtre** — jours ouvrés hors dimanche pour le BA (comme l'hebdo ×6/7), tous les jours pour le POS (comme l'hebdo ×7/7). Réalisé POS = somme des "Semaine N" du classement dont la période chevauche la fenêtre (le classement n'a pas de date par ligne). Réalisé BA = form filtré par Timestamp sur la fenêtre ; le graphique "rythme journalier" couvre alors tout le mois (pas juste 7 jours).

Mode DEV uniquement pour l'instant (lit les mêmes CSV que l'hebdo dans `inputs/`) — pas encore automatisé/câblé à l'Apps Script ; à faire plus tard si besoin d'un envoi mensuel automatique. Le thème/palette/helpers pptx partagés avec l'hebdo sont dans `scripts/lib/theme.js` (à réutiliser pour tout futur script de rendu — ne pas redupliquer les constantes).

## Côté Apps Script (chaîne `DTC Pushed`)

1. Coller `appscript/compute_and_trigger.gs` dans le projet Apps Script lié au classeur (`CLASSEMENT_TSA` + `TSA_REF`).
2. Script Property **`GH_TOKEN`** = PAT GitHub (fine-grained sur `lkageteam/dtc-weekly-report`, *Contents: Read and write* ; ou classic `repo`).
3. Le Form BA est déjà câblé (`FORM_SS_ID` + `FORM_GID` en haut du fichier).
4. Exécuter **`installWeeklyTriggers()`** une fois → filet **lundi 8h / 9h / 10h**.
5. Ajouter **`maybeSendReport();`** à la fin de `syncSheetsFromMySQL()` → envoi dès que les données du lundi sont synchronisées.

**Déclenchement automatique** — `maybeSendReport()` envoie la **dernière semaine BA bouclée** (lun→dim), calculée par date (`targetN = floor((aujourd'hui − 15 juin)/7 j)`), **une seule fois** (état `LAST_SENT_WEEK`), et seulement si ses données sont dans le classement. Pourquoi le lundi : le BA ferme dimanche et le POS jeudi → le lundi tout est complet. L'appel dans `syncSheetsFromMySQL()` (événementiel) et le cron du lundi (filet) sont idempotents grâce à l'état.

**Fonctions** : `previewReport()` (logge le JSON sans envoyer) · `buildAndTriggerReport("1")` (force une semaine) · `maybeSendReport()` (auto) · `resetLastSentWeek()` (ré-autorise un renvoi).

## Côté GitHub

- Rien d'obligatoire pour les données (elles arrivent par le payload).
- **Mail → `g.fondzefe@lkaservices.com`** (destinataire fixé dans le workflow), envoyé via l'**API Gmail (OAuth)** — pas de SMTP. Un seul secret :
  - **`GMAIL_TOKEN_JSON`** = le contenu intégral de `connections/token.json` de la chaîne LKA (compte `joselonm11@gmail.com`, scope `gmail.modify`). `scripts/send_mail.js` rafraîchit l'access token et envoie le `.pptx` en pièce jointe, sans aucune dépendance.
  - ⚠️ Le `refresh_token` doit rester valide : l'écran de consentement OAuth (Google Cloud) doit être **« In production »** (sinon les tokens d'apps en *Testing* expirent au bout de 7 jours).
- Le `.pptx` reste aussi récupérable dans les **Artifacts** du run (onglet Actions).

## Règles métier (résumé)

- **Objectif TSA** : 12 M/jour (= somme des `OBJECTIF_DAILY` de `TSA_REF`) → 84 M/semaine.
- **Objectif BA** : 1,5 M/jour global, réparti par % d'effectif régional.
- **Prime TSA** : n°1 de sa région **ET** volume ≥ 450 000 FCFA.
- **Semaine** : par défaut, la dernière `SEMAINE` présente dans le classement.
- **Retours terrain** : répartition cumulée (toutes semaines) des réponses à « Quel est le problème principal rencontré ? » du Form, classées en 9 catégories + « Divers / Autres » (les « non / RAS / aucun » sont exclus).
- Slides : Couverture · Recrutement · Activations POS · Prime & Rang · Activations BA · Week-over-week (POS+BA, 1 slide, graphique matplotlib) · Analyse BA (transactions/montant/personnes actives) · Retours terrain · Clôture.
- Le graphique Week-over-week est une image PNG générée par `python scripts/gen_wow_chart.py inputs/report.json assets/tmp/wow_chart.png` (nécessite `pip install -r scripts/requirements.txt`) — la CI le fait automatiquement avant le rendu. En DEV local sans Python, la slide affiche un message au lieu de planter.
