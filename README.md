# Rapport hebdomadaire DTC Assisted — MTN Bénin × LKA

Génère le rapport `.pptx` (6 slides). **Répartition des rôles :**

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

## Côté Apps Script (chaîne `DTC Pushed`)

1. Coller `appscript/compute_and_trigger.gs` dans le projet Apps Script lié au classeur (celui qui contient `CLASSEMENT_TSA` et `TSA_REF`).
2. Script Property **`GH_TOKEN`** = PAT GitHub (fine-grained sur `lkageteam/dtc-weekly-report`, *Contents: Read and write* ; ou classic `repo`).
3. Le Form BA est déjà câblé (`FORM_SS_ID` + `FORM_GID` en haut du fichier).
4. Appeler **`buildAndTriggerReport()`** à la fin de `syncSheetsFromMySQL()` (ou via un trigger hebdo).
   - `previewReport()` logge le JSON sans déclencher (pour vérifier les chiffres).

## Côté GitHub

- Rien d'obligatoire pour les données (elles arrivent par le payload).
- **Mail (optionnel)** — l'étape d'envoi s'active si `SMTP_HOST` est défini :
  `SMTP_HOST`, `SMTP_PORT`, `SMTP_USER`, `SMTP_PASS`, `MAIL_TO`.
- Le `.pptx` est toujours récupérable dans les **Artifacts** du run (onglet Actions).

## Règles métier (résumé)

- **Objectif TSA** : 12 M/jour (= somme des `OBJECTIF_DAILY` de `TSA_REF`) → 84 M/semaine.
- **Objectif BA** : 1,5 M/jour global, réparti par % d'effectif régional.
- **Prime TSA** : n°1 de sa région **ET** volume ≥ 450 000 FCFA.
- **Semaine** : par défaut, la dernière `SEMAINE` présente dans le classement.
- Slides : Couverture · Recrutement · Activations POS · Prime & Rang · Activations BA · Clôture.
