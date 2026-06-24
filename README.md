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
- Slides : Couverture · Recrutement · Activations POS · Prime & Rang · Activations BA · Clôture.
