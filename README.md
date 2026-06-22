# Rapport hebdomadaire DTC Assisted — MTN Bénin × LKA

Génère automatiquement le rapport `.pptx` (6 slides) à partir des données déjà produites par
la chaîne DTC (Gmail → Apps Script → MySQL → Google Sheets).

## Chaîne complète

```
Gmail MTN ─► Apps Script ─► MySQL ─► syncSheetsFromMySQL()
                                        │ écrit les onglets CLASSEMENT_TSA + TSA_REF
                                        ▼
                            ┌─ Google Sheet (à jour) ─┐
                            │  + Sheet réponses Form BA │
                                        │
   (fin de sync)  Apps Script ─► repository_dispatch ─► GitHub Actions (ce repo)
                                                              │
                       node scripts/progress_report_dtc_assisted.js
                       lit les 3 onglets ─► .pptx ─► artifact (+ mail optionnel)
```

Le rapport est **un consommateur de plus** des onglets que la chaîne maintient déjà : aucune
transformation, les en-têtes des Sheets sont identiques à ceux attendus par le script.

## Données attendues (`inputs/`)

| Fichier | Onglet source | En-têtes |
|---|---|---|
| `classement.csv` | `CLASSEMENT_TSA` | `CORPORATE_NUM, SEMAINE, RANG_GLOBAL, RANG_REGIONAL, TSA, REGION, RBM, VOLUME_XAF, NB_POS_ACTIFS, NB_POS_ASSIGNES, TAUX_ACTIVATION, PRIME_ELIGIBLE` |
| `tsa_ref.csv` | `TSA_REF` | `CORPORATE_NUM, TSA_FULL_NAME, REGION, RBM_NAME, GM_NAME, ACTIVE, OBJECTIF_DAILY` |
| `form.csv` | Réponses du Form BA | colonnes du formulaire (Timestamp, Region, Numéro du BA, Montant…) |

En CI, le workflow télécharge ces 3 onglets (URLs en secrets) avant de lancer le script.

## Règles métier (résumé)

- **Objectif TSA** : 12 M/jour (= somme des `OBJECTIF_DAILY` de `TSA_REF`) → 84 M/semaine.
- **Objectif BA** : 1,5 M/jour global, réparti par % d'effectif régional (≈ 6 787/BA/jour).
- **Prime TSA** : être n°1 de sa région **ET** volume ≥ 450 000 FCFA.
- **Semaine** : sans argument, le script prend la **dernière `SEMAINE`** présente dans `classement.csv`.
- Ordre des slides : Couverture · Recrutement · Activations POS · Prime & Rang · Activations BA · Clôture.

## Lancer en local

```bash
npm install
# déposer classement.csv / tsa_ref.csv / form.csv dans ./inputs
node scripts/progress_report_dtc_assisted.js          # dernière semaine auto
node scripts/progress_report_dtc_assisted.js 1        # forcer Semaine 1
# → ./outputs/Progress_Report_DTC_Assisted_LKA_SemaineN.pptx
```

## Configuration GitHub (Settings → Secrets and variables → Actions)

**Données (obligatoire)** — URLs CSV des onglets (publier l'onglet en CSV, ou lien gviz
`…/gviz/tq?tqx=out:csv&sheet=CLASSEMENT_TSA` si le Sheet est partagé « tous avec le lien ») :

- `CLASSEMENT_CSV_URL`
- `TSA_REF_CSV_URL`
- `FORM_CSV_URL`

**Mail (optionnel)** — l'étape d'envoi s'active seulement si `SMTP_HOST` est défini :

- `SMTP_HOST`, `SMTP_PORT`, `SMTP_USER`, `SMTP_PASS`, `MAIL_TO`

## Déclencheur depuis l'Apps Script existant

À appeler à la fin de `syncSheetsFromMySQL()` (ou via un trigger hebdo). Stocker le PAT
(scope `repo`) dans les Script Properties, pas en clair.

```javascript
function triggerWeeklyReport() {
  var token = PropertiesService.getScriptProperties().getProperty('GH_TOKEN');
  UrlFetchApp.fetch('https://api.github.com/repos/lkageteam/dtc-weekly-report/dispatches', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'token ' + token, Accept: 'application/vnd.github+json' },
    payload: JSON.stringify({ event_type: 'weekly-report' }),  // option : client_payload: { semaine: '2' }
    muteHttpExceptions: true
  });
}
```

## Lancement manuel

Onglet **Actions → Rapport hebdomadaire DTC → Run workflow** (champ `semaine` optionnel).
Le `.pptx` est récupérable dans les **Artifacts** du run.
