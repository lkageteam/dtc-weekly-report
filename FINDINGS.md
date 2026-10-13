# dtc-weekly-report — faits mesurés

## 2026-09-03 — Le mensuel n'a jamais eu de déclencheur. Il n'en a jamais eu.

**Mesuré**, pas supposé. Le rapport d'août n'est pas « en retard » : rien, nulle
part, n'a jamais été construit pour le produire.

Le rapport **hebdomadaire** est, lui, correctement branché sur un FAIT :

| Maillon | Fichier | Ce qu'il fait |
| :--- | :--- | :--- |
| observation du fait | `appscript/compute_and_trigger.gs::maybeAutoSendWeekly` | `_rptWeekPresent_(N)` — la semaine N est-elle DANS le classement ? |
| réveil | même fichier, l. 76 | `repository_dispatch` `event_type: 'weekly-report'` |
| rendu + envoi | `.github/workflows/weekly-report.yml` (l. 10-11) | `types: [weekly-report, weekly-report-preview]` |
| idempotence | `scripts/etat_envois.js` + `state/envois_hebdo.json` | une période déjà envoyée n'est pas renvoyée |

Le **mensuel** n'a AUCUN de ces quatre maillons :

- `grep -n "monthly\|Mensuel\|mensuel" appscript/compute_and_trigger.gs` → **0 ligne**.
- `weekly-report.yml` n'accepte que `weekly-report` / `weekly-report-preview` :
  aucun `event_type` mensuel n'existe, donc aucun ne peut être reçu.
- il n'y a pas de `state/envois_mensuel.json` — donc pas de clé d'idempotence,
  donc **aucune protection contre le doublon** le jour où on l'automatisera.
- `scripts/monthly_report_dtc_assisted.js` n'est atteignable que depuis un shell
  local, à la main. Les fichiers `outputs/` le confirment : `202606` écrit le
  09/07, `202607` écrit le 06/08 — deux gestes manuels, à neuf jours d'écart de
  cadence. Août n'a pas été fait parce que personne ne l'a tapé.

➡️ **La cause n'est pas une panne, c'est une absence.** Ne pas chercher un job
rouge : il n'y en a pas, et un job vert n'aurait rien prouvé non plus.

### Le fait qui devrait le déclencher

« Le mois M est bouclé » = les DEUX conditions, pas une seule :

1. le dernier jour du mois M est passé (sinon on livre un mois incomplet), **et**
2. `pos_sales` porte bien une ligne pour **le dernier jour** de M —
   `SELECT MAX(sale_date) FROM lka_client_mtn.pos_sales >= LAST_DAY(M)`.

⚠️ Aucune des deux ne suffit seule, et **la 2 est celle qui bloque aujourd'hui** :
au 03/09, `MAX(sale_date)` vaut `2026-08-30`. Le 31/08 n'est pas chargé. Août a
donc été livré sur 30 jours, et la slide le dit. Un déclencheur qui n'aurait
testé que la condition 1 aurait livré le même trou **sans le dire**.

⚠️ Ne PAS reprendre le test de l'hebdo (`_rptWeekPresent_`) pour le mensuel : il
interroge le classement, qui est hebdomadaire. La semaine 11 (21→27/08) y était
dès le 28/08, quand le mois n'était pas fini.

⚠️ Et **aucun cron** : le fait est observable (`pos_sales` le porte), donc un
cron mensuel serait un second sondage, pas un filet.

### Ce qui reste à faire, et qui n'a PAS été fait le 2026-09-03

Le câblage complet (`maybeAutoSendMonthly` + `event_type: monthly-report` +
`state/envois_mensuel.json`) n'a pas été posé : il ouvre un **envoi automatique
de mail** vers de vraies personnes, sans clé d'idempotence existante. C'est un
premier envoi dans un canal neuf — arbitrage du propriétaire, pas du script.
Août 2026 a donc été produit **à la main** ce jour-là.

---

## 2026-09-03 — Un MOIS se compte en JOURS. La première version comptait en semaines.

**Erreur commise puis corrigée le jour même**, sur remarque du propriétaire. Elle
est écrite ici parce qu'elle est invisible à l'exécution : les deux versions
sortent en code 0, avec des chiffres plausibles.

`classement.csv` (vue `v_classement_tsa_semaine`) **ne porte pas de date** :
chaque ligne porte une SEMAINE POS (ven→jeu, ancrée au 12/06). Agréger un mois en
retenant « les semaines qui chevauchent le mois » déborde donc des deux côtés.

| juillet 2026, Nord Ouest | fenêtre réelle | volume |
| :--- | :--- | ---: |
| par semaines (S3→S8) | 26/06 → 06/08, **42 jours** | 25 910 800 |
| par jours (`pos_sales`) | 01/07 → 31/07, **31 jours** | 19 557 450 |

+32 % attribués à juillet, dont une partie appartient à juin et une partie à août.
Le classement des 7 en est changé : par semaines le Nord Est était 2ᵉ de juillet
avec 20,7 M, par jours il est 2ᵉ avec 15,4 M, et les rangs 3/4 (Sud Est / Sud
Ouest) s'inversent. Le récit change aussi : par semaines juillet (25,9 M) écrasait
août (18,7 M) ; par jours août (20,2 M) dépasse juillet (19,6 M).

➡️ Le jour EXISTE : `lka_client_mtn.pos_sales`, une ligne par POS et par jour,
156 960 lignes du 12/06 au 30/08. `fetch_live_data.py` le dumpe désormais en
`pos_daily.csv`, et `lib/pos_jour.js` en est la seule porte d'entrée. Le rapport
**hebdomadaire** garde le classement — sa fenêtre EST une semaine.

⚠️ **La cible se multiplie par les jours PRÉSENTS, pas par ceux du calendrier.**
Au 03/09, `pos_sales` s'arrête au 30/08 : août compte 30 jours, pas 31. Multiplier
la cible par 31 aurait inventé un retard d'un jour sur tout le monde.

⚠️ Le seuil de prime TSA suivait la même faute : `primeThreshold × nb de semaines`
valait 6 semaines pour les 31 jours de juillet. Il est désormais
`primeThreshold × jours / 7`.

---

## 2026-09-03 — Prime RBM : le calcul vit dans `scripts/lib/prime_rbm.js`

Critère tranché par le propriétaire : **volume POS du mois**, un seul primé sur
les 7, **mois pleins uniquement** (juin exclu — programme démarré le 15/06).

Ce que ça change, mesuré sur les données live du 03/09 : le critère n'est pas
neutre, il désigne des lauréats différents.

Chiffres au JOUR (voir la section précédente) :

| Mois | 1er (volume POS) | volume | atteinte | 2e |
| :--- | :--- | ---: | ---: | :--- |
| juillet (31 j) | **VINASSOU LOUIS** — Nord Ouest | 19 557 450 | 43,7 % | Adjicoude Bienvenu (Nord Est) 15 416 950 |
| août (30 j) | **VINASSOU LOUIS** — Nord Ouest | 20 222 850 | 46,7 % | Adjicoude Bienvenu (Nord Est) 16 585 300 |

Même lauréat les deux mois. Le critère n'est pas neutre : sur le volume BA,
juillet aurait primé ADJICOUDE BIENVENU et août ATTIKA ALBERT — deux lauréats
différents. Ne pas changer de critère sans redire lequel
dans la slide : la colonne « CRITÈRE » y est écrite en clair pour ça.

⚠️ `scripts/fetch_live_data.py` **ne charge pas** `D:\LKA\DTC Pushed\.env` ; sans
export préalable des variables il tombe sur `localhost` et rend
`ConnectionRefusedError`. Lancer :
`set -a && . "/d/LKA/DTC Pushed/.env" && set +a && python scripts/fetch_live_data.py`
