#!/usr/bin/env python3
# ════════════════════════════════════════════════════════════════════
#  Récupère les données LIVE directement depuis la base MySQL de la
#  chaîne DTC (celle qui alimente aussi les Sheets CLASSEMENT_TSA / TSA_REF)
#  et écrit les 3 CSV attendus par les scripts de rendu (weekly & monthly)
#  dans un dossier "inputs_live" — pour ne jamais avoir à dépendre d'un
#  export Sheets/Form manuel, ni des CSV de test (souvent stale) de inputs/.
#
#  Usage : python scripts/fetch_live_data.py [--out DIR]
#   → DIR/classement.csv, DIR/tsa_ref.csv, DIR/form.csv
#   (DIR défaut = ./inputs_live à côté de ce repo)
#
#  Connexion : credentials dans D:\LKA\DTC Pushed\.env (chaîne LKA×MTN,
#  MySQL flaky — retry avec backoff, cf. mémoire [[dtc-data-pipeline]]).
# ════════════════════════════════════════════════════════════════════
import csv
import os
import sys
import time
from pathlib import Path

import pymysql

MYSQL_HOST = os.getenv("MYSQL_HOST", "75.119.154.255")
MYSQL_PORT = int(os.getenv("MYSQL_PORT", "3306"))
MYSQL_USER = os.getenv("MYSQL_USER", "root")
MYSQL_PASSWORD = os.getenv("MYSQL_PASSWORD", "LkaRoot2025Secure!")
MYSQL_DATABASE = os.getenv("MYSQL_DATABASE", "lka_client_mtn")


def connect(retries=9, delay=4):
    last = None
    for i in range(retries):
        try:
            return pymysql.connect(
                host=MYSQL_HOST, port=MYSQL_PORT, user=MYSQL_USER,
                password=MYSQL_PASSWORD, database=MYSQL_DATABASE, connect_timeout=15,
            )
        except Exception as e:
            last = e
            print(f"  … tentative {i + 1}/{retries} échouée ({e}), retry dans {delay}s", file=sys.stderr)
            time.sleep(delay)
    raise last


def write_csv(path, header, rows):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)
    print(f"✅ {path} ({len(rows)} lignes)")


def main():
    out = Path(sys.argv[sys.argv.index("--out") + 1]) if "--out" in sys.argv else Path(__file__).parent.parent / "inputs_live"
    out.mkdir(parents=True, exist_ok=True)

    print(f"Connexion à {MYSQL_HOST}:{MYSQL_PORT}/{MYSQL_DATABASE} …")
    conn = connect()
    cur = conn.cursor()

    # CLASSEMENT_TSA (v_classement_tsa_semaine) → Prime&Rang + réalisation POS
    cur.execute("""
        SELECT tsa_corporate, tsa_full_name, region, rbm_name, semaine, volume_xaf,
               nb_pos_actifs, rang_regional, taux_activation
        FROM v_classement_tsa_semaine
    """)
    rows = [[r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8]] for r in cur.fetchall()]
    write_csv(out / "classement.csv",
              ["TSA", "TSA_FULL_NAME", "REGION", "RBM", "SEMAINE", "VOLUME_XAF", "NB_POS_ACTIFS", "RANG_REGIONAL", "TAUX_ACTIVATION"],
              rows)

    # TSA_REF (tsa_reference_mtn) → cibles POS journalières par région
    cur.execute("SELECT corporate_num, tsa_full_name, region, rbm_name, objectif_daily FROM tsa_reference_mtn WHERE active = 1")
    rows = [[r[0], r[1], r[2], r[3], r[4]] for r in cur.fetchall()]
    write_csv(out / "tsa_ref.csv", ["CORPORATE_NUM", "TSA_FULL_NAME", "REGION", "RBM", "OBJECTIF_DAILY"], rows)

    # FORM BA (dtc_ba_activations) → Activations BA par région, jour par jour
    # numero_ba (numéro de téléphone du BA) = identifiant stable pour compter les BA actifs
    # (contrairement à l'email du Form, qui n'est pas mirroré ici et que les agents changeaient souvent).
    cur.execute("SELECT timestamp, region, montant_activation, numero_ba FROM dtc_ba_activations ORDER BY timestamp")
    def fmt_ts(dt):  # M/D/YYYY H:MM:SS — format attendu par parseTS() côté JS
        return f"{dt.month}/{dt.day}/{dt.year} {dt.hour}:{dt.minute:02d}:{dt.second:02d}"
    rows = [[fmt_ts(r[0]), r[1], r[2], r[3]] for r in cur.fetchall()]
    write_csv(out / "form.csv", ["Timestamp", "Region", "Montant d'activation", "Numero du BA"], rows)

    cur.close()
    conn.close()
    print(f"\n📂 Données live écrites dans {out} — lancer les scripts de rendu avec INPUT_DIR={out}")


if __name__ == "__main__":
    main()
