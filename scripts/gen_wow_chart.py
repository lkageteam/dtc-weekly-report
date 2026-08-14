#!/usr/bin/env python3
# ════════════════════════════════════════════════════════════════════
#  Génère le graphique WEEK-OVER-WEEK (POS + BA, un seul PNG, 2 panneaux)
#  avec matplotlib — remplace l'ancien dessin natif pptxgenjs (formes qui
#  provoquaient un "réparer" à l'ouverture quand les échelles % et absolue
#  n'étaient pas alignées sur le même zéro).
#
#  Usage : python scripts/gen_wow_chart.py [report.json] [out.png]
#   défaut : ./inputs/report.json → ./assets/tmp/wow_chart.png
#
#  Règle de conception : les deux axes (% à gauche, absolu/FCFA à droite)
#  sont TOUJOURS centrés symétriquement sur 0 (même si les valeurs sont
#  positives) → le "0" des deux échelles tombe exactement à la même
#  hauteur, donc la courbe/les labels de montant ne sont jamais contraints
#  à l'intérieur de l'échelle des barres de %.
# ════════════════════════════════════════════════════════════════════
import json
import sys
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.font_manager as fm
import matplotlib.pyplot as plt
import numpy as np
from scipy.interpolate import make_interp_spline

ROOT = Path(__file__).parent.parent
BG = "#FFFFFF"; INK = "#1A1A1A"; MUTE = "#6B6B6B"; MUTE2 = "#9A9A9A"; LINE = "#E4E4E4"
GREEN = "#2E9E4F"; REDX = "#D1453B"; YELDK = "#E0B400"

#: Nombre de semaines AFFICHÉES — fenêtre glissante sur les plus récentes.
#: Le calcul, lui, continue de porter sur toute l'histoire (voir `draw_panel`).
FENETRE = 5


def register_fonts():
    reg = ROOT / "fonts" / "MTNBrighterSans-Regular.ttf"
    bold = ROOT / "fonts" / "MTNBrighterSans-Bold.ttf"
    fam_reg = fam_bold = "sans-serif"
    if reg.exists():
        fm.fontManager.addfont(str(reg)); fam_reg = fm.FontProperties(fname=str(reg)).get_name()
    if bold.exists():
        fm.fontManager.addfont(str(bold)); fam_bold = fm.FontProperties(fname=str(bold)).get_name()
    plt.rcParams["font.family"] = fam_reg
    return fam_reg, fam_bold


def fmt_wow(v):
    av = abs(v)
    if av >= 1e6:
        return f"{v / 1e6:.{0 if av >= 1e7 else 1}f} M".replace(".", ",")
    if av >= 1e3:
        return f"{round(v / 1e3)} k"
    return str(round(v))


def draw_panel(ax, ax2, title, weeks, values, objectif, fam_bold):
    # ── FENÊTRE GLISSANTE ────────────────────────────────────────────────
    # Le graphique n'affiche que les FENETRE dernières semaines : au-delà, les
    # colonnes se resserrent et les libellés de période deviennent illisibles.
    # ⚠️ Le % vs semaine précédente est calculé sur la série COMPLÈTE, PUIS
    # découpé. Découper d'abord priverait la plus ancienne semaine affichée de
    # son prédécesseur : elle perdrait sa barre alors que la donnée existe.
    serie = np.array(values, dtype=float)
    total = len(serie)
    pct_complet = [None] + [
        (serie[i] - serie[i - 1]) / serie[i - 1] * 100 if serie[i - 1] else None
        for i in range(1, total)
    ]
    debut = max(0, total - FENETRE)
    values = serie[debut:]
    pct = pct_complet[debut:]
    weeks = list(weeks)[debut:]

    n = len(values)
    x = np.arange(n)

    mags = [abs(p) for p in pct if p is not None]
    max_mag = max(10, *(mags or [10])) * 1.25
    ax.set_ylim(-max_mag, max_mag)
    ax.axhline(0, color=LINE, lw=1, zorder=1)
    ax.set_ylabel("% vs semaine précédente", fontsize=9, color=MUTE)

    # axe absolu (droite) — centré sur 0 comme l'axe % pour un alignement garanti
    vmax = max(abs(values.max()), abs(objectif or 0), 1)
    q = vmax * 1.25
    ax2.set_ylim(-q, q)

    # barres de variation (vert hausse / rouge baisse) — depuis la 1re semaine
    # affichée, son % venant de la semaine juste avant la fenêtre
    for i in range(n):
        p = pct[i]
        if p is None:
            continue
        color = GREEN if p >= 0 else REDX
        ax.bar(x[i], p, width=0.5, color=color, zorder=3)
        # Le % va DANS la barre quand elle est assez haute, dehors sinon. Deux
        # familles d'étiquettes (% des barres, montants de la courbe) se
        # disputaient la même bande au-dessus du zéro et se recouvraient — les
        # rentrer dans la barre les sépare pour de bon, sans cadre blanc.
        if abs(p) >= max_mag * 0.20:
            ax.text(x[i], p - (max_mag * 0.045 if p >= 0 else -max_mag * 0.045), f"{p:+.1f} %",
                    ha="center", va="top" if p >= 0 else "bottom", fontsize=9,
                    fontweight="bold", color="white", fontfamily=fam_bold, zorder=10)
        else:
            ax.text(x[i], p + (max_mag * 0.06 if p >= 0 else -max_mag * 0.06), f"{p:+.1f} %",
                    ha="center", va="bottom" if p >= 0 else "top", fontsize=9,
                    fontweight="bold", color=color, fontfamily=fam_bold,
                    bbox=dict(facecolor="white", edgecolor="none", alpha=0.85, pad=1.5), zorder=10)

    # courbe lissée du total réalisé (axe absolu)
    if n >= 3:
        xs = np.linspace(0, n - 1, 200)
        spline = make_interp_spline(x, values, k=min(3, n - 1))
        ys = spline(xs)
        ax2.plot(xs, ys, color=YELDK, lw=2.5, zorder=4, solid_capstyle="round")
    else:
        ax2.plot(x, values, color=YELDK, lw=2.5, zorder=4)
    ax2.scatter(x, values, color=YELDK, edgecolor="white", s=45, zorder=8, linewidths=1.2)
    for i, v in enumerate(values):
        # Toujours AU-DESSUS du point. L'alternance haut/bas servait à écarter les
        # étiquettes quand 8 colonnes se serraient ; à 5 la place horizontale suffit,
        # et alterner renvoyait une étiquette sur deux dans la zone des barres.
        offset = q * 0.13
        ax2.text(x[i], v + offset, fmt_wow(v), ha="center", va="bottom", fontsize=8.5, fontweight="bold", color=INK, fontfamily=fam_bold,
                 bbox=dict(facecolor="white", edgecolor="none", alpha=0.85, pad=1.5), zorder=9)

    # ligne d'objectif (axe absolu), clampée si hors échelle
    oy, suf = objectif, ""
    if objectif is not None:
        if objectif > q: oy, suf = q, " ↑"
        elif objectif < -q: oy, suf = -q, " ↓"
        ax2.axhline(oy, color=INK, lw=1.25, ls=(0, (5, 3)), zorder=2)
        ax2.text(n - 1, oy + q * 0.045, f"Objectif {fmt_wow(objectif)}{suf}", ha="right", va="bottom", fontsize=8.5, fontweight="bold", color=INK, fontfamily=fam_bold,
                 bbox=dict(facecolor="white", edgecolor="none", alpha=0.9, pad=1.5), zorder=7)

    ax.set_xlim(-0.6, n - 1 + 0.9)
    # Libellés horizontaux : à 5 colonnes, « 31 juil – 6 août » tient sans rotation,
    # et un texte droit se lit plus vite qu'un texte incliné.
    ax.set_xticks(x); ax.set_xticklabels(weeks, fontsize=9, color=INK, fontfamily=fam_bold)
    ax.set_title(title, fontsize=13, fontweight="bold", color=INK, fontfamily=fam_bold, loc="left", pad=10)
    for spine in ("top", "right", "left"): ax.spines[spine].set_visible(False)
    for spine in ("top", "left", "right"): ax2.spines[spine].set_visible(False)
    ax.spines["bottom"].set_color(LINE); ax2.spines["bottom"].set_color(LINE)
    ax.tick_params(axis="y", labelsize=8, colors=MUTE2)
    ax2.set_yticks([])
    ax.tick_params(axis="x", length=0)


def main():
    report_path = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "inputs" / "report.json"
    out_path = Path(sys.argv[2]) if len(sys.argv) > 2 else ROOT / "assets" / "tmp" / "wow_chart.png"
    out_path.parent.mkdir(parents=True, exist_ok=True)

    data = json.loads(report_path.read_text(encoding="utf-8"))
    wow = data.get("wow") or {}
    # `periodes` (« 3 – 9 août ») remplace `weeks` (« S8 ») quand le producteur la
    # fournit. Le repli sur `weeks` garde le graphique fonctionnel avec un
    # report.json produit avant ce changement.
    weeks = wow.get("periodes") or wow.get("weeks") or []
    if len(weeks) < 2:
        print("⚠️ Pas assez de semaines pour un graphique WoW (< 2) — aucun PNG généré.")
        return

    fam_reg, fam_bold = register_fonts()
    fig, (axL, axR) = plt.subplots(1, 2, figsize=(13.6, 5.0), dpi=200)
    fig.patch.set_alpha(0)
    axL2, axR2 = axL.twinx(), axR.twinx()

    draw_panel(axL, axL2, "POS — Volume réalisé", weeks, wow.get("pos") or [], wow.get("posObjectif"), fam_bold)
    draw_panel(axR, axR2, "BA — Montant réalisé", weeks, wow.get("ba") or [], wow.get("baObjectif"), fam_bold)

    # légende partagée en haut de figure
    handles = [
        plt.Rectangle((0, 0), 1, 1, color=GREEN, label="Hausse s/s"),
        plt.Rectangle((0, 0), 1, 1, color=REDX, label="Baisse s/s"),
        plt.Line2D([0], [0], color=YELDK, lw=2.5, label="Total réalisé"),
        plt.Line2D([0], [0], color=INK, lw=1.5, ls=(0, (5, 3)), label="Objectif"),
    ]
    fig.legend(handles=handles, loc="upper center", ncol=4, frameon=False, fontsize=10,
               bbox_to_anchor=(0.5, 0.99), prop=fm.FontProperties(fname=str(ROOT / "fonts" / "MTNBrighterSans-Regular.ttf")) if (ROOT / "fonts" / "MTNBrighterSans-Regular.ttf").exists() else None)

    fig.tight_layout(rect=[0, 0, 1, 0.85])
    fig.savefig(out_path, transparent=True, bbox_inches=None)
    plt.close(fig)
    print(f"✅ {out_path}")


if __name__ == "__main__":
    main()
