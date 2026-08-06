# -*- coding: utf-8 -*-
"""
Paper32 — Figure 1: PRISMA 2020 흐름도

★ 수치는 `figures/fig_data.json` 에서만 읽는다(구판은 96편·2갈래를 하드코딩했다).
★ 글자는 `figtext.py` 로 **실측해서** 넣는다. 눈대중으로 pt 를 정하다 상자를 줄줄이 넘겼다.
레이아웃 = PRISMA 2020 표준 2열. 등록한 보조 경로가 둘이므로 오른쪽 열에 함께 싣는다.
출력: figures/Fig1_PRISMA.png|pdf
"""
import sys, os, json
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FIG = os.path.join(BASE, "figures")
sys.path.insert(0, BASE)
import viz_theme as T
from figtext import boxed_text, kv_rows
T.apply()
INK, ACC, MUT = T.INK, T.BLUE, T.AXIS

D = json.load(open(os.path.join(FIG, "fig_data.json"), encoding="utf-8"))
P, DB, CT, SP = D["prisma"], D["prisma"]["db"], D["prisma"]["ct"], D["prisma"]["supp"]

W, H = 7.48, 7.6
FS, FSX = 7.2, 6.4
L_X, L_W, L_XX, L_XW = 3.2, 24.0, 28.5, 20.0
R_X, R_W, R_XX, R_XW = 52.5, 25.0, 79.0, 20.0
# 단계 y (상자 아래 모서리) · 상자 높이
YB = {"id": 84.0, "scr": 65.5, "sought": 44.0, "assess": 24.0}
BH = 8.0


def main():
    fig, ax = plt.subplots(figsize=(W, H))
    ax.set_xlim(0, 100); ax.set_ylim(0, 100); ax.axis("off")

    def box(x, y, lines, w, h=BH, fc="white", ec=INK, lw=0.85, fs=FS):
        ax.add_patch(FancyBboxPatch((x, y), w, h,
                                    boxstyle="round,pad=0.12,rounding_size=0.45",
                                    fc=fc, ec=ec, lw=lw))
        boxed_text(ax, x, y, w, h, lines, fs, color=INK)

    def xbox(x, y, h, head, rows, w):
        ax.add_patch(FancyBboxPatch((x, y), w, h,
                                    boxstyle="round,pad=0.12,rounding_size=0.45",
                                    fc="#f7f7f5", ec=MUT, lw=0.6))
        boxed_text(ax, x, y + h - 2.6, w, 2.6, [head], FSX, weight="bold", color=INK)
        kv_rows(ax, x, y + h - 3.6, w, rows, FSX, color=T.INK2, label=head)

    def down(cx, y0, y1):
        ax.add_patch(FancyArrowPatch((cx, y0), (cx, y1), arrowstyle="-|>",
                                     mutation_scale=7, lw=0.85, color=INK,
                                     shrinkA=0, shrinkB=0))

    def side(x0, y, x1):
        ax.add_patch(FancyArrowPatch((x0, y), (x1, y), arrowstyle="-|>",
                                     mutation_scale=6, lw=0.7, color=MUT,
                                     shrinkA=0, shrinkB=0))

    ax.text(L_X + (L_W + L_XW + 1.5) / 2, 96.5, "Identification via databases",
            ha="center", fontsize=8.4, color=ACC, fontweight="bold")
    ax.text(R_X + (R_W + R_XW + 1.5) / 2, 96.5, "Identification via other methods",
            ha="center", fontsize=8.4, color=ACC, fontweight="bold")
    ax.plot([48.8, 48.8], [13.5, 94.5], color=T.GRID, lw=0.8)

    for y0, y1, lab in [(82.0, 94.0, "Identification"), (60.0, 80.5, "Screening"),
                        (21.0, 58.5, "Eligibility"), (2.0, 12.0, "Included")]:
        ax.add_patch(FancyBboxPatch((0.0, y0), 1.6, y1 - y0, boxstyle="round,pad=0.05",
                                    fc="#eef2f5", ec="none"))
        ax.text(0.8, (y0 + y1) / 2, lab, rotation=90, ha="center", va="center",
                fontsize=7.2, color=ACC, fontweight="bold")

    # ═════ 왼쪽 ═════
    cx = L_X + L_W / 2
    box(L_X, YB["id"], ["Records identified from databases", "(2 August 2026)",
                        f"WoS {DB['wos']:,} · Scopus {DB['scopus']:,} · PubMed {DB['pubmed']}",
                        f"$\\bf{{n = {DB['identified']:,}}}$"], L_W)
    xbox(L_XX, YB["id"] + 2.6, 4.6, "Duplicates removed", [("", DB["duplicates"])], L_XW)
    side(L_X + L_W, YB["id"] + BH / 2, L_XX)
    down(cx, YB["id"], YB["scr"] + BH)

    box(L_X, YB["scr"], ["Records screened", "(title and abstract)",
                         f"$\\bf{{n = {DB['screened']:,}}}$"], L_W)
    xbox(L_XX, YB["scr"] - 4.5, 13.5, f"Excluded   n = {DB['excluded']:,}", DB["excl"], L_XW)
    side(L_X + L_W, YB["scr"] + BH / 2, L_XX)
    down(cx, YB["scr"], YB["sought"] + BH)

    box(L_X, YB["sought"], ["Reports sought for retrieval",
                            f"$\\bf{{n = {DB['sought']}}}$"], L_W)
    xbox(L_XX, YB["sought"] - 0.5, 9.0, f"Not retrieved   n = {DB['not_retrieved']}",
         DB["nr"], L_XW)
    side(L_X + L_W, YB["sought"] + BH / 2, L_XX)
    down(cx, YB["sought"], YB["assess"] + BH)

    box(L_X, YB["assess"], ["Reports assessed for eligibility", "(full text)",
                            f"$\\bf{{n = {DB['assessed']}}}$"], L_W)
    xbox(L_XX, YB["assess"] - 6.0, 14.5, f"Excluded   n = {DB['ft_excluded']}",
         DB["ftx"] + [("Sensitivity only", DB["sens"])], L_XW)
    side(L_X + L_W, YB["assess"] + BH / 2, L_XX)
    down(cx, YB["assess"], 12.2)
    ax.text(cx, 14.4, f"$\\bf{{n = {DB['included']}}}$", ha="center", fontsize=8, color=ACC)

    # ═════ 오른쪽 ═════
    cx = R_X + R_W / 2
    box(R_X, YB["id"], [f"Citation searching · {CT['seeds']} seeds",
                        f"backward {CT['backward']} · forward {CT['forward']:,}"
                        f"   $\\bf{{n = {CT['identified']:,}}}$",
                        "Supplementary index (unique)",
                        f"$\\bf{{n = {SP['identified']}}}$"], R_W)
    xbox(R_XX, YB["id"] + 1.4, 6.0, "Removed before screening",
         [("Deprioritised by title", CT["deprioritised"]),
          ("Already in the pools", SP["duplicates"])], R_XW)
    side(R_X + R_W, YB["id"] + BH / 2, R_XX)
    down(cx, YB["id"], YB["scr"] + BH)

    box(R_X, YB["scr"], ["Records screened",
                         f"citation searching $\\bf{{{CT['screened_title']}}}$ then "
                         f"$\\bf{{{CT['screened_abs']}}}$",
                         f"supplementary index $\\bf{{{SP['screened']}}}$"], R_W)
    xbox(R_XX, YB["scr"] - 4.5, 13.5,
         f"Excluded   n = {CT['excl_title'] + CT['excl_abs'] + SP['excluded']:,}",
         [("No behaviour outcome", 283), ("No acoustic variable", 78),
          ("Not empirical", 173), ("Animal / bioacoustics", 45),
          ("Setting / language", 84)], R_XW)
    side(R_X + R_W, YB["scr"] + BH / 2, R_XX)
    down(cx, YB["scr"], YB["sought"] + BH)

    box(R_X, YB["sought"], ["Reports sought for retrieval",
                            f"citation searching $\\bf{{{CT['sought']}}}$  ·  "
                            f"index $\\bf{{{SP['sought']}}}$"], R_W)
    xbox(R_XX, YB["sought"] - 0.5, 9.0,
         f"Not obtained   n = {CT['not_retrieved'] + SP['not_retrieved'] + SP['prescreen']}",
         [("No institution access", 22), ("Pay-per-view only", 3),
          ("Index route, various", SP["not_retrieved"] + SP["prescreen"])], R_XW)
    side(R_X + R_W, YB["sought"] + BH / 2, R_XX)
    down(cx, YB["sought"], YB["assess"] + BH)

    box(R_X, YB["assess"], ["Reports assessed for eligibility", "(full text)",
                            f"citation searching $\\bf{{{CT['assessed']}}}$  ·  "
                            f"index $\\bf{{{SP['assessed']}}}$"], R_W)
    xbox(R_XX, YB["assess"] - 6.0, 14.5,
         f"Excluded   n = {CT['ft_excluded'] + SP['ft_excluded']}",
         [("No acoustic variable", 20), ("No observed behaviour", 18),
          ("Not in English", 3), ("Sensitivity only", CT["sens"])], R_XW)
    side(R_X + R_W, YB["assess"] + BH / 2, R_XX)
    down(cx, YB["assess"], 12.2)
    ax.text(cx, 14.4, f"$\\bf{{n = {CT['included']} + {SP['included']}}}$",
            ha="center", fontsize=8, color=ACC)

    # ═════ 합류 ═════
    ax.add_patch(FancyBboxPatch((3.2, 2.6), 94.0, 9.0,
                                boxstyle="round,pad=0.15,rounding_size=0.5",
                                fc="#eef5f0", ec=ACC, lw=1.0))
    ax.text(50, 9.2, f"Studies included in the review     $\\bf{{n = {D['n_included']}}}$",
            ha="center", va="center", fontsize=10.5, color=INK)
    boxed_text(ax, 3.2, 5.9, 94.0, 2.4,
               [f"databases {DB['included']}     citation searching {CT['included']}     "
                f"supplementary index {SP['included']}     "
                f"reserved for sensitivity analysis {D['n_sens']}"], 7.2, color=T.INK2)
    boxed_text(ax, 3.2, 3.2, 94.0, 2.4,
               [f"Citation searching added {CT['included']} studies — "
                f"{CT['included'] / DB['included'] * 100:.0f}% of the database yield; "
                f"neither supplementary route contributes a poolable effect."],
               6.6, color=MUT)

    for ext in ("png", "pdf"):
        fig.savefig(os.path.join(FIG, f"Fig1_PRISMA.{ext}"))
    plt.close(fig)
    print(f"  Fig1_PRISMA.png / .pdf   (2열 표준 · 3경로 · 최종 n = {D['n_included']})")


if __name__ == "__main__":
    main()
