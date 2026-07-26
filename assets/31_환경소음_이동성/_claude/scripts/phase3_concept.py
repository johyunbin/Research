# -*- coding: utf-8 -*-
# Fig 2 (재작도): 식별전략을 직관적으로. (a) 인과·교란 (b) 같은 날 동 사이 비교(공통요인 상쇄).
import os
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
import figstyle as FS

FS.apply_style()
ROOT = r"T:\00_BACKUP\01_한양대학교\00_연구실\00_프로젝트\★_논문화 프로젝트\31_환경소음_이동성"
FIG = os.path.join(ROOT, "_claude", "figures")
GREEN, MOB, NOISE, CONF = "#7FB8A4", FS.ACCENT["mobility"], FS.ACCENT["noise"], "#DDD3C2"

def box(ax, x, y, w, h, lines, fc, fs=9, tc="#222", bold_first=False):
    ax.add_patch(FancyBboxPatch((x - w/2, y - h/2), w, h, boxstyle="round,pad=0.01,rounding_size=0.03",
                                fc=fc, ec="#888", lw=1.1))
    if isinstance(lines, str):
        lines = [lines]
    n = len(lines); y0 = y + (n-1)*0.18
    for i, ln in enumerate(lines):
        ax.text(x, y0 - i*0.36, ln, ha="center", va="center", fontsize=fs,
                color=tc, fontweight="bold" if (bold_first and i == 0) else "normal")

def arrow(ax, p0, p1, color="#555", style="-|>", ls="-", lw=1.8):
    ax.add_patch(FancyArrowPatch(p0, p1, arrowstyle=style, mutation_scale=15, color=color, ls=ls, lw=lw, shrinkA=4, shrinkB=4))

fig, ax = plt.subplots(1, 2, figsize=(11.5, 5.2))

# ===== (a) 인과·교란 =====
a = ax[0]; a.set_xlim(0, 10); a.set_ylim(0, 10); a.axis("off"); a.set_title("(a)", loc="left", fontsize=12)
box(a, 2.0, 8.4, 3.0, 1.3, ["Graded social", "distancing"], GREEN)
box(a, 7.6, 8.4, 3.3, 1.5, ["How many people", "are in a", "neighbourhood", "(measured mobility)"], MOB, tc="white", fs=8.5)
box(a, 7.6, 3.3, 3.0, 1.2, ["Urban noise", "(S-DoT sensor)"], NOISE, tc="white")
arrow(a, (3.5, 8.4), (5.9, 8.4))
arrow(a, (7.6, 7.6), (7.6, 4.0), color=NOISE, lw=2.2)
a.text(8.0, 5.8, "effect we\nestimate", fontsize=8.5, style="italic", color="#444", ha="left")
# confounders
box(a, 2.0, 4.6, 3.2, 1.3, ["Weather, day of", "week, holidays"], CONF, fs=8.5)
box(a, 2.0, 1.9, 3.2, 1.3, ["Each sensor's offset", "& slow drift over years"], CONF, fs=8.5)
arrow(a, (3.0, 4.4), (6.1, 3.5), color="#999", ls="--", lw=1.3)
arrow(a, (3.2, 2.1), (6.1, 3.1), color="#999", ls="--", lw=1.3)
a.text(5.0, 0.5, "dashed = confounders we must remove (panel b)", fontsize=8, color="#777", ha="center")

# ===== (b) 같은 날 동 사이 비교 =====
b = ax[1]; b.set_xlim(0, 10); b.set_ylim(0, 10); b.axis("off"); b.set_title("(b)", loc="left", fontsize=12)
b.text(5.0, 9.4, "Compare neighbourhoods on the SAME day", fontsize=10, fontweight="bold", ha="center", color="#333")
box(b, 2.6, 7.0, 3.6, 1.9, ["Commercial dong", "people: emptied", "→ quieter?"], "#A9C8C0", fs=9)
box(b, 7.4, 7.0, 3.6, 1.9, ["Residential dong", "people: about normal", "→ baseline"], "#C7C0DA", fs=9)
b.add_patch(FancyArrowPatch((4.5, 7.0), (5.5, 7.0), arrowstyle="<->", mutation_scale=16, color=NOISE, lw=2.2))
b.text(5.0, 7.7, "?", fontsize=14, ha="center", color=NOISE, fontweight="bold")
box(b, 5.0, 3.8, 8.2, 1.9, ["Same day → same weather, same city-wide trend,", "same sensor drift  —  all shared, so they cancel out"], "#EDE8DE", fs=9)
arrow(b, (2.6, 6.0), (3.4, 4.8), color="#999", lw=1.2)
arrow(b, (7.4, 6.0), (6.6, 4.8), color="#999", lw=1.2)
box(b, 5.0, 1.3, 7.4, 1.1, ["Remaining difference = effect of mobility on noise"], GREEN, fs=9.2, tc="white", bold_first=True)
arrow(b, (5.0, 2.85), (5.0, 1.9), color="#777", lw=1.5)

plt.tight_layout()
out = os.path.join(FIG, "fig_identification_concept.png")
plt.savefig(out)
print(f"-> {out}")
