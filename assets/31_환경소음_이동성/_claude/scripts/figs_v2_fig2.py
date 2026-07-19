# -*- coding: utf-8 -*-
# Fig 2 (identification): (a) 인과 도식(DAG) (b) 고정효과 사다리.
# v2: 텍스트박스 도해 탈피 — 정돈된 노드·화살표·타이포 위계.
import os
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
import figstyle_v2 as FS

FS.apply_style()

def box(ax, xy, w, h, text, fc, ec, fs=9.5, tc="#1A1A1A", bold=False, lw=1.2):
    b = FancyBboxPatch(xy, w, h, boxstyle="round,pad=0.012,rounding_size=0.018",
                       fc=fc, ec=ec, lw=lw, mutation_aspect=1.0)
    ax.add_patch(b)
    ax.text(xy[0] + w / 2, xy[1] + h / 2, text, ha="center", va="center", fontsize=fs,
            color=tc, fontweight="bold" if bold else "normal", linespacing=1.35)
    return b

def arrow(ax, p0, p1, color, lw=1.8, ls="-", style="-|>", ms=14, rad=0.0):
    a = FancyArrowPatch(p0, p1, arrowstyle=style, mutation_scale=ms, color=color,
                        lw=lw, linestyle=ls, shrinkA=2, shrinkB=2,
                        connectionstyle=f"arc3,rad={rad}")
    ax.add_patch(a)

fig, ax = plt.subplots(1, 2, figsize=(12.6, 4.7), gridspec_kw={"wspace": 0.06})
for a in ax:
    a.set_xlim(0, 1); a.set_ylim(0, 1); a.axis("off")

C_POL = "#DEEBF4"; C_MOB = "#C7DCEC"; C_NOI = "#F7DFD0"; C_CONF = "#EFEFEC"
E_POL = FS.MOB; E_NOI = FS.NOISE; E_CONF = "#9A9A9A"

# ---------- (a) 인과 구조 ----------
FS.panel_label(ax[0], "a")
box(ax[0], (0.02, 0.66), 0.27, 0.24, "Graded social\ndistancing\n(policy, 2020–2022)", C_POL, E_POL)
box(ax[0], (0.375, 0.66), 0.27, 0.24, "Neighbourhood\nmobility\n(de-facto population)", C_MOB, E_POL, bold=True)
box(ax[0], (0.73, 0.30), 0.25, 0.24, "Urban noise\n(S-DoT sensor)", C_NOI, E_NOI, bold=True)
box(ax[0], (0.02, 0.06), 0.27, 0.20, "Weather · day of week\nholidays · season", C_CONF, E_CONF, fs=8.6)
box(ax[0], (0.375, 0.06), 0.27, 0.20, "Sensor calibration offset\n+ multi-year drift", C_CONF, E_CONF, fs=8.6)
arrow(ax[0], (0.29, 0.78), (0.375, 0.78), "#444444", lw=1.6)
arrow(ax[0], (0.645, 0.72), (0.755, 0.54), FS.NOISE, lw=2.6)
ax[0].text(0.555, 0.475, "dose–response β\n(effect we estimate)", fontsize=8.8, color=FS.NOISE,
           ha="center", fontstyle="italic", fontweight="bold")
arrow(ax[0], (0.155, 0.26), (0.72, 0.37), E_CONF, lw=1.3, ls=(0, (4, 3)), rad=-0.18)
arrow(ax[0], (0.645, 0.16), (0.76, 0.30), E_CONF, lw=1.3, ls=(0, (4, 3)))
ax[0].text(0.5, -0.045, "dashed = confounding and measurement artefacts removed by the design in (b)",
           fontsize=8.2, color="#666666", ha="center")

# ---------- (b) 고정효과 사다리 ----------
FS.panel_label(ax[1], "b")
steps = [
    ("Raw S-DoT level", "contaminated by sensor offset, drift,\nweather, city-wide trends", C_CONF, E_CONF, False),
    ("+ Sensor fixed effects", "removes each sensor's\ncalibration offset", C_MOB, E_POL, False),
    ("+ Date fixed effects", "removes weather, calendar, city trend,\nnetwork-wide drift (all date-common)", C_MOB, E_POL, False),
]
y0, h, gap = 0.79, 0.17, 0.065
for i, (t, note, fc, ec, bold) in enumerate(steps):
    y = y0 - i * (h + gap)
    box(ax[1], (0.02, y), 0.42, h, t, fc, ec, fs=9.3, bold=True)
    ax[1].text(0.47, y + h / 2, note, fontsize=8.4, va="center", color="#555555", linespacing=1.3)
    arrow(ax[1], (0.23, y - 0.005), (0.23, y - gap + 0.005), "#444444", lw=1.5)
box(ax[1], (0.02, 0.045), 0.62, 0.175,
    "Identifying variation:\nsame-day differences in mobility\nacross neighbourhoods",
    "#E4F0E9", "#009E73", fs=9.0, bold=True)

plt.tight_layout()
out = os.path.join(FS.FIG, "fig2_identification.png")
plt.savefig(out)
print("→", out)
