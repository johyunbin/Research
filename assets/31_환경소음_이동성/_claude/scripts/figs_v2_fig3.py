# -*- coding: utf-8 -*-
# Fig 3 (dose-response 통합): (a) 12개 분할추정 grouped forest(단일 축·FDR 마커·우측 수치열)
#                             (b) sensor FE vs two-way FE 비교(dot-CI) — 구 Fig 4b 흡수.
import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
import figstyle_v2 as FS

FS.apply_style()
seg = pd.read_csv(os.path.join(FS.P, "phase3_segmented_fdr.csv"))
dn = pd.read_csv(os.path.join(FS.P, "phase2g_daynight_results.csv"))

GROUPS = [
    ("By outcome", [("Daytime L_day", "Daytime  L$_{day}$"),
                    ("Nighttime L_night", "Night-time  L$_{night}$"),
                    ("Day-night gap", "Day–night gap")]),
    ("By land use (daytime)", [("Commercial", "Commercial"), ("Mixed", "Mixed"),
                               ("Residential", "Residential")]),
    ("By day type (daytime)", [("Weekday", "Weekday"), ("Weekend/holiday", "Weekend / holiday")]),
    ("By season (daytime)", [("DJF", "Winter (DJF)"), ("MAM", "Spring (MAM)"),
                             ("JJA", "Summer (JJA)"), ("SON", "Autumn (SON)")]),
]
rows = []
for gname, items in GROUPS:
    rows.append(("HDR", gname, None))
    for key, lab in items:
        r = seg[seg["group"] == key].iloc[0]
        rows.append(("EST", lab, r))

fig = plt.figure(figsize=(13.4, 6.4))
gs = fig.add_gridspec(1, 2, width_ratios=[1.55, 1.0], wspace=0.52)
axA = fig.add_subplot(gs[0]); axB = fig.add_subplot(gs[1])

# ---------- (a) grouped forest ----------
ys = np.arange(len(rows))[::-1]
for y, (kind, lab, r) in zip(ys, rows):
    if kind == "HDR":
        axA.text(-0.02, y, lab, transform=axA.get_yaxis_transform(), fontsize=9.3,
                 fontweight="bold", va="center", ha="right", color="#1A1A1A")
        continue
    b, se, p, fdr = r["beta"], r["se"], r["pval"], bool(r["sig_FDR_05"])
    lo, hi = b - 1.96 * se, b + 1.96 * se
    if fdr:
        c, mfc = FS.SIG, FS.SIG
    elif p < 0.05:
        c, mfc = FS.SIG, "white"
    else:
        c, mfc = FS.NS, "white"
    axA.errorbar(b, y, xerr=[[b - lo], [hi - b]], fmt="o", color=c, mfc=mfc, mec=c,
                 ms=6.5, capsize=3.2, lw=1.6, mew=1.6, zorder=3)
    star = "**" if fdr else ("*" if p < 0.05 else "")
    axA.text(1.015, y, f"{b:+.2f} [{lo:+.2f}, {hi:+.2f}]{star}",
             transform=axA.get_yaxis_transform(), fontsize=8.2, va="center",
             family="Arial", color="#1A1A1A")
axA.axvline(0, color=FS.NEUTRAL, lw=.9, ls="--", zorder=1)
axA.set_yticks([y for y, (k, _, _) in zip(ys, rows) if k == "EST"])
axA.set_yticklabels([lab for k, lab, _ in rows if k == "EST"], fontsize=9)
axA.set_ylim(-0.8, len(rows) - 0.2)
axA.set_xlim(-1.4, 2.4)
axA.set_xlabel("Dose–response β (dB per log-unit mobility)")
axA.text(1.015, len(rows) - 0.2, "β [95% CI]", transform=axA.get_yaxis_transform(),
         fontsize=8.2, fontweight="bold", va="bottom")
FS.panel_label(axA, "a", dy=0.03)
handles = [Line2D([0], [0], marker="o", ls="", color=FS.SIG, mfc=FS.SIG, ms=6.5, label="FDR < 0.05"),
           Line2D([0], [0], marker="o", ls="", color=FS.SIG, mfc="white", mew=1.6, ms=6.5, label="p < 0.05 (nominal)"),
           Line2D([0], [0], marker="o", ls="", color=FS.NS, mfc="white", mew=1.6, ms=6.5, label="Not significant")]
axA.legend(handles=handles, loc="lower left", fontsize=8, handletextpad=0.15, borderaxespad=0.2)
axA.spines["left"].set_visible(False)
axA.tick_params(axis="y", length=0)

# ---------- (b) sensor FE vs two-way FE ----------
comp = [("Daytime L_day", "Daytime\nL$_{day}$"), ("Nighttime L_night", "Night-time\nL$_{night}$")]
xpos = np.array([0.0, 1.0]); off = 0.16
for j, (key, lab) in enumerate(comp):
    r = dn.iloc[j]
    for k, (b, se, c, lbl, mk) in enumerate([
            (r["b_within"], r["se_within"], "#9A9A9A", "Sensor FE (+ weather/calendar)", "s"),
            (r["b_2way"], r["se_2way"], FS.SIG, "Two-way FE (sensor + date)", "o")]):
        x = xpos[j] + (k - 0.5) * 2 * off
        lo, hi = b - 1.96 * se, b + 1.96 * se
        axB.errorbar(x, b, yerr=[[b - lo], [hi - b]], fmt=mk, color=c, mfc=c if k else "white",
                     ms=7, capsize=4, lw=1.8, mew=1.7,
                     label=lbl if j == 0 else None, zorder=3)
        p = r["p_within"] if k == 0 else r["p_2way"]
        star = "*" if p < 0.05 else " (ns)"
        axB.annotate(f"{b:+.2f}{star}", (x, hi), textcoords="offset points", xytext=(0, 5),
                     ha="center", fontsize=8.4, color=c)
axB.axhline(0, color=FS.NEUTRAL, lw=.9, ls="--", zorder=1)
axB.set_xticks(xpos); axB.set_xticklabels([lab for _, lab in comp], fontsize=9.5)
axB.set_xlim(-0.55, 1.55)
axB.set_ylim(-1.2, 4.4)
axB.set_ylabel("β (dB per log-unit mobility)")
axB.legend(loc="upper left", fontsize=8.4)
axB.text(-0.42, 2.45, "night-time association loses\nsignificance under date FE\n(day-night difference\nnot established)",
         fontsize=8, color="#555555", ha="left", linespacing=1.35)
FS.panel_label(axB, "b", dy=0.03)

plt.tight_layout()
out = os.path.join(FS.FIG, "fig3_doseresponse.png")
plt.savefig(out)
print("→", out)
