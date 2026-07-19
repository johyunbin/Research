# -*- coding: utf-8 -*-
# Fig 8 (구 Fig 10, drift validation): (a) 연추세 S-DoT vs 검교정망 (b) offset 산점.
# v2: 라벨 겹침 수정·드리프트 주석 화살표·1:1선 라벨·고대비.
import os, csv
import statistics as st
import pandas as pd
import matplotlib.pyplot as plt
import figstyle_v2 as FS

FS.apply_style()
tr = pd.read_csv(os.path.join(FS.P, "phase5_calibration_trends.csv")).set_index("series")
ov = list(csv.DictReader(open(os.path.join(FS.P, "phase5_calibval.csv"), encoding="utf-8-sig")))
offs = [float(r["offset_day"]) for r in ov]

YRS = ["2020", "2021", "2022", "2023"]; yy = [int(y) for y in YRS]
LINES = [("S-DoT (this network)", "sdot", FS.NOISE, "o"),
         ("Calibrated automatic, road", "cal_auto_road", "#009E73", "s"),
         ("Calibrated manual, general", "cal_manual_general", FS.MOB, "^"),
         ("Calibrated manual, road", "cal_manual_road", "#8C6BB1", "D")]

fig, ax = plt.subplots(1, 2, figsize=(11.4, 5.1))

# (a) 연추세 (2020 기준 변화)
for name, key, c, mk in LINES:
    d = [tr.loc[key, y] - tr.loc[key, "2020"] for y in YRS]
    ax[0].plot(yy, d, marker=mk, color=c, lw=1.9, ms=6.5, mfc="white", mew=1.6, label=name)
ax[0].axhline(0, color=FS.NEUTRAL, lw=.8, ls="--")
ax[0].set_xticks(yy)
ax[0].set_xlabel("Year"); ax[0].set_ylabel("Annual level change from 2020 (dB)")
ax[0].set_ylim(-2.75, 2.6)
ax[0].legend(loc="upper center", bbox_to_anchor=(0.5, -0.16), ncol=2, fontsize=8,
             columnspacing=1.2, handletextpad=0.5)
ax[0].set_box_aspect(1)
d23 = tr.loc["sdot", "2023"] - tr.loc["sdot", "2020"]
ax[0].annotate(f"{d23:+.1f} dB\n(sensor drift)", xy=(2023, d23), xytext=(2022.05, -1.35),
               fontsize=8.6, color=FS.NOISE, fontweight="bold", ha="center",
               arrowprops=dict(arrowstyle="->", color=FS.NOISE, lw=1.2))
FS.panel_label(ax[0], "a", dy=0.02)

# (b) offset 산점
ZC = {"일반": ("Residential/general", "#009E73"), "도로": ("Roadside", FS.NOISE)}
for z, (lab, c) in ZC.items():
    xs = [float(r["cal_day"]) for r in ov if r["zone"] == z]
    ys = [float(r["sdot_day"]) for r in ov if r["zone"] == z]
    ax[1].scatter(xs, ys, s=34, color=c, alpha=.85, edgecolors="white", linewidths=.5, label=lab)
ax[1].plot([40, 80], [40, 80], ls=":", color="#999999", lw=1.1, zorder=0)
ax[1].text(76.5, 78.2, "1:1", fontsize=8, color="#777777", rotation=38)
ax[1].set_xlim(40, 80); ax[1].set_ylim(40, 80); ax[1].set_box_aspect(1)
ax[1].set_xlabel("Calibrated L$_{Aeq}$, daytime (dB)")
ax[1].set_ylabel("Nearby S-DoT L$_{day}$ (dB)")
ax[1].legend(loc="upper left", fontsize=8.4)
ax[1].text(.04, .70, f"S-DoT reads {st.mean(offs):+.1f} dB on average\nvs calibrated stations (n = {len(offs)})",
           transform=ax[1].transAxes, fontsize=8.2,
           bbox=dict(boxstyle="round,pad=0.32", fc="white", ec="#CCCCCC"))
FS.panel_label(ax[1], "b", dy=0.02)

plt.tight_layout()
out = os.path.join(FS.FIG, "fig8_drift_validation.png")
plt.savefig(out)
print(f"S-DoT Δ={d23:+.2f} dB · mean offset {st.mean(offs):+.2f} dB (n={len(offs)}) → {out}")
