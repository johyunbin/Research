# -*- coding: utf-8 -*-
# Supplementary figures:
#  S1 상업 vs 주거 상세 (지도+산점, 구 Fig 9)  S2 검교정 대조 상세 (산점+offset 분포, 구 calibval)
#  S3 토지이용 분류 지도 (구 Fig 8a) + 분류지표 분포.
import os, csv
import numpy as np
import pandas as pd
from scipy import stats
import statistics as st
import matplotlib.pyplot as plt
from matplotlib.cm import ScalarMappable
from matplotlib.colors import TwoSlopeNorm
from matplotlib.patches import Patch
import figstyle_v2 as FS

FS.apply_style()
PATS, KEYS = FS.load_seoul()
g = pd.read_csv(os.path.join(FS.P, "phase2c_dong_change.csv"), dtype={"dose_key": str})
stat = g.set_index("dose_key")

# ---------- S1: 상업 vs 주거 상세 ----------
nlim = max(abs(g["noise_anom"].quantile(.05)), abs(g["noise_anom"].quantile(.95)))
normN = TwoSlopeNorm(vcenter=0, vmin=-nlim, vmax=nlim); cmN = plt.get_cmap(FS.CMAP_DIV)
fig, axes = plt.subplots(2, 2, figsize=(11.6, 9.6))
for row, lu in enumerate(["commercial", "residential"]):
    sub = stat[stat["landuse"] == lu]
    cols = [cmN(normN(sub.loc[k, "noise_anom"])) if k in sub.index else FS.NA_FILL for k in KEYS]
    axm = axes[row, 0]
    FS.draw_polys(axm, PATS, cols, lw=.25)
    FS.panel_label(axm, "ac"[row], dy=0.03)
    axm.set_title(f"{FS.LANDUSE_LABEL[lu]} dongs (n={len(sub)})", fontsize=9.6)
    cb = plt.colorbar(ScalarMappable(norm=normN, cmap=cmN), ax=axm, fraction=.042, pad=.02, extend="both")
    cb.set_label("ΔL$_{day}$ vs city mean (dB)", fontsize=8.4)
    cb.outline.set_visible(False); cb.ax.tick_params(labelsize=8)

    axs = axes[row, 1]
    d = g[(g["landuse"] == lu) & (g["n_sens"] >= 2)].dropna(subset=["mob_drop", "noise_anom"])
    axs.scatter(d["mob_drop"], d["noise_anom"], s=10 + d["n_sens"] * 4.5, color=FS.LANDUSE[lu],
                alpha=.75, edgecolors="white", linewidths=.3)
    ts = stats.theilslopes(d["noise_anom"].values, d["mob_drop"].values)
    xs = np.linspace(d["mob_drop"].min(), d["mob_drop"].max(), 50)
    axs.plot(xs, ts[1] + ts[0] * xs, color="#1A1A1A", lw=1.8, label="Theil–Sen")
    rp = stats.pearsonr(d["mob_drop"], d["noise_anom"])[0]
    rs = stats.spearmanr(d["mob_drop"], d["noise_anom"])[0]
    axs.axhline(0, color=FS.NEUTRAL, lw=.7, ls="--")
    axs.set_xlim(-20, 20); axs.set_ylim(-5, 5)
    axs.set_xlabel("Daytime mobility reduction (%)")
    axs.set_ylabel("Relative noise change ΔL$_{day}$ (dB)")
    axs.legend(loc="upper left", fontsize=8)
    axs.text(.97, .04, f"Pearson r = {rp:+.2f}\nSpearman ρ = {rs:+.2f}\nTheil–Sen {ts[0]:+.3f} dB/%\nn = {len(d)} dongs (≥2 sensors)",
             transform=axs.transAxes, fontsize=7.8, ha="right", va="bottom",
             bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="#CCCCCC"))
    FS.panel_label(axs, "bd"[row], dy=0.03)
plt.tight_layout()
out = os.path.join(FS.FIG, "figS1_landuse_detail.png")
plt.savefig(out); plt.close(fig)
print("→", out)

# ---------- S2: 검교정 상세 (2022 S-DoT vs 2024 검교정 산점 + offset 분포) ----------
ov = list(csv.DictReader(open(os.path.join(FS.P, "phase5_calibval.csv"), encoding="utf-8-sig")))
offs = [float(r["offset_day"]) for r in ov]
cal = [float(r["cal_day"]) for r in ov]; sdo = [float(r["sdot_day"]) for r in ov]
zones = [r["zone"] for r in ov]
fig, ax = plt.subplots(1, 2, figsize=(11.0, 5.0))
ZC = {"일반": ("Residential/general", "#009E73"), "도로": ("Roadside", FS.NOISE)}
for z, (lab, c) in ZC.items():
    xs = [x for x, zz in zip(cal, zones) if zz == z]; ys = [y for y, zz in zip(sdo, zones) if zz == z]
    ax[0].scatter(xs, ys, s=34, color=c, alpha=.85, edgecolors="white", linewidths=.5, label=lab)
ax[0].plot([40, 80], [40, 80], ls=":", color="#999999", lw=1.1, zorder=0)
ax[0].text(76.5, 78.2, "1:1", fontsize=8, color="#777777", rotation=38)
rp = stats.pearsonr(cal, sdo)[0]; rs = stats.spearmanr(cal, sdo)[0]
ax[0].set_xlim(40, 80); ax[0].set_ylim(40, 80); ax[0].set_box_aspect(1)
ax[0].set_xlabel("Calibrated L$_{Aeq}$ daytime, 2024 (dB)")
ax[0].set_ylabel("Nearby S-DoT L$_{day}$, 2022 (dB)")
ax[0].legend(loc="upper left", fontsize=8.2)
ax[0].text(.04, .66, f"Pearson r = {rp:+.2f}\nSpearman ρ = {rs:+.2f}\nn = {len(ov)} pairs (≤500 m)",
           transform=ax[0].transAxes, fontsize=8,
           bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="#CCCCCC"))
FS.panel_label(ax[0], "a", dy=0.02)
ax[1].hist(offs, bins=18, color="#B9CFE3", edgecolor=FS.MOB, lw=.6)
ax[1].axvline(st.mean(offs), color=FS.NOISE, lw=2.0, label=f"mean {st.mean(offs):+.1f} dB")
ax[1].axvline(0, color=FS.NEUTRAL, lw=.8, ls="--")
ax[1].set_xlabel("S-DoT − calibrated L$_{Aeq}$ offset, daytime (dB)")
ax[1].set_ylabel("Stations"); ax[1].set_box_aspect(1)
ax[1].legend(loc="upper right", fontsize=8.4)
FS.panel_label(ax[1], "b", dy=0.02)
plt.tight_layout()
out = os.path.join(FS.FIG, "figS2_calibration_detail.png")
plt.savefig(out); plt.close(fig)
print("→", out)

# ---------- S3: 토지이용 분류 지도 + 분류지표 분포 ----------
fig = plt.figure(figsize=(11.8, 5.2))
gs = fig.add_gridspec(1, 2, width_ratios=[1.25, 1.0], wspace=0.22)
axm = fig.add_subplot(gs[0]); axh = fig.add_subplot(gs[1])
cols = [FS.LANDUSE[stat.loc[k, "landuse"]] if k in stat.index else FS.NA_FILL for k in KEYS]
FS.draw_polys(axm, PATS, cols, lw=.25)
axm.legend(handles=[Patch(facecolor=FS.LANDUSE[k], edgecolor="white", label=FS.LANDUSE_LABEL[k])
                    for k in ["commercial", "mixed", "residential"]],
           loc="lower left", fontsize=8.4, handlelength=1.2)
FS.add_scalebar(axm, loc=(0.72, 0.05)); FS.add_north(axm)
FS.panel_label(axm, "a", dy=0.03)
q1, q2 = g["comm_idx"].quantile([1/3, 2/3])
for lu in ["commercial", "mixed", "residential"]:
    d = g[g["landuse"] == lu]
    axh.hist(d["comm_idx"], bins=np.linspace(0.5, 3.5, 40), color=FS.LANDUSE[lu],
             alpha=.85, label=FS.LANDUSE_LABEL[lu])
for q in (q1, q2):
    axh.axvline(q, color="#1A1A1A", lw=1.0, ls="--")
axh.set_xlabel("Daytime / night-time de-facto population ratio (post-lifting)")
axh.set_ylabel("Dongs")
axh.legend(loc="upper right", fontsize=8.4)
axh.text(.985, .70, f"tercile cuts:\n{q1:.2f} / {q2:.2f}", transform=axh.transAxes,
         fontsize=8, ha="right", bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="#CCCCCC"))
FS.panel_label(axh, "b", dy=0.03)
plt.tight_layout()
out = os.path.join(FS.FIG, "figS3_landuse_classification.png")
plt.savefig(out); plt.close(fig)
print("→", out)
