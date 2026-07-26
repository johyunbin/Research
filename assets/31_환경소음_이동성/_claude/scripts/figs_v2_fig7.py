# -*- coding: utf-8 -*-
# Fig 7 (구 Fig 8+9 통합, spatial null): (a) 이동량 감소 지도 (b) 상대 소음변화 지도
#  (c) 산점(토지이용 색·Theil-Sen·Spearman vs Pearson 대비). 상세 분리는 Supp S1.
import os
import numpy as np
import pandas as pd
from scipy import stats
import matplotlib.pyplot as plt
from matplotlib.cm import ScalarMappable
from matplotlib.colors import Normalize, TwoSlopeNorm
import figstyle_v2 as FS

FS.apply_style()
PATS, KEYS = FS.load_seoul()
g = pd.read_csv(os.path.join(FS.P, "phase2c_dong_change.csv"), dtype={"dose_key": str})
stat = g.set_index("dose_key")

# 인쇄 실크기(6.5in) — 지도 2매 상단 병렬 + 산점 하단 전폭
fig = plt.figure(figsize=(6.5, 7.4))
gs = fig.add_gridspec(2, 2, height_ratios=[1.0, 1.05], hspace=0.24, wspace=0.30)
axA = fig.add_subplot(gs[0, 0]); axB = fig.add_subplot(gs[0, 1]); axC = fig.add_subplot(gs[1, :])

# (a) 이동량 감소율 지도
vmax = max(8, g["mob_drop"].quantile(.95))
normA = Normalize(0, vmax); cmA = plt.get_cmap(FS.CMAP_SEQ_DROP)
cols = [cmA(normA(stat.loc[k, "mob_drop"])) if k in stat.index else FS.NA_FILL for k in KEYS]
FS.draw_polys(axA, PATS, cols, lw=.25)
FS.panel_label(axA, "a", dy=0.03)
FS.add_scalebar(axA, loc=(0.02, 0.05))
cb = plt.colorbar(ScalarMappable(norm=normA, cmap=cmA), ax=axA, fraction=.042, pad=.02, extend="max")
cb.set_label("Daytime mobility reduction (%)", fontsize=8.6)
cb.outline.set_visible(False); cb.ax.tick_params(labelsize=8)

# (b) 상대 소음변화 지도
nlim = max(abs(g["noise_anom"].quantile(.05)), abs(g["noise_anom"].quantile(.95)))
normB = TwoSlopeNorm(vcenter=0, vmin=-nlim, vmax=nlim); cmB = plt.get_cmap(FS.CMAP_DIV)
cols = [cmB(normB(stat.loc[k, "noise_anom"])) if k in stat.index else FS.NA_FILL for k in KEYS]
FS.draw_polys(axB, PATS, cols, lw=.25)
FS.panel_label(axB, "b", dy=0.03)
cb = plt.colorbar(ScalarMappable(norm=normB, cmap=cmB), ax=axB, fraction=.042, pad=.02, extend="both")
cb.set_label("ΔL$_{day}$ vs city mean (dB)", fontsize=8.6)
cb.outline.set_visible(False); cb.ax.tick_params(labelsize=8)

# (c) 산점 (센서≥2 동)
gg = g[g["n_sens"] >= 2].dropna(subset=["mob_drop", "noise_anom"])
for lu in ["mixed", "residential", "commercial"]:
    d = gg[gg["landuse"] == lu]
    axC.scatter(d["mob_drop"], d["noise_anom"], s=10 + d["n_sens"] * 4.5,
                color=FS.LANDUSE[lu], alpha=.75, edgecolors="white", linewidths=.3,
                label=f"{FS.LANDUSE_LABEL[lu]} (n={len(d)})")
ts = stats.theilslopes(gg["noise_anom"].values, gg["mob_drop"].values)
xs = np.linspace(gg["mob_drop"].min(), gg["mob_drop"].max(), 50)
axC.plot(xs, ts[1] + ts[0] * xs, color="#1A1A1A", lw=2.0, label="Theil–Sen (all)")
axC.axhline(0, color=FS.NEUTRAL, lw=.7, ls="--")
axC.set_xlim(-20, 20); axC.set_ylim(-5, 5)
axC.set_xlabel("Daytime mobility reduction (%)")
axC.set_ylabel("Relative noise change ΔL$_{day}$ (dB)", labelpad=1)
axC.legend(loc="upper left", fontsize=7.4, handletextpad=0.2, labelspacing=0.3)
FS.panel_label(axC, "c", dy=0.03)

rho_all = stats.spearmanr(gg["mob_drop"], gg["noise_anom"])[0]
com = gg[gg["landuse"] == "commercial"]
r_p = stats.pearsonr(com["mob_drop"], com["noise_anom"])[0]
r_s = stats.spearmanr(com["mob_drop"], com["noise_anom"])[0]
axC.text(0.975, 0.035,
         f"all dongs: Spearman ρ = {rho_all:+.2f} · Theil–Sen {ts[0]:+.3f} dB/%\n"
         f"commercial: Pearson r = {r_p:+.2f} vs Spearman ρ = {r_s:+.2f}\n"
         f"(Pearson inflated by a few extreme dongs)",
         transform=axC.transAxes, fontsize=7.4, ha="right", va="bottom", color="#444444",
         bbox=dict(boxstyle="round,pad=0.32", fc="white", ec="#CCCCCC", alpha=.92))

plt.tight_layout()
out = os.path.join(FS.FIG, "fig7_spatial_null.png")
plt.savefig(out)
print(f"동 {len(g)} (산점 {len(gg)}) · 전체 ρ={rho_all:+.3f} · 상업 Pearson {r_p:+.2f} vs Spearman {r_s:+.2f}")
print("→", out)
