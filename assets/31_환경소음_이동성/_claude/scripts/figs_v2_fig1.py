# -*- coding: utf-8 -*-
# Fig 1 (study area): (a) S-DoT 센서망 + 토지이용 분류 (b) 동별 기준 주간소음(viridis).
# v2: 고대비 색·스케일바·방위·패널라벨·범례 확대.
import os
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.cm import ScalarMappable
from matplotlib.colors import Normalize
import figstyle_v2 as FS

FS.apply_style()
PATS, KEYS = FS.load_seoul()

sd = pd.read_csv(os.path.join(FS.CACHE, "sensor_lu.csv"), dtype={"serial": str, "adm_cd": str})
base = pd.read_csv(os.path.join(FS.CACHE, "dong_base_leq.csv"), dtype={"dose_key": str}).set_index("dose_key")["base_Lday"]

fig, ax = plt.subplots(1, 2, figsize=(13.2, 5.6))

# (a) 센서망 + 토지이용
ax[0].add_collection(plt.matplotlib.collections.PolyCollection(
    PATS, facecolors="#F2F1ED", edgecolors="#FFFFFF", linewidths=.4))
for lu in ["mixed", "residential", "commercial"]:
    d = sd[sd["lu"] == lu]
    ax[0].scatter(d["lon"], d["lat"], s=13, color=FS.LANDUSE[lu], alpha=.9,
                  edgecolors="white", linewidths=.25,
                  label=f"{FS.LANDUSE_LABEL[lu]} (n={len(d)})")
ax[0].autoscale(); FS.style_map_ax(ax[0])
leg = ax[0].legend(loc="lower left", fontsize=9, markerscale=1.7, handletextpad=0.25,
                   borderaxespad=0.1, labelspacing=0.35)
FS.panel_label(ax[0], "a")
FS.add_scalebar(ax[0]); FS.add_north(ax[0])

# (b) 기준 주간소음 choropleth
vmin, vmax = base.quantile(.05), base.quantile(.95)
norm = Normalize(vmin, vmax); cm = plt.get_cmap(FS.CMAP_SEQ_NOISE)
cols = [cm(norm(base[k])) if k in base.index else FS.NA_FILL for k in KEYS]
FS.draw_polys(ax[1], PATS, cols, lw=.35)
FS.panel_label(ax[1], "b")
sm = ScalarMappable(norm=norm, cmap=cm); sm.set_array([])
cb = plt.colorbar(sm, ax=ax[1], fraction=.042, pad=.02, extend="both")
cb.set_label("Baseline daytime level L$_{day}$ (dB)", fontsize=9.5)
cb.outline.set_visible(False); cb.ax.tick_params(labelsize=8.5)

plt.tight_layout()
out = os.path.join(FS.FIG, "fig1_study_area.png")
plt.savefig(out)
print(f"센서 {len(sd)} | " + " ".join(f"{k}:{v}" for k, v in sd.lu.value_counts().items()))
print(f"기준 Lday p05-95 {vmin:.1f}~{vmax:.1f} dB → {out}")
