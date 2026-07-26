# -*- coding: utf-8 -*-
# Phase 2-F 토지이용별 공간패턴 — 2x2: 상업 동(지도+산점도) vs 주거 동(지도+산점도).
# 상업 동에서 (이동량감소 vs 상대소음변화) 관계가 더 뚜렷한지 대비. 파스텔·Arial·지리비율.
import os, json
import numpy as np
import pandas as pd
from scipy import stats
import matplotlib.pyplot as plt
from matplotlib.collections import PolyCollection
from matplotlib.cm import ScalarMappable
from matplotlib.colors import Normalize
import figstyle as FS

FS.apply_style()
MIN_SENS = 2   # 추세·상관은 센서 2개 이상 동만(단일센서=드리프트 노이즈)
ROOT = r"T:\00_BACKUP\01_한양대학교\00_연구실\00_프로젝트\★_논문화 프로젝트\31_환경소음_이동성"
P = os.path.join(ROOT, "data", "processed")
FIG = os.path.join(ROOT, "_claude", "figures")
GEO = os.path.join(ROOT, "data", "reference", "admdong_seoul_ver20220101.geojson")

g = pd.read_csv(os.path.join(P, "phase2c_dong_change.csv"), dtype={"dose_key": str})
stat = g.set_index("dose_key")
REV = {"11740525": "11740520", "11740526": "11740520"}
geo = json.load(open(GEO, encoding="utf-8"))
def polys_of(gm):
    return gm["coordinates"] if gm["type"] == "MultiPolygon" else [gm["coordinates"]]
PATS, KEYS = [], []
for f in geo["features"]:
    code = str(f["properties"]["adm_cd2"])[:8]
    key = REV.get(code, code)
    for poly in polys_of(f["geometry"]):
        PATS.append(np.array(poly[0])); KEYS.append(key)
NLIM = max(abs(g["noise_anom"].quantile(.05)), abs(g["noise_anom"].quantile(.95)))

def small_map(ax, lu, title):
    sub = set(stat.index[stat["landuse"] == lu])
    norm = Normalize(-NLIM, NLIM); cm = FS.PASTEL_DIV
    cols = [cm(norm(stat.loc[k, "noise_anom"])) if (k in sub) else "#F3F3F1" for k in KEYS]
    ax.add_collection(PolyCollection(PATS, facecolors=cols, edgecolors=FS.EDGE, linewidths=.25))
    ax.autoscale(); FS.style_map_ax(ax); ax.set_title(title)
    sm = ScalarMappable(norm=norm, cmap=cm); sm.set_array([])
    cb = plt.colorbar(sm, ax=ax, fraction=.043, pad=.02); cb.set_label("ΔL$_{day}$ vs city mean (dB)", fontsize=8.5)
    cb.outline.set_visible(False)

def scatter_group(ax, lu, title):
    d = g[(g["landuse"] == lu) & (g["n_sens"] >= MIN_SENS)].dropna(subset=["mob_drop", "noise_anom"])
    c = FS.LANDUSE_COLORS[lu]
    ax.scatter(d["mob_drop"], d["noise_anom"], s=10 + d["n_sens"] * 6, color=c, alpha=.75,
               edgecolors="white", linewidths=.4)
    txt = ""
    if len(d) > 5 and d["mob_drop"].std() > 1e-6:
        ts = stats.theilslopes(d["noise_anom"].values, d["mob_drop"].values)   # robust 회귀
        xs = np.linspace(d["mob_drop"].min(), d["mob_drop"].max(), 50)
        ax.plot(xs, ts[1] + ts[0] * xs, color=c, lw=2)
        rho = stats.spearmanr(d["mob_drop"], d["noise_anom"])[0]
        pr = stats.pearsonr(d["mob_drop"], d["noise_anom"])[0]
        txt = f"Theil-Sen={ts[0]:+.3f} dB/%\nSpearman ρ={rho:+.2f}  (Pearson r={pr:+.2f})\nn={len(d)} dongs (≥{MIN_SENS} sensors)"
        print(f"  {lu}: Theil-Sen={ts[0]:+.4f} Spearman={rho:+.3f} Pearson={pr:+.3f} n={len(d)}")
    ax.axhline(0, color=FS.ACCENT["neutral"], lw=.7, ls="--")
    ax.text(.04, .04, txt, transform=ax.transAxes, fontsize=8, va="bottom",
            bbox=dict(boxstyle="round,pad=0.3", fc="white", ec=c, alpha=.85))
    ax.set_xlabel("Daytime mobility reduction (%)")
    ax.set_ylabel("Relative noise change ΔL$_{day}$ (dB)")
    ax.set_title(title)
    ax.set_xlim(-20, 20); ax.set_ylim(-5, 5)   # 이상치(소수 동) 제외하고 본체에 맞춰 — 회귀·상관은 전체 데이터 기준

print("=== 토지이용 그룹별 산점 회귀 ===")
fig, ax = plt.subplots(2, 2, figsize=(11.4, 10.0))
small_map(ax[0, 0], "commercial", "(a)")
scatter_group(ax[0, 1], "commercial", "(b)")
small_map(ax[1, 0], "residential", "(c)")
scatter_group(ax[1, 1], "residential", "(d)")
# suptitle 제거(캡션이 대신함)
plt.tight_layout(rect=[0, 0, 1, 0.98])
out = os.path.join(FIG, "fig_spatial_bylanduse.png")
plt.savefig(out)
print(f"-> {out}")
