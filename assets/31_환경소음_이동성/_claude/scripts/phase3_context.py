# -*- coding: utf-8 -*-
# Fig 1 (맥락): (a) S-DoT 센서망 + 활동유형(주간/야간 인구비) 분류  (b) 기준 소음수준(동별 post-lift Leq_day).
# 논문1(B&E) Fig1 '연구지역+기능별 zone' 패턴 차용.
import os, json
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.collections import PolyCollection
from matplotlib.cm import ScalarMappable
from matplotlib.colors import Normalize
from matplotlib.patches import Patch
import figstyle as FS

FS.apply_style()
ROOT = r"T:\00_BACKUP\01_한양대학교\00_연구실\00_프로젝트\★_논문화 프로젝트\31_환경소음_이동성"
P = os.path.join(ROOT, "data", "processed")
FIG = os.path.join(ROOT, "_claude", "figures")
GEO = os.path.join(ROOT, "data", "reference", "admdong_seoul_ver20220101.geojson")

sd = pd.read_csv(os.path.join(P, "sensor_dong_map.csv"), dtype={"serial": str, "adm_cd": str})
sd = sd[sd["matched"] == 1]
pan = pd.read_csv(os.path.join(P, "analysis_panel.csv"), dtype={"serial": str, "dose_key": str},
                  usecols=["serial", "dose_key", "date", "Leq_day", "lp_day", "lp_night"])
pan["date"] = pd.to_datetime(pan["date"])
post = pan[(pan["date"] >= "2022-07-01") & (pan["date"] <= "2023-12-31")]
comm = post.groupby("dose_key").apply(lambda g: g["lp_day"].mean() / g["lp_night"].mean(), include_groups=False)
q1, q2 = comm.quantile([1/3, 2/3])
def lu(c):
    return "commercial" if c >= q2 else ("residential" if c <= q1 else "mixed")
sd["lu"] = sd["adm_cd"].map(lambda c: lu(comm.get(c, np.nan)) if pd.notna(comm.get(c, np.nan)) else "mixed")
base_leq = post.groupby("dose_key")["Leq_day"].mean()

geo = json.load(open(GEO, encoding="utf-8"))
def polys_of(gm):
    return gm["coordinates"] if gm["type"] == "MultiPolygon" else [gm["coordinates"]]
REV = {"11740525": "11740520", "11740526": "11740520"}
PATS, KEYS = [], []
for f in geo["features"]:
    code = str(f["properties"]["adm_cd2"])[:8]; key = REV.get(code, code)
    for poly in polys_of(f["geometry"]):
        PATS.append(np.array(poly[0])); KEYS.append(key)

fig, ax = plt.subplots(1, 2, figsize=(13.5, 5.4))
# (a) 센서망 + 활동유형
ax[0].add_collection(PolyCollection(PATS, facecolors="#F4F3EF", edgecolors="white", linewidths=.35))
for luk in ["residential", "mixed", "commercial"]:
    d = sd[sd["lu"] == luk]
    ax[0].scatter(d["lon"], d["lat"], s=9, color=FS.LANDUSE_COLORS[luk], alpha=.82,
                  edgecolors="white", linewidths=.18, label=f"{FS.LANDUSE_LABEL[luk].split(' (')[0]} (n={len(d)})")
ax[0].autoscale(); FS.style_map_ax(ax[0])
ax[0].set_title("(a)")
ax[0].legend(loc="lower left", fontsize=7.8, handlelength=1.0)
# (b) 기준 소음수준 (동별 post-lift Leq_day)
vmin, vmax = base_leq.quantile(.05), base_leq.quantile(.95)
norm = Normalize(vmin, vmax); cm = plt.get_cmap("YlGnBu")
cols = [cm(norm(base_leq.get(k))) if k in base_leq.index else FS.NA_FILL for k in KEYS]
ax[1].add_collection(PolyCollection(PATS, facecolors=cols, edgecolors="white", linewidths=.35))
ax[1].autoscale(); FS.style_map_ax(ax[1])
ax[1].set_title("(b)")
smm = ScalarMappable(norm=norm, cmap=cm); smm.set_array([])
cb = plt.colorbar(smm, ax=ax[1], fraction=.043, pad=.02); cb.set_label("L$_{day}$ (dB)", fontsize=9); cb.outline.set_visible(False)
# suptitle 제거(캡션이 대신함)
plt.tight_layout()
out = os.path.join(FIG, "fig_study_area.png")
plt.savefig(out)
print(f"센서 {len(sd)} | 상업 {(sd.lu=='commercial').sum()} 혼합 {(sd.lu=='mixed').sum()} 주거 {(sd.lu=='residential').sum()}")
print(f"기준 Leq_day 범위 {base_leq.min():.1f}~{base_leq.max():.1f} (p05-95 {vmin:.1f}~{vmax:.1f})")
print(f"-> {out}")
