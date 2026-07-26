# -*- coding: utf-8 -*-
# small-multiples: 거리두기 국면별 동 주간 이동량(lp_day_rel) 지도 (논문1 Fig2-3 '차원×시간 행렬' 패턴).
# spatiotemporal dose 진화 — 상업 도심이 비었다가(2020-12·2021-07) 회복(위드코로나·해제).
import os, json
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.collections import PolyCollection
from matplotlib.cm import ScalarMappable
from matplotlib.colors import TwoSlopeNorm
import figstyle as FS

FS.apply_style()
ROOT = r"T:\00_BACKUP\01_한양대학교\00_연구실\00_프로젝트\★_논문화 프로젝트\31_환경소음_이동성"
P = os.path.join(ROOT, "data", "processed")
FIG = os.path.join(ROOT, "_claude", "figures")
GEO = os.path.join(ROOT, "data", "reference", "admdong_seoul_ver20220101.geojson")

# 국면 윈도우(수도권)
PHASES = [
    ("2020-12  2.5-tier + 5-person ban", "2020-12-08", "2021-02-14"),
    ("2021-07  Level 4 (capital area)", "2021-07-12", "2021-09-30"),
    ("2021-11  With-COVID (relaxation)", "2021-11-01", "2021-12-05"),
    ("2022-03  just before full lift", "2022-03-01", "2022-04-17"),
]
lp = pd.read_csv(os.path.join(P, "livingpop_dong_daily_2020-2023.csv"), dtype={"adm_cd": str})
lp["date"] = pd.to_datetime(lp["date"])
XW = {"11305590": "11305595", "11305600": "11305603", "11305606": "11305608",
      "11305610": "11305615", "11305620": "11305625", "11305630": "11305635",
      "11740525": "11740520", "11740526": "11740520"}
lp["dose_key"] = lp["adm_cd"].map(lambda c: XW.get(c, c))
agg = lp.groupby(["dose_key", "date"], as_index=False)["lp_day"].sum()
base = agg[(agg["date"] >= "2022-07-01") & (agg["date"] <= "2023-12-31")].groupby("dose_key")["lp_day"].mean()

geo = json.load(open(GEO, encoding="utf-8"))
def polys_of(gm):
    return gm["coordinates"] if gm["type"] == "MultiPolygon" else [gm["coordinates"]]
REV = {"11740525": "11740520", "11740526": "11740520"}
PATS, KEYS = [], []
for f in geo["features"]:
    code = str(f["properties"]["adm_cd2"])[:8]; key = REV.get(code, code)
    for poly in polys_of(f["geometry"]):
        PATS.append(np.array(poly[0])); KEYS.append(key)

norm = TwoSlopeNorm(vcenter=1.0, vmin=0.7, vmax=1.15)   # 1=baseline, <1 emptied
cm = FS.PASTEL_DIV   # 비움(낮음)=teal(cool/조용), 증가(높음)=rose(warm/붐빔) — 직관적
fig, axes = plt.subplots(2, 2, figsize=(10.0, 9.0))   # 2x2 (Word 단 너비 적합)
ax = axes.flat
for k, (lab, d0, d1) in enumerate(PHASES):
    win = agg[(agg["date"] >= d0) & (agg["date"] <= d1)].groupby("dose_key")["lp_day"].mean()
    rel = (win / base).reindex(base.index)
    cols = [cm(norm(rel.get(key))) if key in rel.index and pd.notna(rel.get(key)) else FS.NA_FILL for key in KEYS]
    ax[k].add_collection(PolyCollection(PATS, facecolors=cols, edgecolors="white", linewidths=.2))
    ax[k].autoscale(); FS.style_map_ax(ax[k]); ax[k].set_title(lab, fontsize=10)
    ax[k].text(.5, -.02, f"median {rel.median():.2f}× baseline", transform=ax[k].transAxes, ha="center", fontsize=8, color="#555")
sm = ScalarMappable(norm=norm, cmap=cm); sm.set_array([])
cb = fig.colorbar(sm, ax=axes, fraction=.030, pad=.02)
cb.set_label("Daytime mobility relative to post-lift baseline", fontsize=9); cb.outline.set_visible(False)
# suptitle 제거(캡션이 대신함)
out = os.path.join(FIG, "fig_phase_mobility_maps.png")
plt.savefig(out)
print("국면별 median 이동량(×baseline):")
for lab, d0, d1 in PHASES:
    win = agg[(agg["date"] >= d0) & (agg["date"] <= d1)].groupby("dose_key")["lp_day"].mean()
    print(f"  {lab[:28]}: {(win/base).median():.3f} (min동 {(win/base).min():.2f})")
print(f"-> {out}")
