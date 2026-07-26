# -*- coding: utf-8 -*-
# Fig 4 (구 Fig 5, phase mobility maps): 국면별 동 주간 이동량 지도 4매.
# v2: RdBu_r 고대비·1.0 중심 발산·패널라벨·간결 제목·가로 공용 컬러바·스케일바.
import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.cm import ScalarMappable
from matplotlib.colors import TwoSlopeNorm
import figstyle_v2 as FS

FS.apply_style()
PATS, KEYS = FS.load_seoul()

PHASES = [
    ("a", "Dec 2020 — Level 2.5 + 5-person ban", "2020-12-08", "2021-02-14"),
    ("b", "Jul 2021 — Level 4 (strongest)", "2021-07-12", "2021-09-30"),
    ("c", "Nov 2021 — With-COVID relaxation", "2021-11-01", "2021-12-05"),
    ("d", "Mar 2022 — just before full lifting", "2022-03-01", "2022-04-17"),
]
XW = {"11305590": "11305595", "11305600": "11305603", "11305606": "11305608",
      "11305610": "11305615", "11305620": "11305625", "11305630": "11305635",
      "11740525": "11740520", "11740526": "11740520"}
lp = pd.read_csv(os.path.join(FS.P, "livingpop_dong_daily_2020-2023.csv"), dtype={"adm_cd": str})
lp["date"] = pd.to_datetime(lp["date"])
lp["dose_key"] = lp["adm_cd"].map(lambda c: XW.get(c, c))
agg = lp.groupby(["dose_key", "date"], as_index=False)["lp_day"].sum()
base = agg[(agg["date"] >= "2022-07-01") & (agg["date"] <= "2023-12-31")].groupby("dose_key")["lp_day"].mean()

norm = TwoSlopeNorm(vcenter=1.0, vmin=0.75, vmax=1.15)
cm = plt.get_cmap(FS.CMAP_DIV)

fig, axes = plt.subplots(2, 2, figsize=(6.5, 6.9))   # 인쇄 실크기
for axi, (pl, title, d0, d1) in zip(axes.flat, PHASES):
    win = agg[(agg["date"] >= d0) & (agg["date"] <= d1)].groupby("dose_key")["lp_day"].mean()
    rel = (win / base).reindex(base.index)
    cols = [cm(norm(rel[k])) if k in rel.index and pd.notna(rel.get(k)) else FS.NA_FILL for k in KEYS]
    FS.draw_polys(axi, PATS, cols, lw=.18)
    axi.set_title(f"({pl}) {title}", fontsize=8.6, loc="left")
    axi.text(.5, -.035, f"city-wide median {rel.median():.2f}× baseline",
             transform=axi.transAxes, ha="center", fontsize=7.5, color="#555555")
FS.add_scalebar(axes.flat[0], loc=(0.02, 0.05))

sm = ScalarMappable(norm=norm, cmap=cm); sm.set_array([])
cb = fig.colorbar(sm, ax=axes, orientation="horizontal", fraction=.035, pad=.05,
                  aspect=42, extend="both")
cb.set_label("Daytime de-facto population relative to post-lifting baseline "
             "(blue = emptied, red = fuller)", fontsize=9.3)
cb.outline.set_visible(False); cb.ax.tick_params(labelsize=8.5)

out = os.path.join(FS.FIG, "fig4_phase_maps.png")
plt.savefig(out)
for pl, title, d0, d1 in PHASES:
    win = agg[(agg["date"] >= d0) & (agg["date"] <= d1)].groupby("dose_key")["lp_day"].mean()
    rel = win / base
    print(f"({pl}) {title[:34]:36s} median {rel.median():.3f} · min {rel.min():.2f}")
print("→", out)
