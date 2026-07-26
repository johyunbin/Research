# -*- coding: utf-8 -*-
# Fig 3 기능분할 dose-response forest (2x2): 결과(주/야/gap) × 토지이용 × 평일/주말 × 계절.
# 전부 센서+날짜 양방향 FE(drift-robust). 효과가 어느 기능에서 (안)나오는지 한눈에.
import os
import numpy as np
import pandas as pd
import statsmodels.api as sm
import matplotlib.pyplot as plt
import figstyle as FS

FS.apply_style()
ROOT = r"T:\00_BACKUP\01_한양대학교\00_연구실\00_프로젝트\★_논문화 프로젝트\31_환경소음_이동성"
P = os.path.join(ROOT, "data", "processed")
FIG = os.path.join(ROOT, "_claude", "figures")

pan = pd.read_csv(os.path.join(P, "analysis_panel.csv"), dtype={"serial": str, "dose_key": str},
                  usecols=["serial", "dose_key", "date", "Leq_day", "Leq_night",
                           "lp_day_logrel", "lp_night_logrel", "lp_day", "lp_night",
                           "is_weekend", "season"])
pan["date"] = pd.to_datetime(pan["date"])
pan["gap"] = pan["Leq_day"] - pan["Leq_night"]
post = pan[(pan["date"] >= "2022-07-01") & (pan["date"] <= "2023-12-31")]
comm = post.groupby("dose_key").apply(lambda g: g["lp_day"].mean() / g["lp_night"].mean(), include_groups=False)
q1, q2 = comm.quantile([1/3, 2/3])
pan["lu"] = pan["dose_key"].map(lambda c: "commercial" if comm.get(c, np.nan) >= q2
                                else ("residential" if comm.get(c, np.nan) <= q1 else "mixed"))

def twoway(d, ycol, xcol, g1="serial", g2="date", clust="dose_key", iters=18):
    d = d.dropna(subset=[ycol, xcol, g1, g2, clust])
    tmp = pd.DataFrame({"y": d[ycol].astype(float).values, "x": d[xcol].astype(float).values,
                        g1: d[g1].values, g2: d[g2].values})
    last = None
    for _ in range(iters):
        for c in ("y", "x"):
            tmp[c] -= tmp.groupby(g1)[c].transform("mean")
            tmp[c] -= tmp.groupby(g2)[c].transform("mean")
        v = float(np.var(tmp["x"].values))
        if last is not None and abs(v - last) < 1e-9 * max(1.0, last):
            break
        last = v
    r = sm.OLS(tmp["y"].values, sm.add_constant(tmp["x"].values, has_constant="add")).fit(
        cov_type="cluster", cov_kwds={"groups": d[clust].values})
    return r.params[1], r.bse[1], r.pvalues[1], len(d)

# 4개 기능축 (전부 주간 Leq_day ~ lp_day_logrel, 단 outcome 패널만 결과변수 변경)
panels = {
    "(a) By outcome": [("Daytime L_day", pan, "Leq_day", "lp_day_logrel"),
                       ("Nighttime L_night", pan, "Leq_night", "lp_night_logrel"),
                       ("Day-night gap", pan, "gap", "lp_day_logrel")],
    "(b) By land-use (daytime)": [("Commercial", pan[pan.lu == "commercial"], "Leq_day", "lp_day_logrel"),
                                  ("Mixed", pan[pan.lu == "mixed"], "Leq_day", "lp_day_logrel"),
                                  ("Residential", pan[pan.lu == "residential"], "Leq_day", "lp_day_logrel")],
    "(c) By day type (daytime)": [("Weekday", pan[pan.is_weekend == 0], "Leq_day", "lp_day_logrel"),
                                  ("Weekend/holiday", pan[pan.is_weekend == 1], "Leq_day", "lp_day_logrel")],
    "(d) By season (daytime)": [(s, pan[pan.season == s], "Leq_day", "lp_day_logrel")
                                for s in ["DJF", "MAM", "JJA", "SON"]],
}
results = {}
print("기능분할 dose-response (양방향 FE):")
for pk, specs in panels.items():
    rows = []
    for lab, d, y, x in specs:
        b, se, p, n = twoway(d, y, x)
        rows.append((lab, b, se, p, n))
        print(f"  {pk[:18]:18s} {lab:16s} β={b:+.3f} (95%CI {b-1.96*se:+.2f}~{b+1.96*se:+.2f}) p={p:.2f} n={n:,}")
    results[pk] = rows

fig, axes = plt.subplots(2, 2, figsize=(9.2, 9.0))   # 정사각형
COL = FS.ACCENT["mobility"]
for axi, (pk, rows) in zip(axes.flat, results.items()):
    yp = np.arange(len(rows))[::-1]
    for yy, (lab, b, se, p, n) in zip(yp, rows):
        sig = p < .05
        axi.errorbar(b, yy, xerr=1.96 * se, fmt="o", color=COL if sig else "#B9C7BF", ms=7,
                     capsize=4, lw=1.7, mfc="white", mew=1.7)
        axi.annotate(f"{b:+.2f}{'*' if sig else ''}", (b, yy), textcoords="offset points",
                     xytext=(0, 10), ha="center", fontsize=8)
    axi.axvline(0, color=FS.ACCENT["neutral"], lw=.8, ls="--")
    axi.set_yticks(yp)
    axi.set_yticklabels([r[0].replace("L_day", "L$_{day}$").replace("L_night", "L$_{night}$") for r in rows])
    axi.set_ylim(-0.6, len(rows) - 0.4)
    axi.set_xlabel("β (dB per log-unit mobility)")
    axi.set_title(pk.split(")")[0] + ")", fontsize=11)
# suptitle 제거(캡션이 대신함)
plt.tight_layout()
out = os.path.join(FIG, "fig_segmented_forest.png")
plt.savefig(out)
flat = [[pk] + list(r) for pk, rows in results.items() for r in rows]
pd.DataFrame(flat, columns=["panel", "group", "beta", "se", "pval", "n"]).to_csv(
    os.path.join(P, "phase3_segmented_results.csv"), index=False, encoding="utf-8-sig")
print(f"-> {out}")
