# -*- coding: utf-8 -*-
# Fig 4 기능별 시계열 (2x1, 논문1 Fig4 패턴): 토지이용별 (a) 주간 이동량 (b) 주야 gap 궤적.
# 상업 동이 제한기에 비었다 회복 / 주야 gap 압축 — 기능별 차등을 시간축으로.
import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import figstyle as FS

FS.apply_style()
ROOT = r"T:\00_BACKUP\01_한양대학교\00_연구실\00_프로젝트\★_논문화 프로젝트\31_환경소음_이동성"
P = os.path.join(ROOT, "data", "processed")
FIG = os.path.join(ROOT, "_claude", "figures")

pan = pd.read_csv(os.path.join(P, "analysis_panel.csv"), dtype={"serial": str, "dose_key": str},
                  usecols=["dose_key", "date", "dLeq_day", "lp_day_rel", "lp_day", "lp_night"])
pan["date"] = pd.to_datetime(pan["date"])
pan["week"] = pan["date"].dt.to_period("W").dt.start_time
# 주별 도시평균 dLeq_day (공통 드리프트/계절) — 그룹에서 빼면 drift-robust 상대 소음
city_wk = pan.groupby("week")["dLeq_day"].mean().rename("city_dLeq")
post = pan[(pan["date"] >= "2022-07-01") & (pan["date"] <= "2023-12-31")]
comm = post.groupby("dose_key").apply(lambda g: g["lp_day"].mean() / g["lp_night"].mean(), include_groups=False)
q1, q2 = comm.quantile([1/3, 2/3])
pan["lu"] = pan["dose_key"].map(lambda c: "commercial" if comm.get(c, np.nan) >= q2
                                else ("residential" if comm.get(c, np.nan) <= q1 else "mixed"))

anchors = {"2020-08-30": "21h cap", "2020-12-23": "5-person", "2021-07-12": "Lv.4",
           "2021-11-01": "With-COVID", "2022-04-18": "Full lift"}
LU = ["commercial", "mixed", "residential"]

def wk_by_lu(col):
    # 동별 주평균 -> 토지이용 주평균(동 동일가중)
    dw = pan.groupby(["lu", "dose_key", "week"])[col].mean().reset_index()
    return dw.groupby(["lu", "week"])[col].mean().reset_index()

mob = wk_by_lu("lp_day_rel")
# 토지이용별 주간 소음(그룹 주평균) - 도시 주평균 = drift-robust 상대소음
noi = wk_by_lu("dLeq_day").merge(city_wk, on="week")
noi["rel_noise"] = noi["dLeq_day"] - noi["city_dLeq"]
# 4주 이동평균(주별 노이즈 평활)
noi = noi.sort_values(["lu", "week"])
noi["rel_noise_s"] = noi.groupby("lu")["rel_noise"].transform(lambda s: s.rolling(4, min_periods=1, center=True).mean())

fig, ax = plt.subplots(2, 1, figsize=(10.5, 7.4), sharex=True, gridspec_kw={"hspace": 0.12})
for lu in LU:
    d = mob[mob.lu == lu]
    ax[0].plot(d["week"], d["lp_day_rel"], color=FS.LANDUSE_COLORS[lu], lw=1.8,
               label=FS.LANDUSE_LABEL[lu].split(" (")[0])
ax[0].axhline(1, color=FS.ACCENT["neutral"], lw=.7, ls="--")
ax[0].set_ylabel("Daytime mobility (rel. to baseline)")
ax[0].set_title("(a)", fontsize=11)
ax[0].legend(loc="lower right", ncol=3, fontsize=8.5)
ax[1].axhline(0, color=FS.ACCENT["neutral"], lw=.7, ls="--")
for lu in LU:
    d = noi[noi.lu == lu]
    ax[1].plot(d["week"], d["rel_noise_s"], color=FS.LANDUSE_COLORS[lu], lw=1.9,
               label=FS.LANDUSE_LABEL[lu].split(" (")[0])
ax[1].set_ylabel("Noise vs city mean (dB, 4-wk avg)")
ax[1].set_title("(b)", fontsize=11)
for a in ax:
    for d0, lab in anchors.items():
        a.axvline(pd.Timestamp(d0), color=FS.ACCENT["neutral"], lw=.6, alpha=.4)
ymax = ax[0].get_ylim()[1]
for d0, lab in anchors.items():
    ax[0].annotate(lab, (pd.Timestamp(d0), ymax), fontsize=7, rotation=90, va="top", ha="right", alpha=.8, color="#555")
ax[1].xaxis.set_major_locator(mdates.MonthLocator(interval=3))
ax[1].xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m"))
plt.setp(ax[1].xaxis.get_majorticklabels(), rotation=45, ha="right")
# suptitle 제거(캡션이 대신함)
plt.tight_layout()
out = os.path.join(FIG, "fig_landuse_trajectory.png")
plt.savefig(out)
# 수치
print("국면 평균 이동량(토지이용별):")
for lu in LU:
    rest = pan[(pan.lu == lu) & (pan.date >= "2021-07-12") & (pan.date <= "2021-09-30")]
    print(f"  {lu}: 4단계기 lp_day_rel 중앙 {rest['lp_day_rel'].median():.3f}")
print(f"-> {out}")
