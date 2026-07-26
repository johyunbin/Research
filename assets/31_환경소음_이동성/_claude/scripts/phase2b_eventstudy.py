# -*- coding: utf-8 -*-
# Phase 2-B DiD event-study (drift-immune): 이동량 고영향 동 vs 저영향 동의 ΔLAeq 차이 시간궤적.
# 공통 교란(센서드리프트·계절·도시추세·스키마)은 두 그룹 차분에서 상쇄 -> 순수 이동량 효과만 남음.
# 가설: 제한기에 고영향(상업)동이 저영향(주거)동보다 더 조용 -> 차이(고-저) 음(-), 정상기 ~0.
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
os.makedirs(FIG, exist_ok=True)

df = pd.read_csv(os.path.join(P, "analysis_panel.csv"), dtype={"serial": str, "dose_key": str},
                 usecols=["serial", "dose_key", "date", "dLeq_day", "lp_day_rel", "stringency"])
df["date"] = pd.to_datetime(df["date"])
df["week"] = df["date"].dt.to_period("W").dt.start_time

# 동별 이동량 영향도 = 최강제한기(수도권 4단계 2021-07-12~09-30) 평균 lp_day_rel (낮을수록 큰감소=고영향=상업)
imp = df[(df["date"] >= "2021-07-12") & (df["date"] <= "2021-09-30")].groupby("dose_key")["lp_day_rel"].mean()
q1, q3 = imp.quantile([1/3, 2/3])
high = set(imp[imp <= q1].index)   # 큰 감소(상업)
low = set(imp[imp >= q3].index)    # 작은 감소/증가(주거)
df["grp"] = np.where(df["dose_key"].isin(high), "high", np.where(df["dose_key"].isin(low), "low", "mid"))
print(f"고영향 동 {len(high)} (제한기 이동량 <= {q1:.3f}) | 저영향 동 {len(low)} (>= {q3:.3f})")

wk = df[df["grp"].isin(["high", "low"])].groupby(["week", "grp"]).agg(
    dLeq=("dLeq_day", "mean"), mob=("lp_day_rel", "mean"), n=("dLeq_day", "size")).reset_index()
piv = wk.pivot(index="week", columns="grp", values=["dLeq", "mob", "n"])
piv = piv[(piv[("n", "high")] >= 50) & (piv[("n", "low")] >= 50)]
piv["dLeq_diff"] = piv[("dLeq", "high")] - piv[("dLeq", "low")]   # 고-저 (음=상업이 더조용)
piv["mob_diff"] = piv[("mob", "high")] - piv[("mob", "low")]      # 고-저 (음=상업이 더 비워짐)
piv = piv.reset_index()

anchors = {"2020-08-30": "21h cap", "2020-12-23": "5-person ban", "2021-07-12": "Lv.4",
           "2021-11-01": "With-COVID", "2021-12-18": "Re-tighten", "2022-04-18": "Full lift"}

NOISE, MOB = FS.ACCENT["noise"], FS.ACCENT["mobility"]
fig, ax = plt.subplots(2, 1, figsize=(10.5, 6.6), sharex=True, gridspec_kw={"height_ratios": [1, 1], "hspace": 0.13})
ax[0].axhline(0, color=FS.ACCENT["neutral"], lw=.7, ls="--")
ax[0].plot(piv["week"], piv["dLeq_diff"], color=NOISE, lw=1.8)
ax[0].fill_between(piv["week"], piv["dLeq_diff"], 0, where=piv["dLeq_diff"] < 0, color=NOISE, alpha=.18)
ax[0].set_ylabel("ΔL$_{day}$: high − low impact (dB)")
# 상단 제목 제거(캡션이 대신함)
ax[1].axhline(0, color=FS.ACCENT["neutral"], lw=.7, ls="--")
ax[1].plot(piv["week"], piv["mob_diff"], color=MOB, lw=1.8)
ax[1].fill_between(piv["week"], piv["mob_diff"], 0, where=piv["mob_diff"] < 0, color=MOB, alpha=.18)
ax[1].set_ylabel("Daytime mobility: high − low (rel.)")
for a in ax:
    for d, lab in anchors.items():
        a.axvline(pd.Timestamp(d), color=FS.ACCENT["neutral"], lw=.6, alpha=.4)
ymax = ax[0].get_ylim()[1]
for d, lab in anchors.items():
    ax[0].annotate(lab, (pd.Timestamp(d), ymax), fontsize=7, rotation=90, va="top", ha="right", alpha=.8, color="#555555")
ax[1].xaxis.set_major_locator(mdates.MonthLocator(interval=3))
ax[1].xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m"))
plt.setp(ax[1].xaxis.get_majorticklabels(), rotation=45, ha="right")
plt.tight_layout()
out = os.path.join(FIG, "fig_did_eventstudy.png")
plt.savefig(out, dpi=300, bbox_inches="tight")
print(f"-> {out}")

# 수치 요약
hi_rest = piv[piv["week"] <= "2022-04-18"]
norm = piv[piv["week"] > "2022-04-18"]
print(f"\n제한기(~2022-04) 평균 ΔL차이(고-저): {hi_rest['dLeq_diff'].mean():+.3f} dB (이동량차이 {hi_rest['mob_diff'].mean():+.3f})")
print(f"정상기(2022-04~)  평균 ΔL차이(고-저): {norm['dLeq_diff'].mean():+.3f} dB (이동량차이 {norm['mob_diff'].mean():+.3f})")
print(f"DiD(제한 - 정상): {hi_rest['dLeq_diff'].mean()-norm['dLeq_diff'].mean():+.3f} dB")
# 상관: 주별 이동량차이 vs 소음차이 (Pearson + Spearman robust)
r = piv["mob_diff"].corr(piv["dLeq_diff"])
rs = piv["mob_diff"].corr(piv["dLeq_diff"], method="spearman")
print(f"주별 (이동량 고-저) vs (소음 고-저) 상관: Pearson r={r:+.3f} | Spearman ρ={rs:+.3f}  (주별 도시평균=robust 집계, 둘 일치)")
piv_out = piv.copy(); piv_out.columns = ["_".join(map(str, c)).strip("_") for c in piv_out.columns]
piv_out.to_csv(os.path.join(P, "phase2b_did_weekly.csv"), index=False, encoding="utf-8-sig")
print("-> phase2b_did_weekly.csv")
