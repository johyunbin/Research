# -*- coding: utf-8 -*-
# Phase 4 강건성: ①placebo 순열검정(날짜 내 dose 셔플) ②placebo 결과변수(기상~이동량) ③비선형성(2차·분위) ④분할추정 FDR.
# 식별 정당화: '같은 날 동 사이 변이' 신호가 기계적 아티팩트가 아님을 무작위화추론으로 입증.
import os
import numpy as np
import pandas as pd
import statsmodels.api as sm
from scipy import stats
import matplotlib.pyplot as plt
import figstyle as FS

FS.apply_style()
ROOT = r"T:\00_BACKUP\01_한양대학교\00_연구실\00_프로젝트\★_논문화 프로젝트\31_환경소음_이동성"
P = os.path.join(ROOT, "data", "processed")
FIG = os.path.join(ROOT, "_claude", "figures")
rng = np.random.default_rng(20260622)

pan = pd.read_csv(os.path.join(P, "analysis_panel.csv"), dtype={"serial": str, "dose_key": str},
                  usecols=["serial", "dose_key", "date", "Leq_day", "lp_day_logrel", "temp_mean"])
d = pan.dropna(subset=["Leq_day", "lp_day_logrel", "serial", "date", "dose_key"]).reset_index(drop=True)
print(f"표본: {len(d):,} sensor-days | 센서 {d.serial.nunique()} | 동 {d.dose_key.nunique()}")

# --- 빠른 양방향 demean (factorize + bincount) ---
g1 = pd.factorize(d["serial"].values)[0]; n1 = g1.max() + 1; c1 = np.bincount(g1, minlength=n1)
g2 = pd.factorize(d["date"].values)[0]; n2 = g2.max() + 1; c2 = np.bincount(g2, minlength=n2)
gd = pd.factorize(d["dose_key"].values)[0]   # 클러스터(동)

def demean(v, iters=15):
    v = v.astype(float).copy()
    for _ in range(iters):
        v -= (np.bincount(g1, v, n1) / c1)[g1]
        v -= (np.bincount(g2, v, n2) / c2)[g2]
    return v

y = d["Leq_day"].values.astype(float)
x = d["lp_day_logrel"].values.astype(float)
yt = demean(y)                      # ỹ (한 번만)
xt = demean(x)
beta_hat = np.cov(yt, xt)[0, 1] / np.var(xt)
print(f"\n[기준 재현] 양방향FE β(주간) = {beta_hat:+.4f}  (정본 +0.648 대조)")

# ===== ① Placebo 순열검정: 날짜 내에서 dose를 무작위 재배치 → β≈0 이어야 =====
NPERM = 300
dates = d["date"].values
xser = pd.Series(x)
betas = np.empty(NPERM)
for i in range(NPERM):
    xp = xser.groupby(dates, sort=False).transform(lambda s: rng.permutation(s.values)).values
    xpt = demean(xp)
    betas[i] = np.cov(yt, xpt)[0, 1] / np.var(xpt)
p_perm = (np.sum(np.abs(betas) >= abs(beta_hat)) + 1) / (NPERM + 1)
print(f"① placebo 순열({NPERM}회): null β 평균 {betas.mean():+.4f}, SD {betas.std():.4f}, |null| 최대 {np.abs(betas).max():.3f}")
print(f"   실제 β={beta_hat:+.3f} → 양측 순열 p = {p_perm:.4f}  (작을수록 신호가 진짜)")

# ===== ② Placebo 결과변수: 기상(기온)~이동량. 이동량이 기상을 만들 수 없으므로 β≈0 이어야 =====
tt = demean(d["temp_mean"].values.astype(float))
Xc = sm.add_constant(xt, has_constant="add")
r_temp = sm.OLS(tt, Xc).fit(cov_type="cluster", cov_kwds={"groups": gd})
print(f"\n② placebo 결과변수(기온~이동량): β={r_temp.params[1]:+.4f} (p={r_temp.pvalues[1]:.3f}) → 0 근처여야 정상")

# ===== ③ 비선형성: 2차항 + 분위 구간 =====
x2 = x * x
x2t = demean(x2)
Xq = sm.add_constant(np.column_stack([xt, x2t]), has_constant="add")
r_q = sm.OLS(yt, Xq).fit(cov_type="cluster", cov_kwds={"groups": gd})
print(f"\n③ 비선형성 2차: linear β1={r_q.params[1]:+.3f}(p={r_q.pvalues[1]:.3f}), quad β2={r_q.params[2]:+.3f}(p={r_q.pvalues[2]:.3f})")
print("   → β2 비유의면 선형 근사 타당")
# 분위 binned: x̃을 데실로 나눠 ỹ 평균(선형성 시각 확인)
dec = pd.qcut(xt, 10, labels=False, duplicates="drop")
binned = pd.DataFrame({"dec": dec, "xt": xt, "yt": yt}).groupby("dec").mean()

# ===== ④ 분할추정 FDR 보정 (Table 6의 12개 검정) =====
seg = pd.read_csv(os.path.join(P, "phase3_segmented_results.csv"))
pv = seg["pval"].values
order = np.argsort(pv); m = len(pv)
fdr = np.empty(m)
prev = 1.0
for rank, idx in enumerate(order[::-1]):          # BH 절차
    k = m - rank
    prev = min(prev, pv[idx] * m / k)
    fdr[idx] = prev
seg["p_BH_FDR"] = fdr
seg["sig_FDR_05"] = fdr < 0.05
print(f"\n④ 분할추정 {m}개 검정 BH-FDR: 원래 p<.05 {int((pv<.05).sum())}개 → FDR<.05 {int(seg['sig_FDR_05'].sum())}개 생존")
print(seg[["panel", "group", "beta", "pval", "p_BH_FDR", "sig_FDR_05"]].to_string(index=False))

# 저장
pd.DataFrame({
    "check": ["headline_beta", "placebo_perm_null_mean", "placebo_perm_null_sd", "placebo_perm_p",
              "placebo_outcome_temp_beta", "placebo_outcome_temp_p",
              "nonlin_linear_b1", "nonlin_quad_b2", "nonlin_quad_p"],
    "value": [beta_hat, betas.mean(), betas.std(), p_perm,
              r_temp.params[1], r_temp.pvalues[1], r_q.params[1], r_q.params[2], r_q.pvalues[2]],
}).to_csv(os.path.join(P, "phase4_robustness.csv"), index=False, encoding="utf-8-sig")
seg.to_csv(os.path.join(P, "phase3_segmented_fdr.csv"), index=False, encoding="utf-8-sig")

# ===== 그림: (a) placebo null 분포 + 실제 β  (b) binned dose-response(선형성) =====
fig, ax = plt.subplots(1, 2, figsize=(10.5, 5.0))
ax[0].hist(betas, bins=30, color="#C9D6C3", edgecolor="#7a8a72", alpha=.9)
ax[0].axvline(beta_hat, color=FS.ACCENT["noise"], lw=2.2, label=f"Actual β={beta_hat:+.2f}")
ax[0].axvline(0, color=FS.ACCENT["neutral"], lw=.8, ls="--")
ax[0].set_xlabel("Placebo β (dose shuffled within date)"); ax[0].set_ylabel("Frequency")
ax[0].set_title("(a)"); ax[0].legend(loc="upper left", fontsize=8.5)
ax[0].text(.04, .72, f"permutation p = {p_perm:.3f}\n{NPERM} shuffles", transform=ax[0].transAxes,
           fontsize=8.5, bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="#bbb"))
ax[1].plot(binned["xt"], binned["yt"], "o", color=FS.ACCENT["mobility"], ms=7, mfc="white", mew=1.6)
xs = np.linspace(binned["xt"].min(), binned["xt"].max(), 50)
ax[1].plot(xs, beta_hat * xs, color=FS.ACCENT["noise"], lw=1.8, label=f"Linear fit (β={beta_hat:+.2f})")
ax[1].set_xlabel("Mobility (two-way demeaned log rel.)"); ax[1].set_ylabel("Noise (two-way demeaned, dB)")
ax[1].set_title("(b)"); ax[1].legend(loc="upper left", fontsize=8.5)
for a in ax:
    a.set_box_aspect(1)
plt.tight_layout()
out = os.path.join(FIG, "fig_robustness.png")
plt.savefig(out)
print(f"\n-> {out}")
print("-> phase4_robustness.csv, phase3_segmented_fdr.csv")
