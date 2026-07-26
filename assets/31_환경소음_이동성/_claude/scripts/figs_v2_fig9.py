# -*- coding: utf-8 -*-
# Fig 9 (구 Fig 11, robustness): (a) placebo 순열 null 분포(브레이크 축, dead space 제거)
#                                (b) 데실 dose-response + 95% CI(동클러스터).
# placebo 분포는 phase4와 동일 seed·로직으로 재현 → cache/placebo_betas.csv 캐시.
import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import figstyle_v2 as FS

FS.apply_style()
CACHEF = os.path.join(FS.CACHE, "placebo_betas_dong.csv")   # 동 단위 셔플 정본 (analysis_v9_fixes.py)
rob = pd.read_csv(os.path.join(FS.P, "phase4_robustness.csv")).set_index("check")["value"]
beta_hat = rob["headline_beta"]

if os.path.exists(CACHEF):
    betas = pd.read_csv(CACHEF)["beta"].values
else:
    rng = np.random.default_rng(20260622)
    pan = pd.read_csv(os.path.join(FS.P, "analysis_panel.csv"), dtype={"serial": str, "dose_key": str},
                      usecols=["serial", "dose_key", "date", "Leq_day", "lp_day_logrel", "temp_mean"])
    d = pan.dropna(subset=["Leq_day", "lp_day_logrel", "serial", "date", "dose_key"]).reset_index(drop=True)
    g1 = pd.factorize(d["serial"].values)[0]; n1 = g1.max() + 1; c1 = np.bincount(g1, minlength=n1)
    g2 = pd.factorize(d["date"].values)[0]; n2 = g2.max() + 1; c2 = np.bincount(g2, minlength=n2)
    def demean(v, iters=15):
        v = v.astype(float).copy()
        for _ in range(iters):
            v -= (np.bincount(g1, v, n1) / c1)[g1]
            v -= (np.bincount(g2, v, n2) / c2)[g2]
        return v
    y = d["Leq_day"].values.astype(float); x = d["lp_day_logrel"].values.astype(float)
    yt = demean(y); xt = demean(x)
    bh = np.cov(yt, xt)[0, 1] / np.var(xt)
    print(f"재현 β={bh:+.4f} (정본 {beta_hat:+.4f})")
    dates = d["date"].values; xser = pd.Series(x)
    NPERM = 300; betas = np.empty(NPERM)
    for i in range(NPERM):
        xp = xser.groupby(dates, sort=False).transform(lambda s: rng.permutation(s.values)).values
        xpt = demean(xp)
        betas[i] = np.cov(yt, xpt)[0, 1] / np.var(xpt)
        if (i + 1) % 50 == 0:
            print(f"  perm {i+1}/{NPERM}")
    pd.DataFrame({"beta": betas}).to_csv(CACHEF, index=False)

p_perm = (np.sum(np.abs(betas) >= abs(beta_hat)) + 1) / (len(betas) + 1)
print(f"placebo: mean {betas.mean():+.4f} SD {betas.std():.4f} p={p_perm:.4f} "
      f"(정본 {rob['placebo_perm_null_mean']:+.4f}/{rob['placebo_perm_null_sd']:.4f}/{rob['placebo_perm_p']:.4f})")

dec = pd.read_csv(os.path.join(FS.CACHE, "binned_deciles.csv"))

# 인쇄 실크기(6.5in) — (a) 브레이크축 상단 전폭, (b) 데실 하단 전폭
fig = plt.figure(figsize=(6.5, 6.6))
gs0 = fig.add_gridspec(2, 1, height_ratios=[1.0, 1.05], hspace=0.30)
gsA = gs0[0].subgridspec(1, 2, width_ratios=[3.0, 0.7], wspace=0.08)
axL = fig.add_subplot(gsA[0]); axR = fig.add_subplot(gsA[1], sharey=axL); axB = fig.add_subplot(gs0[1])

# ---------- (a) 브레이크 축: 좌=null 분포, 우=실제 β ----------
axL.hist(betas, bins=30, color="#B9CFE3", edgecolor=FS.MOB, lw=.6)
axL.axvline(0, color=FS.NEUTRAL, lw=.8, ls="--")
axL.set_xlim(-0.09, 0.09)
axL.set_xlabel("Placebo β (dose shuffled within date)")
axL.set_ylabel("Frequency")
axL.text(.05, .84, f"null: {betas.mean():+.3f} ± {betas.std():.3f}\ntwo-sided p = {p_perm:.3f}\n({len(betas)} dong-level shuffles)",
         transform=axL.transAxes, fontsize=8.2,
         bbox=dict(boxstyle="round,pad=0.32", fc="white", ec="#CCCCCC"))
axR.axvline(beta_hat, color=FS.NOISE, lw=2.6)
axR.set_xlim(0.55, 0.75); axR.set_xticks([0.65])
axR.annotate(f"actual β = {beta_hat:+.2f}", xy=(beta_hat, axL.get_ylim()[1] * 0.55),
             xytext=(-10, 18), textcoords="offset points", fontsize=8.6, color=FS.NOISE,
             fontweight="bold", rotation=90, va="center")
axR.spines["left"].set_visible(False)
axR.tick_params(axis="y", left=False, labelleft=False)
# 브레이크 마크
for a, x in ((axL, 1.0), (axR, 0.0)):
    a.plot([x, x], [-0.02, 0.02], transform=a.transAxes, color="k", lw=1,
           clip_on=False, marker=[(-0.4, -1), (0.4, 1)], ms=7, ls="")
FS.panel_label(axL, "a", dy=0.02)

# ---------- (b) 데실 dose-response + 95% CI ----------
axB.errorbar(dec["xt"], dec["yt"], yerr=1.96 * dec["se"], fmt="o", color=FS.MOB,
             ms=6.5, mfc="white", mew=1.6, capsize=3, lw=1.4,
             label="Decile mean ± 95% CI (dong-clustered)")
xs = np.linspace(dec["xt"].min(), dec["xt"].max(), 50)
axB.plot(xs, beta_hat * xs, color=FS.NOISE, lw=1.9, label=f"Linear fit (β = {beta_hat:+.2f})")
axB.axhline(0, color=FS.NEUTRAL, lw=.7, ls="--"); axB.axvline(0, color=FS.NEUTRAL, lw=.7, ls="--")
axB.set_xlabel("Mobility (two-way demeaned, log relative)")
axB.set_ylabel("Noise (two-way demeaned, dB)")
axB.legend(loc="upper left", fontsize=8.2)
FS.panel_label(axB, "b", dy=0.02)

plt.tight_layout()
out = os.path.join(FS.FIG, "fig9_robustness.png")
plt.savefig(out)
print("→", out)
