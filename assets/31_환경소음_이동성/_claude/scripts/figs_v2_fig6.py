# -*- coding: utf-8 -*-
# Fig 6 (구 Fig 7, DiD event study): 고영향−저영향 동 주별 차이 (a) 소음 (b) 이동량.
# v2: **95% CI 밴드(동클러스터 SE)** 추가·강제한기 음영·패널라벨.
import os
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import figstyle_v2 as FS

FS.apply_style()
d = pd.read_csv(os.path.join(FS.CACHE, "did_weekly_ci.csv"), parse_dates=["week"])

fig, ax = plt.subplots(2, 1, figsize=(6.5, 5.8), sharex=True,
                       gridspec_kw={"hspace": 0.14})   # 인쇄 실크기
for a in ax:
    FS.shade_restrictions(a)
    a.axhline(0, color=FS.NEUTRAL, lw=.8, ls="--")

ax[0].fill_between(d["week"], d["dLeq_diff"] - 1.96 * d["dLeq_diff_se"],
                   d["dLeq_diff"] + 1.96 * d["dLeq_diff_se"],
                   color=FS.NOISE, alpha=.16, lw=0, label="95% CI (dong-clustered)")
ax[0].plot(d["week"], d["dLeq_diff"], color=FS.NOISE, lw=1.6)
ax[0].set_ylabel("ΔL$_{day}$ difference,\nhigh − low impact (dB)")
ax[0].legend(loc="upper right", fontsize=8.2)
FS.mark_events(ax[0], y_frac=0.995, fontsize=7.2)
FS.panel_label(ax[0], "a", dy=0.02)

ax[1].fill_between(d["week"], d["mob_diff"] - 1.96 * d["mob_diff_se"],
                   d["mob_diff"] + 1.96 * d["mob_diff_se"],
                   color=FS.MOB, alpha=.16, lw=0)
ax[1].plot(d["week"], d["mob_diff"], color=FS.MOB, lw=1.6)
ax[1].set_ylabel("Daytime mobility difference,\nhigh − low impact (relative)")
FS.mark_events(ax[1], labels=False)
FS.panel_label(ax[1], "b", dy=0.02)

ax[1].xaxis.set_major_locator(mdates.MonthLocator(interval=3))
ax[1].xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m"))
plt.setp(ax[1].xaxis.get_majorticklabels(), rotation=45, ha="right")

r = d["mob_diff"].corr(d["dLeq_diff"]); rs = d["mob_diff"].corr(d["dLeq_diff"], method="spearman")
ax[0].text(0.988, 0.06, f"weekly correlation of the two differences:\nPearson r = {r:+.2f} · Spearman ρ = {rs:+.2f}",
           transform=ax[0].transAxes, fontsize=7.8, ha="right", color="#555555",
           bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="#CCCCCC", alpha=.9))

plt.tight_layout()
out = os.path.join(FS.FIG, "fig6_did_eventstudy.png")
plt.savefig(out)
rest = d[d["week"] <= "2022-04-18"]; norm = d[d["week"] > "2022-04-18"]
print(f"제한기 평균 차이 {rest['dLeq_diff'].mean():+.3f} dB · 정상기 {norm['dLeq_diff'].mean():+.3f} dB")
print(f"상관 Pearson {r:+.3f} Spearman {rs:+.3f} → {out}")
