# -*- coding: utf-8 -*-
# Fig 5 (구 Fig 6, landuse trajectory): (a) 토지이용별 주간 이동량 (b) 도시평균 대비 상대소음(4주 MA).
# v2: 강제한기 음영·이벤트 라벨 상단 정돈·고대비 토지이용 3색·범례 상단 가로.
import os
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import figstyle_v2 as FS

FS.apply_style()
wk = pd.read_csv(os.path.join(FS.CACHE, "weekly_lu.csv"), parse_dates=["week"])
LU = ["commercial", "mixed", "residential"]

fig, ax = plt.subplots(2, 1, figsize=(10.8, 7.2), sharex=True,
                       gridspec_kw={"hspace": 0.14})
for a in ax:
    FS.shade_restrictions(a)

for lu in LU:
    d = wk[wk.lu == lu]
    ax[0].plot(d["week"], d["mob"], color=FS.LANDUSE[lu], lw=1.6, label=FS.LANDUSE_LABEL[lu])
ax[0].axhline(1, color=FS.NEUTRAL, lw=.8, ls="--")
ax[0].set_ylabel("Daytime mobility\n(relative to post-lifting baseline)")
ax[0].legend(loc="lower right", ncol=3, fontsize=8.8, columnspacing=1.2)
FS.mark_events(ax[0], y_frac=0.995, fontsize=7.2)
FS.panel_label(ax[0], "a", dy=0.02)

for lu in LU:
    d = wk[wk.lu == lu]
    ax[1].plot(d["week"], d["rel_noise_s"], color=FS.LANDUSE[lu], lw=1.7, label=FS.LANDUSE_LABEL[lu])
ax[1].axhline(0, color=FS.NEUTRAL, lw=.8, ls="--")
ax[1].set_ylabel("Daytime noise vs city mean\n(dB, 4-week moving average)")
FS.mark_events(ax[1], labels=False)
FS.panel_label(ax[1], "b", dy=0.02)

ax[1].xaxis.set_major_locator(mdates.MonthLocator(interval=3))
ax[1].xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m"))
plt.setp(ax[1].xaxis.get_majorticklabels(), rotation=45, ha="right")
ax[1].text(0.012, 0.05, "shaded bands = strong-restriction periods (stringency ≥ 4)",
           transform=ax[1].transAxes, fontsize=7.6, color="#777777")

plt.tight_layout()
out = os.path.join(FS.FIG, "fig5_trajectory.png")
plt.savefig(out)
print("→", out)
