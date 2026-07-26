# -*- coding: utf-8 -*-
# Phase 2-G 주간/야간 비교 + 주간-야간 gap(L_day - L_night) dose-response.
# gap = 같은 센서·같은 날 차분 -> 센서 오프셋 AND 공통 드리프트 상쇄(주·야 함께 드리프트), baseline 불요.
# 가설: 주간 이동량↓ -> 주간소음이 야간수준으로 -> gap 축소 (lp_day_logrel에 양(+) 계수).
import os
import numpy as np
import pandas as pd
import statsmodels.api as sm
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import figstyle as FS

FS.apply_style()
ROOT = r"T:\00_BACKUP\01_한양대학교\00_연구실\00_프로젝트\★_논문화 프로젝트\31_환경소음_이동성"
P = os.path.join(ROOT, "data", "processed")
FIG = os.path.join(ROOT, "_claude", "figures")

panel = pd.read_csv(os.path.join(P, "analysis_panel.csv"), dtype={"serial": str, "dose_key": str},
                    usecols=["serial", "dose_key", "date", "Leq_day", "Leq_night",
                             "lp_day_logrel", "lp_night_logrel", "stringency",
                             "temp_mean", "precip_mm", "wind_max_ms", "rain_flag", "is_weekend", "holiday"])
panel["date"] = pd.to_datetime(panel["date"])
panel["gap"] = panel["Leq_day"] - panel["Leq_night"]
panel["temp_sq"] = panel["temp_mean"] ** 2
WX = ["temp_mean", "temp_sq", "precip_mm", "wind_max_ms", "rain_flag", "is_weekend", "holiday"]

def within_fe(d, ycol, xcols, g="serial"):
    d = d.dropna(subset=[ycol] + xcols + [g]).copy()
    gg = d.groupby(g)
    Y = d[ycol] - gg[ycol].transform("mean")
    X = pd.DataFrame({c: d[c] - gg[c].transform("mean") for c in xcols})
    X = sm.add_constant(X, has_constant="add")
    r = sm.OLS(Y.values, X.values).fit(cov_type="cluster", cov_kwds={"groups": d[g].values})
    i = 1
    return r.params[i], r.bse[i], r.pvalues[i], len(d)

def twoway_fe(d, ycol, xcol, g1="serial", g2="date", clust="dose_key", iters=20):
    d = d.dropna(subset=[ycol, xcol, g1, g2, clust]).copy()
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
    X = sm.add_constant(tmp["x"].values, has_constant="add")
    r = sm.OLS(tmp["y"].values, X).fit(cov_type="cluster", cov_kwds={"groups": d[clust].values})
    return r.params[1], r.bse[1], r.pvalues[1], len(d)

specs = [("Daytime L_day", "Leq_day", "lp_day_logrel"),
         ("Nighttime L_night", "Leq_night", "lp_night_logrel"),
         ("Day−night gap", "gap", "lp_day_logrel")]
print("=== 주간/야간/gap dose-response ===")
rows = []
for lab, y, x in specs:
    bw, sew, pw, nw = within_fe(panel, y, [x] + WX)
    bt, set_, pt, nt = twoway_fe(panel, y, x)
    print(f"{lab:18s} | 센서FE β={bw:+.3f}(p={pw:.1e}) | 양방향FE β={bt:+.3f} (95%CI {bt-1.96*set_:+.2f}~{bt+1.96*set_:+.2f}, p={pt:.1e}, n={nt:,})")
    rows.append([lab, bw, sew, pw, bt, set_, pt, nt])
res = pd.DataFrame(rows, columns=["outcome", "b_within", "se_within", "p_within", "b_2way", "se_2way", "p_2way", "n"])
res.to_csv(os.path.join(P, "phase2g_daynight_results.csv"), index=False, encoding="utf-8-sig")

# === 그림: (a) 양방향FE forest(day/night/gap) (b) within FE vs 양방향FE 대비(야간 효과 소멸=시간교란) ===
fig, ax = plt.subplots(1, 2, figsize=(10.5, 5.4))   # 각 패널 정사각형(set_box_aspect)
# (a) forest
cols = [FS.ACCENT["noise"], "#5b7fa6", FS.ACCENT["mobility"]]
yp = np.arange(len(res))[::-1]
for i, (_, r) in enumerate(res.iterrows()):
    yy = yp[i]
    ax[0].errorbar(r["b_2way"], yy, xerr=1.96 * r["se_2way"], fmt="o", color=cols[i], ms=8,
                   capsize=4, lw=1.8, mfc="white", mew=1.8)
    sig = "*" if r["p_2way"] < .05 else " (ns)"
    ax[0].annotate(f"β={r['b_2way']:+.2f}{sig}", (r["b_2way"], yy), textcoords="offset points",
                   xytext=(0, 11), ha="center", fontsize=8.5)
ax[0].axvline(0, color=FS.ACCENT["neutral"], lw=.8, ls="--")
ax[0].set_yticks(yp)
ax[0].set_yticklabels([o.replace("L_day", "L$_{day}$").replace("L_night", "L$_{night}$") for o in res["outcome"]])
ax[0].set_xlabel("Dose-response β (dB per log-unit mobility)")
ax[0].set_title("(a)")
ax[0].set_ylim(-0.6, len(res) - 0.4)
# (b) within FE vs 양방향 FE (day, night) — 야간 겉보기 효과가 엄밀식별에서 소멸
dn = res[res["outcome"].isin(["Daytime L_day", "Nighttime L_night"])].reset_index(drop=True)
x = np.arange(len(dn)); w = 0.36
b1 = ax[1].bar(x - w/2, dn["b_within"], w, yerr=1.96 * dn["se_within"], capsize=4,
               color="#C9D6C3", edgecolor="#7a8a72", label="Sensor FE (+ weather/calendar)")
b2 = ax[1].bar(x + w/2, dn["b_2way"], w, yerr=1.96 * dn["se_2way"], capsize=4,
               color=FS.ACCENT["mobility"], edgecolor="#3d6b54", label="Two-way FE (sensor + date)")
# 숫자를 막대가 아닌 '오차막대 끝' 위에 찍어 오차막대와 겹치지 않게
tops, bots = [], []
for xi, (_, r) in zip(x, dn.iterrows()):
    tw = r["b_within"] + 1.96 * r["se_within"]; t2 = r["b_2way"] + 1.96 * r["se_2way"]
    tops += [tw, t2]; bots += [r["b_2way"] - 1.96 * r["se_2way"], r["b_within"] - 1.96 * r["se_within"]]
    ax[1].annotate(f"{r['b_within']:+.2f}{'*' if r['p_within']<.05 else ''}", (xi - w/2, tw),
                   textcoords="offset points", xytext=(0, 6), ha="center", fontsize=8)
    ax[1].annotate(f"{r['b_2way']:+.2f}{'*' if r['p_2way']<.05 else ' ns'}", (xi + w/2, t2),
                   textcoords="offset points", xytext=(0, 6), ha="center", fontsize=8)
ax[1].set_ylim(min(-0.4, min(bots) - 0.3), max(tops) * 1.16)   # 숫자 들어갈 위 여유
ax[1].axhline(0, color=FS.ACCENT["neutral"], lw=.8, ls="--")
ax[1].set_xticks(x); ax[1].set_xticklabels(["Daytime\n(L$_{day}$)", "Nighttime\n(L$_{night}$)"])
ax[1].set_ylabel("Dose-response β (dB / log-unit)")
ax[1].set_title("(b)")
# 야간 within-FE 막대(우측 최고)를 가리지 않도록 레전드를 좌상단 바깥 가까이로
ax[1].legend(loc="upper left", fontsize=7.5, framealpha=0.9, borderpad=0.4)
for a in ax:
    a.set_box_aspect(1)   # 각 패널 정사각형
plt.tight_layout()
out = os.path.join(FIG, "fig_daynight.png")
plt.savefig(out)
print(f"-> {out}")
print(f"-> phase2g_daynight_results.csv")
