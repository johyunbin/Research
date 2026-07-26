# -*- coding: utf-8 -*-
# Phase 2-A 헤드라인: 양방향 고정효과(센서 + 날짜) within-회귀.
# 날짜 FE가 기상·요일·공휴일·도시추세·센서드리프트(날짜공통) 전부 흡수 ->
# 동 단위 이동량의 '같은날짜 내 동간 변이'만으로 순수 dose-response 식별 (가장 엄밀).
# 식별: 같은 날, 자기 baseline 대비 더 비워진 동의 센서가 더 조용해졌나?
import os
import numpy as np
import pandas as pd
import statsmodels.api as sm

ROOT = r"T:\00_BACKUP\01_한양대학교\00_연구실\00_프로젝트\★_논문화 프로젝트\31_환경소음_이동성"
P = os.path.join(ROOT, "data", "processed")
panel = pd.read_csv(os.path.join(P, "analysis_panel.csv"),
                    dtype={"serial": str, "dose_key": str},
                    usecols=["serial", "dose_key", "date", "Leq_day", "Leq_night", "Leq24",
                             "lp_day_logrel", "lp_night_logrel", "lp_mean_logrel"])

def twoway_fe(d, ycol, xcol, g1="serial", g2="date", clust="dose_key", iters=20, tol=1e-9):
    d = d.dropna(subset=[ycol, xcol, g1, g2, clust]).copy()
    tmp = pd.DataFrame({"y": d[ycol].astype(float).values, "x": d[xcol].astype(float).values,
                        g1: d[g1].values, g2: d[g2].values})
    last = None
    for it in range(iters):
        for col in ("y", "x"):
            tmp[col] -= tmp.groupby(g1)[col].transform("mean")
            tmp[col] -= tmp.groupby(g2)[col].transform("mean")
        # 수렴 점검(x의 잔차 분산 변화)
        v = float(np.var(tmp["x"].values))
        if last is not None and abs(v - last) < tol * max(1.0, last):
            break
        last = v
    X = sm.add_constant(tmp["x"].values, has_constant="add")
    res = sm.OLS(tmp["y"].values, X).fit(cov_type="cluster", cov_kwds={"groups": d[clust].values})
    b, se, p = res.params[1], res.bse[1], res.pvalues[1]
    return b, se, p, len(d), d[g1].nunique(), d[g2].nunique(), d[clust].nunique(), it + 1

rows = []
specs = [("Leq_day", "lp_day_logrel", "주간 소음 ~ 동 주간이동량"),
         ("Leq_night", "lp_night_logrel", "야간 소음 ~ 동 야간이동량"),
         ("Leq24", "lp_mean_logrel", "전일 소음 ~ 동 전일이동량")]
print("양방향 FE(센서+날짜) dose-response — 동간 변이만으로 식별\n")
for y, x, lab in specs:
    b, se, p, n, ns, nd, nc, ni = twoway_fe(panel, y, x)
    star = "***" if p < .001 else "**" if p < .01 else "*" if p < .05 else ""
    d30 = b * -0.357   # 이동량 30% 감소 효과
    print(f"=== {lab} ===")
    print(f"  β={b:+.3f} dB/log-unit  se={se:.3f}  p={p:.2e} {star}  (n={n:,} 센서={ns} 날짜={nd} 동={nc}, demean수렴 {ni}회)")
    print(f"  >> 이동량 30%↓ -> {d30:+.2f} dB | 50%↓ -> {b*np.log(0.5):+.2f} dB\n")
    rows.append([lab, y, x, b, se, p, n, ns, nd, nc])

pd.DataFrame(rows, columns=["label", "outcome", "dose", "beta", "se", "pval", "n", "sensors", "dates", "dongs"]).to_csv(
    os.path.join(P, "phase2a_twoway_results.csv"), index=False, encoding="utf-8-sig")
print("-> phase2a_twoway_results.csv 저장")
