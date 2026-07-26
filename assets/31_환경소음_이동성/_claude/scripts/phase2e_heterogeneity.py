# -*- coding: utf-8 -*-
# Phase 2-E (옵션2): dose-response 이질성 — 효과는 상업/활동 동에 집중되는가? (RQ2 + 효과 선명화)
# 가설: 생활인구=소음활동 proxy의 타당성이 토지이용에 따라 다름. 상업동(주간/야간 인구비 高)에서 β 더 큼.
# 양방향 FE(센서+날짜) dose-response를 상업도 그룹별 + 상호작용으로 추정.
import os
import numpy as np
import pandas as pd
import statsmodels.api as sm

ROOT = r"T:\00_BACKUP\01_한양대학교\00_연구실\00_프로젝트\★_논문화 프로젝트\31_환경소음_이동성"
P = os.path.join(ROOT, "data", "처리".replace("처리", "processed"))
panel = pd.read_csv(os.path.join(P, "analysis_panel.csv"), dtype={"serial": str, "dose_key": str},
                    usecols=["serial", "dose_key", "date", "Leq_day", "lp_day_logrel", "lp_day", "lp_night", "year"])
panel["date"] = pd.to_datetime(panel["date"])

# 동 상업도 = post-lift 주간/야간 생활인구 비 (高=상업·업무, 低=주거)
post = panel[(panel["date"] >= "2022-07-01") & (panel["date"] <= "2023-12-31")]
comm = post.groupby("dose_key").apply(
    lambda g: g["lp_day"].mean() / g["lp_night"].mean(), include_groups=False).rename("comm_idx")
panel = panel.merge(comm, on="dose_key", how="left")
q1, q2 = comm.quantile([1/3, 2/3])
print(f"상업도(주간/야간 인구비) 33%={q1:.3f} 67%={q2:.3f} | 범위 {comm.min():.2f}~{comm.max():.2f}")

def twoway_fe(d, ycol, xcols, g1="serial", g2="date", clust="dose_key", iters=20, tol=1e-9):
    d = d.dropna(subset=[ycol] + xcols + [g1, g2, clust]).copy()
    tmp = pd.DataFrame({g1: d[g1].values, g2: d[g2].values})
    cols = [ycol] + xcols
    for c in cols:
        tmp[c] = d[c].astype(float).values
    last = None
    for it in range(iters):
        for c in cols:
            tmp[c] -= tmp.groupby(g1)[c].transform("mean")
            tmp[c] -= tmp.groupby(g2)[c].transform("mean")
        v = float(np.var(tmp[xcols[0]].values))
        if last is not None and abs(v - last) < tol * max(1.0, last):
            break
        last = v
    X = sm.add_constant(tmp[xcols].values, has_constant="add")
    res = sm.OLS(tmp[ycol].values, X).fit(cov_type="cluster", cov_kwds={"groups": d[clust].values})
    return res, len(d), d[g1].nunique(), d[clust].nunique()

print("\n=== 상업도 그룹별 dose-response (Leq_day ~ lp_day_logrel, 센서+날짜 FE) ===")
groups = [("주거 동(하위1/3)", panel[panel["comm_idx"] <= q1]),
          ("혼합 동(중위)", panel[(panel["comm_idx"] > q1) & (panel["comm_idx"] < q2)]),
          ("상업 동(상위1/3)", panel[panel["comm_idx"] >= q2])]
rows = []
for lab, d in groups:
    res, n, ns, nc = twoway_fe(d, "Leq_day", ["lp_day_logrel"])
    b, se, p = res.params[1], res.bse[1], res.pvalues[1]
    star = "***" if p < .001 else "**" if p < .01 else "*" if p < .05 else ""
    print(f"  {lab}: β={b:+.3f} se={se:.3f} p={p:.2e} {star}  (n={n:,} 센서={ns} 동={nc}) | 이동량30%↓→{b*-0.357:+.2f}dB")
    rows.append([lab, b, se, p, n, ns, nc])

# 상호작용: 효과가 상업도에 따라 증가하는가
panel["comm_z"] = (panel["comm_idx"] - comm.mean()) / comm.std()
panel["dose_x_comm"] = panel["lp_day_logrel"] * panel["comm_z"]
res, n, ns, nc = twoway_fe(panel, "Leq_day", ["lp_day_logrel", "dose_x_comm"])
print(f"\n=== 상호작용 (전체, 센서+날짜 FE) ===  n={n:,}")
for nm, i in [("lp_day_logrel(main)", 1), ("dose×상업도(interaction)", 2)]:
    b, se, p = res.params[i], res.bse[i], res.pvalues[i]
    star = "***" if p < .001 else "**" if p < .01 else "*" if p < .05 else ""
    print(f"  {nm:24s} β={b:+.3f} se={se:.3f} p={p:.2e} {star}")
print("  >> interaction 양(+)·유의 = 상업도 높을수록 dose-response 강함 (가설지지)")

# 주별 집계 robustness (측정오차 평활 -> SE 정밀화)
print("\n=== robustness: 동×주 집계 (dong+week FE) ===")
panel["week"] = panel["date"].dt.to_period("W").dt.start_time
dw = panel.groupby(["dose_key", "week"]).agg(
    Leq_day=("Leq_day", "mean"), lp_day_logrel=("lp_day_logrel", "mean")).reset_index()
res, n, ns, nc = twoway_fe(dw, "Leq_day", ["lp_day_logrel"], g1="dose_key", g2="week", clust="dose_key")
b, se, p = res.params[1], res.bse[1], res.pvalues[1]
print(f"  동×주: β={b:+.3f} se={se:.3f} p={p:.2e}  (n={n:,} 동={ns}) | 이동량30%↓→{b*-0.357:+.2f}dB")

pd.DataFrame(rows, columns=["group", "beta", "se", "pval", "n", "sensors", "dongs"]).to_csv(
    os.path.join(P, "phase2e_heterogeneity.csv"), index=False, encoding="utf-8-sig")
print("\n-> phase2e_heterogeneity.csv")
