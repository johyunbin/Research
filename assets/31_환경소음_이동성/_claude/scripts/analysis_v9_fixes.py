# -*- coding: utf-8 -*-
# Codex 게이트 반영 재계산 3종:
#  A. M1 재추정 — own-period dose·동클러스터 SE (Table 5 정정용)
#  B. 동(dong) 단위 placebo 순열 300회 — 날짜 안에서 동별 dose를 동들 사이에 셔플(센서엔 동일 브로드캐스트)
#  C. 동-일 동일가중 M2 민감도 — sensor-weighted vs dong-equal-weight
# 출력: figures_v2/cache/v9fix_results.json + placebo_betas_dong.csv
import os, sys, json
import numpy as np
import pandas as pd
import statsmodels.api as sm
sys.path.insert(0, os.path.dirname(__file__))
import figstyle_v2 as FS

rng = np.random.default_rng(20260719)
OUT = os.path.join(FS.CACHE, "v9fix_results.json")

pan = pd.read_csv(os.path.join(FS.P, "analysis_panel.csv"), dtype={"serial": str, "dose_key": str},
                  usecols=["serial", "dose_key", "date", "Leq_day", "Leq_night", "Leq24",
                           "lp_day_logrel", "lp_night_logrel", "lp_mean_logrel",
                           "temp_mean", "precip_mm", "wind_max_ms", "rain_flag",
                           "is_weekend", "holiday"])
res = {}

# ---------- A. M1 재추정 (센서 FE demean · own dose · 동클러스터) ----------
def m1(ycol, xcol):
    d = pan.dropna(subset=[ycol, xcol]).copy()
    d["temp_sq"] = d["temp_mean"] ** 2
    cols = [xcol, "temp_mean", "temp_sq", "precip_mm", "wind_max_ms", "rain_flag",
            "is_weekend", "holiday"]
    g = pd.factorize(d["serial"].values)[0]; n = g.max() + 1; c = np.bincount(g, minlength=n)
    M = d[[ycol] + cols].astype(float).values
    for j in range(M.shape[1]):
        M[:, j] -= (np.bincount(g, M[:, j], n) / c)[g]
    X = sm.add_constant(M[:, 1:], has_constant="add")
    r = sm.OLS(M[:, 0], X).fit(cov_type="cluster", cov_kwds={"groups": d["dose_key"].values})
    out = {}
    for k, name in enumerate(cols):
        out[name] = {"b": float(r.params[k + 1]), "se": float(r.bse[k + 1]), "p": float(r.pvalues[k + 1])}
    out["_n"] = int(len(d)); out["_sensors"] = int(d["serial"].nunique())
    out["_dongs"] = int(d["dose_key"].nunique())
    out["_r2w"] = float(r.rsquared)
    return out

res["M1_day"] = m1("Leq_day", "lp_day_logrel")
res["M1_night"] = m1("Leq_night", "lp_night_logrel")
res["M1_24"] = m1("Leq24", "lp_mean_logrel")
print("A. M1(동클러스터·own dose):")
for k in ("M1_day", "M1_night", "M1_24"):
    m = res[k]; b = m[list(m.keys())[0]]
    dose = "lp_day_logrel" if k == "M1_day" else ("lp_night_logrel" if k == "M1_night" else "lp_mean_logrel")
    print(f"  {k}: β={m[dose]['b']:+.3f} (SE {m[dose]['se']:.3f}, p={m[dose]['p']:.4f}) n={m['_n']:,}")

# ---------- 공통: 양방향 demean 준비 (주간) ----------
d = pan.dropna(subset=["Leq_day", "lp_day_logrel"]).reset_index(drop=True)
g1 = pd.factorize(d["serial"].values)[0]; n1 = g1.max() + 1; c1 = np.bincount(g1, minlength=n1)
g2, date_keys = pd.factorize(d["date"].values); n2 = g2.max() + 1; c2 = np.bincount(g2, minlength=n2)

def demean(v, iters=15):
    v = v.astype(float).copy()
    for _ in range(iters):
        v -= (np.bincount(g1, v, n1) / c1)[g1]
        v -= (np.bincount(g2, v, n2) / c2)[g2]
    return v

y = d["Leq_day"].values.astype(float)
x = d["lp_day_logrel"].values.astype(float)
yt = demean(y); xt = demean(x)
beta_hat = float(np.cov(yt, xt)[0, 1] / np.var(xt))
print(f"\n기준 β 재현: {beta_hat:+.4f}")

# ---------- B. 동 단위 placebo 순열 ----------
# 동-일 테이블 (dose는 동-일 상수)
dd = d[["dose_key", "date", "lp_day_logrel"]].drop_duplicates(["dose_key", "date"]).reset_index(drop=True)
dd_date = pd.factorize(dd["date"].values)[0]
order = np.argsort(dd_date, kind="stable")
dd_sorted = dd.iloc[order].reset_index(drop=True)
seg_date = pd.factorize(dd_sorted["date"].values)[0]
seg_starts = np.searchsorted(seg_date, np.arange(seg_date.max() + 1))
seg_ends = np.append(seg_starts[1:], len(dd_sorted))
dose_sorted = dd_sorted["lp_day_logrel"].values.astype(float)
# sensor-day → dong-day(sorted) 인덱스
key_map = {k: i for i, k in enumerate(zip(dd_sorted["dose_key"].values, dd_sorted["date"].values))}
row_idx = np.fromiter((key_map[(k, t)] for k, t in zip(d["dose_key"].values, d["date"].values)),
                      dtype=np.int64, count=len(d))
NPERM = 300
betas = np.empty(NPERM)
perm_dose = dose_sorted.copy()
for i in range(NPERM):
    for s, e in zip(seg_starts, seg_ends):
        seg = dose_sorted[s:e]
        perm_dose[s:e] = seg[rng.permutation(e - s)]
    xp = perm_dose[row_idx]
    xpt = demean(xp)
    betas[i] = np.cov(yt, xpt)[0, 1] / np.var(xpt)
    if (i + 1) % 50 == 0:
        print(f"  dong-perm {i+1}/{NPERM}")
p_perm = float((np.sum(np.abs(betas) >= abs(beta_hat)) + 1) / (NPERM + 1))
pd.DataFrame({"beta": betas}).to_csv(os.path.join(FS.CACHE, "placebo_betas_dong.csv"), index=False)
res["placebo_dong"] = {"null_mean": float(betas.mean()), "null_sd": float(betas.std()),
                       "p": p_perm, "n_perm": NPERM, "beta_hat": beta_hat}
print(f"B. 동 단위 순열: null {betas.mean():+.4f}±{betas.std():.4f} · p={p_perm:.4f}")

# ---------- C. 동-일 동일가중 M2 민감도 ----------
dg = d.groupby(["dose_key", "date"], as_index=False).agg(y=("Leq_day", "mean"),
                                                         x=("lp_day_logrel", "first"))
h1 = pd.factorize(dg["dose_key"].values)[0]; m1n = h1.max() + 1; cc1 = np.bincount(h1, minlength=m1n)
h2 = pd.factorize(dg["date"].values)[0]; m2n = h2.max() + 1; cc2 = np.bincount(h2, minlength=m2n)
Y = dg["y"].values.astype(float); X2 = dg["x"].values.astype(float)
for _ in range(20):
    Y -= (np.bincount(h1, Y, m1n) / cc1)[h1]; Y -= (np.bincount(h2, Y, m2n) / cc2)[h2]
    X2 -= (np.bincount(h1, X2, m1n) / cc1)[h1]; X2 -= (np.bincount(h2, X2, m2n) / cc2)[h2]
r = sm.OLS(Y, sm.add_constant(X2, has_constant="add")).fit(
    cov_type="cluster", cov_kwds={"groups": dg["dose_key"].values})
res["dong_equal_weight"] = {"b": float(r.params[1]), "se": float(r.bse[1]),
                            "p": float(r.pvalues[1]), "n": int(len(dg))}
print(f"C. 동 동일가중 M2: β={r.params[1]:+.4f} (SE {r.bse[1]:.4f}, p={r.pvalues[1]:.4f}) n={len(dg):,}")

with open(OUT, "w", encoding="utf-8") as f:
    json.dump(res, f, ensure_ascii=False, indent=1)
print("저장:", OUT)
