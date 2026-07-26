# -*- coding: utf-8 -*-
# Phase 2-A dose-response: 센서 고정효과 within-회귀 (demean by serial) + 센서 clustered SE.
# 결과 = Leq_day/night/24, 주 dose = 동 주간 생활인구 상대변화(lp_day_logrel). 통제 = 기상·요일·공휴일·계절.
# within(센서 demean)이 센서별 검교정 오프셋을 상쇄 -> 비검교정 dB의 상대변화만 사용(설계 핵심).
import os
import numpy as np
import pandas as pd
import statsmodels.api as sm

ROOT = r"T:\00_BACKUP\01_한양대학교\00_연구실\00_프로젝트\★_논문화 프로젝트\31_환경소음_이동성"
P = os.path.join(ROOT, "data", "processed")
OUTDIR = os.path.join(ROOT, "data", "processed")
panel = pd.read_csv(os.path.join(P, "analysis_panel.csv"), dtype={"serial": str, "dose_key": str})
panel["date"] = pd.to_datetime(panel["date"])
panel["temp_sq"] = panel["temp_mean"] ** 2

def within_fe(d, ycol, xcols, group="serial"):
    """센서 demean within estimator + group-clustered SE."""
    d = d.dropna(subset=[ycol] + xcols + [group]).copy()
    g = d.groupby(group)
    Y = d[ycol] - g[ycol].transform("mean")
    X = pd.DataFrame({c: d[c] - g[c].transform("mean") for c in xcols})
    X = sm.add_constant(X, has_constant="add")
    res = sm.OLS(Y.values, X.values).fit(cov_type="cluster", cov_kwds={"groups": d[group].values})
    res._xnames = ["const"] + xcols
    res._nobs = len(d); res._ngroups = d[group].nunique()
    return res

def show(res, title, key="lp_day_logrel"):
    names = res._xnames
    print(f"\n=== {title} ===  (n={res._nobs:,}, 센서={res._ngroups}, R2_within={res.rsquared:.3f})")
    for nm, b, se, p in zip(names, res.params, res.bse, res.pvalues):
        star = "***" if p < .001 else "**" if p < .01 else "*" if p < .05 else ""
        mark = " <--dose" if nm == key else ""
        print(f"  {nm:16s} b={b:+8.3f}  se={se:.3f}  p={p:.2e} {star}{mark}")
    if key in names:
        b = res.params[names.index(key)]
        print(f"  >> 해석: 주간 이동량 30%감소(logrel-0.357) -> Leq_day {b*-0.357:+.2f} dB  (β={b:+.3f} dB/log-unit)")

WX = ["temp_mean", "temp_sq", "precip_mm", "wind_max_ms", "rain_flag", "is_weekend", "holiday"]

print(f"패널 {len(panel):,} sensor-days 로드. 분석 대상 필터링...")
# 주간 모델: 동 주간 생활인구 dose
m1 = within_fe(panel, "Leq_day", ["lp_day_logrel"] + WX)
show(m1, "주간 Leq_day ~ 동 주간이동량(lp_day_logrel) + 기상/요일/공휴일 [센서FE]")

m2 = within_fe(panel, "Leq_night", ["lp_night_logrel" if "lp_night_logrel" in panel else "lp_day_logrel"] + WX)
show(m2, "야간 Leq_night ~ 이동량 + 통제 [센서FE]", key="lp_day_logrel")

m3 = within_fe(panel, "Leq24", ["lp_mean_logrel"] + WX, )
show(m3, "전일 Leq24 ~ 동 전일이동량(lp_mean_logrel) + 통제 [센서FE]", key="lp_mean_logrel")

# 보조 dose: 도시 지하철 + 거리두기 stringency (city-level)
m4 = within_fe(panel, "Leq_day", ["subway_rel", "stringency"] + WX)
show(m4, "주간 Leq_day ~ 지하철상대 + 거리두기stringency + 통제 [센서FE]", key="subway_rel")

# 복합: 동 이동량 + 도시 보조 동시
m5 = within_fe(panel, "Leq_day", ["lp_day_logrel", "subway_rel", "stringency"] + WX)
show(m5, "주간 Leq_day ~ 동이동량 + 지하철 + stringency [센서FE, 복합]")

# 결과표 저장
rows = []
for tag, res, key in [("Leq_day~lp_day", m1, "lp_day_logrel"), ("Leq_night~lp", m2, "lp_day_logrel"),
                       ("Leq24~lp_mean", m3, "lp_mean_logrel"), ("Leq_day~subway+string", m4, "subway_rel"),
                       ("Leq_day~combined", m5, "lp_day_logrel")]:
    nm = res._xnames
    for v in nm:
        i = nm.index(v)
        rows.append([tag, v, res.params[i], res.bse[i], res.pvalues[i], res._nobs, res._ngroups, res.rsquared])
pd.DataFrame(rows, columns=["model", "term", "coef", "se", "pval", "n", "sensors", "r2_within"]).to_csv(
    os.path.join(OUTDIR, "phase2a_doseresponse_results.csv"), index=False, encoding="utf-8-sig")
print(f"\n-> phase2a_doseresponse_results.csv 저장")
