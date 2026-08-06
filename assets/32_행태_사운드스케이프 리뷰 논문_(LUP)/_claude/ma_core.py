# -*- coding: utf-8 -*-
"""
Paper32 — 메타분석 핵심 추정기 (단일 출처)

★ 왜 이 파일이 생겼는가
  같은 τ² 추정 함수가 `ma_v2.py`·`ma_supplementary_analyses.py`·`ma_sensitivity_v2.py`
  **세 곳에 복제**돼 있었고, 셋 다 이름은 `reml_tau2`인데 실제로는 **ML score equation**을
  풀고 있었다(독립 검증에서 적발, 2026-08-06). 원고는 "REML"이라고 보고하고 있었으므로
  사실과 다른 진술이었다. 복제가 원인이므로 여기 하나만 둔다.

  구 구현의 고정점:  Σ wᵢ²(yᵢ−μ̂)² = Σ wᵢ                     ← ML
  올바른 REML     :  Σ wᵢ²(yᵢ−μ̂)² = Σ wᵢ − (Σ wᵢ²)/(Σ wᵢ)     ← 상수항 모형의 REML

  구 추정치는 참 REML보다 τ²를 MA1 33%·MA2 80%·MA3 68%·MA4 17% 작게 잡았고,
  그 결과 예측구간이 실제보다 좁게 보고됐다.
"""
import math
import numpy as np
from scipy import stats


def reml_tau2(y, v, iters=2000, tol=1e-13):
    """상수항 랜덤효과 모형의 REML τ² (Fisher-scoring 스타일 반복)."""
    y = np.asarray(y, float); v = np.asarray(v, float)
    k = len(y)
    if k < 2:
        return 0.0
    t2 = max(float(np.var(y, ddof=1) - np.mean(v)), 0.0)
    for _ in range(iters):
        w = 1.0 / (v + t2)
        sw, sw2 = float(np.sum(w)), float(np.sum(w ** 2))
        mu = float(np.sum(w * y) / sw)
        # REML score: Σw²(y−μ)² − [Σw − Σw²/Σw] = 0  를 τ² 에 대해 갱신
        num = float(np.sum(w ** 2 * ((y - mu) ** 2 - v))) + sw2 / sw
        new = max(num / sw2, 0.0)
        if abs(new - t2) < tol:
            t2 = new
            break
        t2 = new
    return t2


def pool(rows, back_r=False, key_y="g", key_v="v"):
    """REML + Hartung-Knapp 랜덤효과 풀링.

    HK 는 unmodified 형(q<1 이면 분산을 축소)을 쓴다 — 사전 규약이 정한 형태다.
    q<1 인 클러스터는 `hk_q` 로 표면화해 보수적 대안(modified HK)과 함께 보고한다.
    """
    y = np.array([r[key_y] for r in rows], float)
    v = np.array([r[key_v] for r in rows], float)
    k = len(y)
    if k == 0:
        return None
    if k == 1:
        est, se = float(y[0]), math.sqrt(float(v[0]))
        o = dict(k=1, est=est, lo=est - 1.96 * se, hi=est + 1.96 * se,
                 p=float(2 * (1 - stats.norm.cdf(abs(est / se)))), tau2=0.0, I2=0.0, Q=0.0,
                 hk_q=float("nan"), pooled=False)
    else:
        t2 = reml_tau2(y, v)
        w = 1 / (v + t2)
        mu = float(np.sum(w * y) / np.sum(w))
        se = math.sqrt(1 / np.sum(w))
        qhk = float(np.sum(w * (y - mu) ** 2) / (k - 1))
        se_hk = se * math.sqrt(qhk)
        tc = stats.t.ppf(0.975, k - 1)
        wf = 1 / v
        muf = float(np.sum(wf * y) / np.sum(wf))
        Q = float(np.sum(wf * (y - muf) ** 2))
        o = dict(k=k, est=mu, lo=mu - tc * se_hk, hi=mu + tc * se_hk,
                 p=float(2 * (1 - stats.t.cdf(abs(mu / se_hk), k - 1))), tau2=t2,
                 I2=max(0.0, (Q - (k - 1)) / Q) * 100 if Q > 0 else 0.0, Q=Q,
                 hk_q=qhk, se_hk=se_hk, pooled=True)
    if back_r:
        o["r"] = math.tanh(o["est"]); o["r_lo"] = math.tanh(o["lo"]); o["r_hi"] = math.tanh(o["hi"])
    return o


def prediction_interval(res):
    """Higgins–Thompson–Spiegelhalter 95% 예측구간. 자유도는 k−2."""
    k = res["k"]
    if k < 3:
        return None, None
    se = res.get("se_hk") or math.sqrt(max(res["tau2"], 0) / k)
    spread = math.sqrt(res["tau2"] + se ** 2)
    tp = stats.t.ppf(0.975, k - 2)
    return res["est"] - tp * spread, res["est"] + tp * spread


def combine_within_study(effects, rho=0.5):
    """규약 §3 — 동급 다중 효과는 논문 내 평균(효과 간 상관 ρ 가정, Borenstein).

    ȳ = mean(yᵢ) · Var(ȳ) = (1/m²)[Σvᵢ + 2ρ ΣΣ_{i<j} √(vᵢvⱼ)]
    """
    ys = np.array([e[0] for e in effects], float)
    vs = np.array([e[1] for e in effects], float)
    m = len(ys)
    if m == 1:
        return float(ys[0]), float(vs[0])
    cross = sum(math.sqrt(vs[i] * vs[j]) for i in range(m) for j in range(i + 1, m))
    var = (float(np.sum(vs)) + 2 * rho * cross) / (m ** 2)
    return float(np.mean(ys)), var


def var_from_ci(r_lo, r_hi, z=1.96):
    """보고된 상관 CI 로부터 Fisher-z 분산을 역산.

    원문이 n 을 명시하지 않고 CI 만 주는 경우, n 을 추측하는 것보다 CI 를 그대로 쓰는 편이
    정확하고 검증 가능하다(1177 이 이 경우다).
    """
    zl, zh = math.atanh(r_lo), math.atanh(r_hi)
    return ((zh - zl) / (2 * z)) ** 2
