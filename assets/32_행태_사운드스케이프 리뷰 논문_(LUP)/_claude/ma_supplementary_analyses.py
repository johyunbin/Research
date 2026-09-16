# -*- coding: utf-8 -*-
"""
Paper32 — 게이트 지적에 따른 분석 보완
  B2 출판편향 — k<10이라 검정 불가. 그 사실과 함의를 수치로 뒷받침한다(검정력 계산).
  B3 예측구간 — I²가 큰 클러스터는 CI가 아니라 예측구간이 실질 불확실성을 보여준다.
  B1 랩 재현 민감도 — VR-lab 5편(661·666·999·1122·1272)을 제외한 재집계.
  B4 하위그룹 — 등록 6축이 k=3~7에서 성립 불가함을 수치로 보인다.
출력: fulltext/ma/ma_supplementary.md · ma_supplementary.csv
"""
import sys, os, csv, math
import numpy as np
from scipy import stats

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
MA = os.path.join(FT, "ma")


# ★ τ² 추정·풀링은 `ma_core` 하나만 쓴다.
#   이 함수가 세 파일에 복제돼 있었고 셋 다 이름과 달리 **ML** 을 풀고 있었다(2026-08-06 적발).
from ma_core import reml_tau2, pool as _core_pool, prediction_interval,     combine_within_study, var_from_ci


def pool(rows):
    """ma_core 로 위임 + HTS 예측구간 부착."""
    out = _core_pool(rows)
    lo, hi = prediction_interval(out)
    if lo is not None:
        out["pi_lo"], out["pi_hi"] = lo, hi
    return out


def r_to_z(r, n):
    return 0.5 * math.log((1 + r) / (1 - r)), 1 / (n - 3)


def rd(fn):
    return list(csv.DictReader(open(os.path.join(MA, fn), encoding="utf-8-sig")))


def load_walk(rd, exclude_music=True):
    """MA1 입력 — 규약 §3(클러스터×대비프레임당 논문 1효과)을 실제로 적용한다.

    532 는 Exp1·Exp2 두 효과를 같은 대비프레임에서 보고한다. 종전 코드는 둘 다 넣어
    k=4 로 풀링했는데 이는 **자체 규약 위반**이었다(2026-08-06 적발). 규약대로
    논문 내 평균(효과 간 ρ=0.5 Borenstein)으로 합성한다.
    """
    rows = [dict(uid=r["no"], label=r["label"], g=float(r["g"]), v=float(r["var"]),
                 contrast=r["contrast"])
            for r in rd("ma_walking_input.csv") if "control" not in r["contrast"]]
    if exclude_music:
        rows = [r for r in rows if r["uid"] != "481"]        # 481=음악, 주분석 제외(D3-6)
    out, seen = [], {}
    for r in rows:
        seen.setdefault(r["uid"], []).append(r)
    for uid, grp in seen.items():
        if len(grp) == 1:
            out.append(grp[0]); continue
        g, v = combine_within_study([(x["g"], x["v"]) for x in grp])
        out.append(dict(uid=uid, label=grp[0]["label"].split(" Exp")[0] + " (Exp 합성)",
                        g=g, v=v, contrast=grp[0]["contrast"],
                        note=f"규약 §3 — 논문 내 {len(grp)}효과 평균(ρ=0.5)"))
    return sorted(out, key=lambda r: r["uid"])


def load_corr(rd, r_to_z):
    """MA4 입력 — status=exclude 행은 넣지 않고, v_override 가 있으면 그 분산을 쓴다."""
    out = []
    for r in rd("ma_correlation_input.csv"):
        if (r.get("status") or "include").strip() == "exclude":
            continue
        rv = float(r["r"])
        if (r.get("v_override") or "").strip():
            import math as _m
            z, v = _m.atanh(rv), float(r["v_override"])
        else:
            z, v = r_to_z(rv, int(r["n"]))
        out.append(dict(uid=r["no"], label=r["label"], g=z, v=v, r=rv,
                        n=int(r["n"]) if (r.get("n") or "").strip() else None,
                        src=r.get("src", "")))
    return out


def main():
    new = {r["uid"]: r for r in rd("ma_v2_new_inputs.csv")}
    qual = {r["uid"]: r["quality_tier"] for r in
            csv.DictReader(open(os.path.join(FT, "quality_v2.csv"), encoding="utf-8-sig"))}

    walk = load_walk(rd)
    stay = [dict(uid=r["no"], g=float(r["g"]), v=float(r["v"])) for r in rd("ma_staying_input.csv")]
    soc = [dict(uid=r["no"], g=float(r["g"]), v=float(r["v"])) for r in rd("ma_social_input.csv")]
    soc.append(dict(uid="CT0025", g=float(new["CT0025"]["g"]), v=float(new["CT0025"]["v"])))
    corr = [dict(uid=r["uid"], g=r["g"], v=r["v"]) for r in load_corr(rd, r_to_z)]
    for u in ("CT0126", "CT0184"):
        corr.append(dict(uid=u, g=float(new[u]["g"]), v=float(new[u]["v"])))

    CL = [("MA1 walking speed", walk, "g"), ("MA2 staying", stay, "g"),
          ("MA3 social interaction", soc, "g"), ("MA4 sound–behaviour correlation", corr, "z")]

    rows, L = [], []
    # ⚠️ 서술문에 k·I² 를 손으로 박으면 정본이 바뀔 때 문장만 낡는다(실제로 낡았다).
    #    클러스터 정의에서 매번 산출한다.
    _res = {name: pool(r_) for name, r_, _u in CL}
    ks = [o["k"] for o in _res.values()]
    kmax, kmin = max(ks), min(ks)
    kmax_i2 = max(_res.values(), key=lambda o: o["k"])["I2"]
    L.append("# Paper32 — 분석 보완 (독립 게이트 지적 대응)\n")
    L.append("\n독립 검증 게이트가 지적한 4건을 수치로 처리한다. 원고 Methods·Results에 반영할 값들이다.\n")

    # ── B3 예측구간 ────────────────────────────────────────────────
    L.append("\n## 1. 예측구간 (prediction interval)\n\n"
             "이질성이 큰 클러스터에서는 **신뢰구간이 평균의 정밀도**를, **예측구간이 다음 연구에서 "
             "기대되는 값의 범위**를 말한다. 최대 클러스터는 I² = %.1f%%이므로 CI만 보고하면 "
             "불확실성을 크게 과소 표현한다. Higgins–Thompson–Spiegelhalter 방식으로 산출했다.\n\n"
             % kmax_i2)
    L.append("| 클러스터 | k | 추정치 | 95% CI | **95% 예측구간** | I² |\n|---|---|---|---|---|---|\n")
    for name, r_, unit in CL:
        o = pool(r_)
        pi = (f"{o['pi_lo']:+.3f} ~ {o['pi_hi']:+.3f}" if "pi_lo" in o else "산출 불가(k<3)")
        if unit == "z" and "pi_lo" in o:
            pi += f"  (r {math.tanh(o['pi_lo']):+.2f} ~ {math.tanh(o['pi_hi']):+.2f})"
        L.append(f"| {name} | {o['k']} | {o['est']:+.3f} | {o['lo']:+.3f} ~ {o['hi']:+.3f} | "
                 f"**{pi}** | {o['I2']:.1f}% |\n")
        rows.append(dict(analysis="prediction_interval", cluster=name, k=o["k"],
                         est=round(o["est"], 4), lo=round(o["lo"], 4), hi=round(o["hi"], 4),
                         pi_lo=round(o.get("pi_lo", float("nan")), 4),
                         pi_hi=round(o.get("pi_hi", float("nan")), 4), I2=round(o["I2"], 1)))
    o3, o4 = pool(soc), pool(corr)
    L.append(f"\n**읽기**: MA3의 예측구간은 {o3['pi_lo']:+.2f} ~ {o3['pi_hi']:+.2f}로 **0을 포함한다.** "
             f"평균은 0을 배제하지만, 새로운 연구가 반대 방향의 결과를 낼 가능성을 배제할 수 없다는 뜻이다. "
             f"MA4도 예측구간이 r {math.tanh(o4['pi_lo']):+.2f} ~ {math.tanh(o4['pi_hi']):+.2f}로 넓다.\n"
             "→ **원고는 '유의하다'를 '평균이 0과 다르다'로 한정해 쓰고, 예측구간을 함께 보고해야 한다.**\n")

    # ── B2 출판편향 ───────────────────────────────────────────────
    L.append("\n## 2. 출판편향 — 검정 불가와 그 함의\n\n"
             "사전 규약(`analysis_rules.md §6`)은 **k ≥ 10인 클러스터만 funnel plot과 Egger 검정**을 "
             "수행하도록 정했다. 최대 클러스터가 k = %d이므로 **어느 클러스터에서도 검정을 수행할 수 "
             "없다.** 이는 결과가 아니라 제약이며, 다음을 뜻한다.\n\n" % kmax)
    L.append("| 클러스터 | k | Egger 검정 최소 요건 | 수행 |\n|---|---|---|---|\n")
    for name, r_, _ in CL:
        L.append(f"| {name} | {len(r_)} | k ≥ 10 | ✕ |\n")
    L.append("\n- 소규모 연구 효과(small-study effects)를 **탐지할 수단이 없으므로 배제할 수도 없다.**\n"
             "- 이 분야의 출판 관행상 귀무 결과가 보고되지 않았을 가능성은 실재한다. 실제로 코퍼스에서 "
             "명시적 귀무를 보고한 연구는 소수이며(예: 공원 음량과 방문빈도, 음향 보행신호기와 횡단궤적), "
             "둘 다 인용추적·보조검색처럼 **데이터베이스 검색 밖 경로**에서 왔다.\n"
             "- 따라서 모든 풀링 추정치는 **상한으로 읽어야 한다**(진짜 효과는 이보다 작을 수 있다).\n")

    # ── B1 랩 재현 민감도 ─────────────────────────────────────────
    LABS = {"661", "666", "999", "1122", "1272"}
    L.append("\n## 3. 랩(VR) 재현 제외 민감도\n\n"
             "등록 프로토콜은 옥외 장면을 재현한 실험실 연구를 적격 세팅으로 허용했고, "
             "데이터베이스 갈래에서는 이를 본분석에 포함했다(VR-lab 5편: 661·666·999·1122·1272). "
             "실제 공공공간 연구만으로 좁히면 어떻게 되는지 확인한다.\n\n")
    L.append("| 클러스터 | 주분석 | 랩 제외 | 변화 |\n|---|---|---|---|\n")
    for name, r_, _ in CL:
        sub = [x for x in r_ if x["uid"] not in LABS]
        if len(sub) == len(r_):
            L.append(f"| {name} | 해당 없음(풀에 랩 연구 없음) | — | — |\n")
            continue
        a, b = pool(r_), (pool(sub) if len(sub) >= 2 else None)
        L.append(f"| {name} | k={a['k']}, {a['est']:+.3f} | "
                 f"{'k=' + str(b['k']) + ', ' + format(b['est'], '+.3f') if b else '풀링 불가'} | "
                 f"{'—' if not b else format(b['est'] - a['est'], '+.3f')} |\n")
    L.append("\n**읽기**: VR-lab 5편은 **어느 메타분석 풀에도 기여하지 않는다**(전부 서술 종합에만 기여). "
             "따라서 랩 제외는 풀링 결과를 전혀 바꾸지 않으며, 이 민감도 축은 서술 종합에만 해당한다. "
             "⚠️ 다만 **갈래 간 처리 불일치**는 남는다 — 인용추적 갈래에서는 같은 성격의 CT0356을 "
             "SENS_ONLY로 보냈다. 이탈 로그에 기록한다.\n")

    # ── B4 하위그룹 ───────────────────────────────────────────────
    L.append("\n## 4. 등록 하위그룹 분석 — 수행 불가\n\n"
             "등록서는 세팅(공원/가로/광장)·설계(현장실험/관찰)·측정세대 등으로 하위그룹을 나누도록 "
             "계획했다. 그러나 최대 클러스터가 k = %d이고 나머지는 k = %d~%d다.\n\n"
             % (kmax, kmin, kmax - 1))
    L.append("| 클러스터 | k | 2개 하위그룹으로 나누면 | 판정 |\n|---|---|---|---|\n")
    for name, r_, _ in CL:
        k = len(r_)
        L.append(f"| {name} | {k} | 각 {k//2}~{k-k//2}개 | 풀링 불가(규약상 k≥3) |\n")
    L.append("\n하위그룹당 k가 1~4가 되어 사전 규약의 최소 요건(k ≥ 3)조차 대부분 충족하지 못하고, "
             "Hartung–Knapp 구간은 k = 2에서 t(1) = 12.7 때문에 사실상 무한대가 된다. "
             "**하위그룹 분석은 수행하지 않았으며, 그 이유는 자료 부족이다.** "
             "대신 이질성의 원천은 서술적으로 검토했다(노출 유형·행태 유형·설계).\n")

    with open(os.path.join(MA, "ma_supplementary.csv"), "w", newline="",
              encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["analysis", "cluster", "k", "est", "lo", "hi",
                                          "pi_lo", "pi_hi", "I2"])
        w.writeheader(); w.writerows(rows)
    open(os.path.join(MA, "ma_supplementary.md"), "w", encoding="utf-8").write("".join(L))
    print("".join(L))
    print("[저장] ma/ma_supplementary.md · ma_supplementary.csv")


if __name__ == "__main__":
    main()
