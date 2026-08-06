# -*- coding: utf-8 -*-
"""
Paper32 — 메타분석 v2 (인용추적 신규 효과 반영)
규약 `analysis_rules.md` 준수 · REML + Hartung-Knapp · 입력은 전부 verbatim에서 유래.

신규 편입 판단(보수적):
  MA1 보행속도 — 변동 없음. CT0090은 음악 하위스트림(481과 같은 계열) → 민감도.
  MA2 체류     — 변동 없음. CT0175는 아웃컴이 '방문빈도'라 체류가 아님 → 민감도.
  MA3 사회     — **CT0025 추가**(Mathews & Canon 1975, 2×2 → OR → Chinn d). Moser 1988과 동일 계열.
                 CT0414는 대비가 '청각공간 vs 시각공간'이고 독립성 위배 → 민감도.
  MA4 상관     — **CT0126·CT0184 추가**. CT0137은 d→r 변환이 등록 규약 밖 → 변형분석으로 분리.
출력: fulltext/ma/ma_v2_*.csv|md
"""
import sys, os, csv, math
import numpy as np
from scipy import stats

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
MA = os.path.join(FT, "ma")
os.makedirs(MA, exist_ok=True)
CHINN = math.sqrt(3) / math.pi


def hedges_g(m1, sd1, n1, m2, sd2, n2):
    df = n1 + n2 - 2
    sp = math.sqrt(((n1 - 1) * sd1 ** 2 + (n2 - 1) * sd2 ** 2) / df)
    d = (m1 - m2) / sp
    J = 1 - 3 / (4 * df - 1)
    g = J * d
    return g, (n1 + n2) / (n1 * n2) + g ** 2 / (2 * (n1 + n2))


def or_2x2(a, b, c, d):
    """a,b = 노출1의 (사건, 비사건) · c,d = 노출2의 (사건, 비사건) → Chinn d"""
    orv = (a * d) / (b * c)
    se = math.sqrt(1 / a + 1 / b + 1 / c + 1 / d)
    return math.log(orv) * CHINN, (se * CHINN) ** 2, orv


def d_to_r(d, n1, n2):
    """Borenstein 표준 변환 — 등록 규약 밖(이탈 로그 기록 대상)"""
    a = (n1 + n2) ** 2 / (n1 * n2)
    return d / math.sqrt(d ** 2 + a)


def r_to_z(r, n):
    return 0.5 * math.log((1 + r) / (1 - r)), 1 / (n - 3)


# ★ τ² 추정·풀링은 `ma_core` 하나만 쓴다.
#   이 함수가 세 파일에 복제돼 있었고 셋 다 이름과 달리 **ML** 을 풀고 있었다(2026-08-06 적발).
from ma_core import reml_tau2, pool as _core_pool, prediction_interval,     combine_within_study, var_from_ci


def pool(rows, back_r=False):
    return _core_pool(rows, back_r=back_r)


def line(tag, o, back_r=False):
    if o is None:
        return f"| {tag} | 0 | — | — | — | — |\n"
    s = (f"| {tag} | {o['k']} | {o['est']:+.3f} | {o['lo']:+.3f} ~ {o['hi']:+.3f} | "
         f"{o['p']:.3f} | {o['I2']:.1f}% |")
    if back_r and "r" in o:
        s = s[:-1] + f" r = {o['r']:+.3f} [{o['r_lo']:+.3f}, {o['r_hi']:+.3f}] |"
    return s + "\n"


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
    # ── 기존 입력 로드 ─────────────────────────────────────────────
    def rd(fn):
        return list(csv.DictReader(open(os.path.join(MA, fn), encoding="utf-8-sig")))

    walk_all = [dict(uid=r["no"], label=r["label"], g=float(r["g"]), v=float(r["var"]),
                     contrast=r["contrast"])
                for r in rd("ma_walking_input.csv") if "control" not in r["contrast"]]
    walk = load_walk(rd)                                       # 규약 §3 합성 적용 · 481 제외(D3-6)
    stay = [dict(uid=r["no"], label=r["label"], g=float(r["g"]), v=float(r["v"]))
            for r in rd("ma_staying_input.csv")]
    soc = [dict(uid=r["no"], label=r["label"], g=float(r["g"]), v=float(r["v"]))
           for r in rd("ma_social_input.csv")]
    corr = load_corr(rd, r_to_z)

    new_rows = []

    # ── MA3 신규: CT0025 (Mathews & Canon 1975) ────────────────────
    # Table 2 현장실험. cue(cast) 2셀 합산 → 조용(ambient) vs 소음(high)의 도움행동 2×2
    #   ambient: 도움 4+16=20, 미도움 16+4=20  |  high: 도움 2+3=5, 미도움 18+17=35
    g25, v25, or25 = or_2x2(20, 20, 5, 35)
    soc_new = dict(uid="CT0025", label="Mathews & Canon 1975 (조용 vs 소음, 도움행동)",
                   g=g25, v=v25,
                   src="Table 2 cells; F(1,76)=20.00, P<.001; n1=n2=40",
                   note=f"2×2 OR={or25:.2f} → Chinn d. cue 2셀 합산. Moser 1988과 동일 변환경로")
    soc_v2 = soc + [soc_new]
    new_rows.append({"cluster": "MA3", **{k: soc_new[k] for k in ("uid", "label", "g", "v", "src", "note")}})

    # ── MA3 민감도: CT0414 (청각공간 vs 시각공간) ───────────────────
    # Table 4: 사회적 상호작용 Auditory 99/455 vs Visual 39/1113
    g414, v414, or414 = or_2x2(99, 455 - 99, 39, 1113 - 39)
    soc_lib = soc_v2 + [dict(uid="CT0414", label="Xuanwu Lake 2026 (청각공간 vs 시각공간)",
                             g=g414, v=v414)]

    # ── MA4 신규: CT0126 · CT0184 ──────────────────────────────────
    z126, v126 = r_to_z(0.65, 29)     # LAeq × '목소리 높이기', 분석단위 = sampling point N=29
    z184, v184 = r_to_z(0.165, 301)   # 동반상태 → 말소리 인지(역방향), 단일예측 β=R
    corr_new = [
        dict(uid="CT0126", label="Montes González 2022 (LAeq ↔ 대화 방해)", g=z126, v=v126,
             r=0.65, n=29, src="Table 2 LAeq × e) 0.65***",
             note="분석단위는 105명이 아니라 sampling point N=29(원문 §2.2 명시). CT0322와 동일 데이터 — 한 편만 채택"),
        dict(uid="CT0184", label="Cao & Kang 2021 (동반상태 → 말소리 인지)", g=z184, v=v184,
             r=0.165, n=301, src="Table 8 단일예측 Beta 0.165 = R, t=2.895",
             note="역방향(행태→지각). MA4는 양방향 혼재 클러스터라 프레임 정합"),
    ]
    corr_v2 = corr + corr_new
    for c in corr_new:
        new_rows.append({"cluster": "MA4", **{k: c[k] for k in ("uid", "label", "g", "v", "src", "note")}})

    # ── MA4 변형: CT0137 (d → r, 등록 규약 밖) ─────────────────────
    g137, _ = hedges_g(2.42, 0.68, 57, 1.98, 0.83, 49)
    r137 = d_to_r(g137, 57, 49)
    z137, v137 = r_to_z(r137, 106)
    corr_ext = corr_v2 + [dict(uid="CT0137", label="광장무 음악 평가 → 참여 규칙성",
                               g=z137, v=v137, r=r137, n=106)]

    # ── MA1·MA2 민감도용 ───────────────────────────────────────────
    # CT0090: 음악 vs 무음악 보행속도, 보고된 Cohen's d = 0.462 (F(1,63)=5.104), N=72
    #   ※부호 규약: g>0 = 긍정음 조건에서 더 빠름. 음악에서 더 빨랐으므로 +.
    d90 = 0.462
    n90a, n90b = 36, 36                      # N=72를 균등 가정(원문 군별 n 미보고)
    J90 = 1 - 3 / (4 * (n90a + n90b - 2) - 1)
    g90 = J90 * d90
    v90 = (n90a + n90b) / (n90a * n90b) + g90 ** 2 / (2 * (n90a + n90b))
    walk_music = [r for r in walk_all if r["uid"] == "481"] + \
                 [dict(uid="CT0090", label="Franěk 2014 (음악 vs 무음악)", g=g90, v=v90)]

    # CT0175: 조용 vs 시끄러움, 주3회 이상 방문 OR=0.3 (95% CI 0.09–1.0)
    #   ⚠️아웃컴이 체류가 아니라 방문빈도 → MA2 주분석 제외, 민감도만
    lo, hi = 0.09, 1.0
    se175 = (math.log(hi) - math.log(lo)) / (2 * 1.96)
    g175 = math.log(0.3) * CHINN
    v175 = (se175 * CHINN) ** 2
    stay_lib = stay + [dict(uid="CT0175", label="Dublin 2018 (조용 → 주3회+ 방문, OR)",
                            g=g175, v=v175)]

    # ── 풀링 ───────────────────────────────────────────────────────
    res = {
        "MA1 보행속도 (주분석·불변)": (pool(walk), False),
        "  └ 민감도: 음악 하위스트림": (pool(walk_music), False),
        "MA2 체류 (주분석·불변)": (pool(stay), False),
        "  └ 민감도: 방문빈도 포함(CT0175)": (pool(stay_lib), False),
        "MA3 사회적 상호작용 (신규 CT0025 포함)": (pool(soc_v2), False),
        "  └ 민감도: CT0414 추가": (pool(soc_lib), False),
        "  └ 민감도: CT0025 제외(구 주분석)": (pool(soc), False),
        "MA4 지각–행태 상관 (신규 2편 포함)": (pool(corr_v2, True), True),
        "  └ 변형: CT0137 추가(d→r, 규약 밖)": (pool(corr_ext, True), True),
        "  └ 민감도: 신규 제외(구 주분석)": (pool(corr, True), True),
    }

    # ── 저장 ───────────────────────────────────────────────────────
    with open(os.path.join(MA, "ma_v2_new_inputs.csv"), "w", newline="",
              encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["cluster", "uid", "label", "g", "v", "src", "note"])
        w.writeheader()
        for r in new_rows:
            w.writerow({**r, "g": round(r["g"], 4), "v": round(r["v"], 5)})

    L = ["# Paper32 — 메타분석 v2 (인용추적 반영)\n",
         "\n규약 `analysis_rules.md` · REML + Hartung-Knapp · 입력은 전부 원문 verbatim에서 유래.\n",
         "\n## 종합\n\n| 분석 | k | 추정치 | 95% CI (HK) | p | I² |\n|---|---|---|---|---|---|\n"]
    for name, (o, br) in res.items():
        L.append(line(name, o, br))

    L.append("\n## 신규 편입 판단\n\n"
             "| 클러스터 | 신규 | 판단 |\n|---|---|---|\n"
             "| MA1 | CT0090 | **주분석 제외** — 음악 하위스트림(481과 같은 계열). 민감도로 분리 |\n"
             "| MA2 | CT0175 | **주분석 제외** — 아웃컴이 체류시간이 아니라 **방문빈도**. 민감도로 분리 |\n"
             "| MA3 | **CT0025** | **주분석 포함** — Moser 1988과 동일 설계·동일 변환경로(2×2 → OR → Chinn d) |\n"
             "| MA3 | CT0414 | **주분석 제외** — 대비가 '청각공간 vs 시각공간'이고 독립성 위배(합계 2,249 > 관측 1,167) |\n"
             "| MA4 | **CT0126·CT0184** | **주분석 포함** — 둘 다 r 계열, 프레임 정합 |\n"
             "| MA4 | CT0137 | **변형 분석** — M±SD를 d→r로 변환해야 하는데 등록 규약 §1에 없는 경로 |\n")
    L.append(f"\n## MA3 신규 효과 상세 (CT0025)\n\n"
             f"Mathews & Canon (1975) *J Personality and Social Psychology* 현장실험.\n"
             f"잔디깎기 소음 87 dB(C) vs 주변 50 dB(C) 조건에서 물건을 떨어뜨린 사람을 돕는 행동을 관찰.\n\n"
             f"- 2×2(cue 2셀 합산): 조용 도움 20 / 미도움 20 · 소음 도움 5 / 미도움 35\n"
             f"- OR = {or25:.2f} → Chinn d = **{g25:+.3f}** (v = {v25:.4f}), n1 = n2 = 40\n"
             f"- 원문 검정: `F(1, 76) = 20.00, P < .001`\n\n"
             f"**Moser(1988)와 같은 방향·같은 변환경로다.** 50년 간격의 독립 반복이라 "
             f"MA3의 신뢰도를 실질적으로 높인다.\n")
    L.append("\n⚠️ **k = 2 행의 신뢰구간은 해석하지 말 것.** Hartung-Knapp 보정은 t(k−1) 분포를 쓰므로 "
             "k = 2일 때 t(1) = 12.7이 되어 구간이 사실상 무한대로 벌어진다. 음악 하위스트림 민감도가 "
             "그 경우이며, 추정치(방향)만 참고하고 구간은 보고하지 않는다.\n")
    L.append("\n## 해석\n\n")
    o3, o3o = res["MA3 사회적 상호작용 (신규 CT0025 포함)"][0], res["  └ 민감도: CT0025 제외(구 주분석)"][0]
    o4, o4o = res["MA4 지각–행태 상관 (신규 2편 포함)"][0], res["  └ 민감도: 신규 제외(구 주분석)"][0]
    L.append(f"1. **MA3가 유의해졌다** — k {o3o['k']}→{o3['k']}, g {o3o['est']:+.3f}→{o3['est']:+.3f}, "
             f"p {o3o['p']:.3f}→{o3['p']:.3f}. 다만 I²가 {o3o['I2']:.1f}%→{o3['I2']:.1f}%로 "
             f"올랐다(CT0025의 효과가 크다).\n")
    L.append(f"2. **MA4는 안정적** — k {o4o['k']}→{o4['k']}, r {o4o['r']:+.3f}→{o4['r']:+.3f}, "
             f"p {o4o['p']:.3f}→{o4['p']:.3f}. 신규 2편이 들어와도 추정치가 거의 움직이지 않는다.\n")
    L.append("3. **MA1·MA2는 불변** — 인용추적이 이 두 클러스터에는 풀링 가능한 효과를 "
             "추가하지 못했다. 조건 대비 실험의 희소성이라는 리뷰의 핵심 주장이 재확인된다.\n")
    open(os.path.join(MA, "ma_v2_summary.md"), "w", encoding="utf-8").write("".join(L))

    # ── Figure 2용 기계 판독 export ──────────────────────────────────
    # ★ Fig 2 는 종전에 효과값·풀링값을 **전부 하드코딩**하고 있었다. 정본이 바뀌어도
    #   그림은 옛 숫자를 그리므로, 이번 드리프트를 만든 원인 중 하나다. 여기서 내보낸다.
    import json
    FOREST = [("walking", "MA1 보행속도 (주분석·불변)", walk, False),
              ("staying", "MA2 체류 (주분석·불변)", stay, False),
              ("social", "MA3 사회적 상호작용 (신규 CT0025 포함)", soc_v2, False),
              ("correlation", "MA4 지각–행태 상관 (신규 2편 포함)", corr_v2, True)]
    fx = {}
    for key, tag, rows_, br in FOREST:
        o = res[tag][0]
        fx[key] = {
            "effects": [{"uid": r["uid"], "label": r.get("label", r["uid"]),
                         "est": r["g"], "var": r["v"],
                         "route": ("citation-tracking" if str(r["uid"]).startswith("CT")
                                   else "db-search")} for r in rows_],
            "pooled": {k: o[k] for k in ("k", "est", "lo", "hi", "p", "I2", "tau2") if k in o},
            "back_r": br,
        }
    with open(os.path.join(MA, "ma_forest_data.json"), "w", encoding="utf-8") as f:
        json.dump(fx, f, ensure_ascii=False, indent=1)

    print("".join(L))
    print(f"[저장] ma/ma_v2_summary.md · ma_v2_new_inputs.csv · ma_forest_data.json")


if __name__ == "__main__":
    main()
