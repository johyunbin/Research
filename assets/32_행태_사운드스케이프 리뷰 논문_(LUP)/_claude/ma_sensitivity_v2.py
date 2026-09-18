# -*- coding: utf-8 -*-
"""
Paper32 — 민감도 분석 v2 (MA3 k=4 · MA4 k=7 — 2026-09-18 80 편입 반영. 코퍼스 수는 corpus_v4_verdicts 에서 센다)
등록 규약 `analysis_rules.md §6`의 5축 + 추가 2축을 갱신된 풀에 다시 적용한다.
  ① leave-one-out ② MMAT 저품질 제외 ③ 관측 n 재계산 ④ rho 제외
  ⑤ 가정 의존 입력 제외 ⑥ MA1 풀 정의 대안 ⑦ ★신규: 갈래 제외(인용추적 빼기)
출력: fulltext/ma/ma_sensitivity_v2.csv · ma_sensitivity_v2.md
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


def pool(rows, back_r=False):
    return _core_pool(rows, back_r=back_r)


def hedges_g(m1, sd1, n1, m2, sd2, n2):
    df = n1 + n2 - 2
    sp = math.sqrt(((n1 - 1) * sd1 ** 2 + (n2 - 1) * sd2 ** 2) / df)
    d = (m1 - m2) / sp
    J = 1 - 3 / (4 * df - 1)
    g = J * d
    return g, (n1 + n2) / (n1 * n2) + g ** 2 / (2 * (n1 + n2))


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
    qual = {r["uid"]: r["quality_tier"] for r in
            csv.DictReader(open(os.path.join(FT, "quality_v2.csv"), encoding="utf-8-sig"))}
    new = {r["uid"]: r for r in rd("ma_v2_new_inputs.csv")}

    # ── 풀 구성 (ma_v2.py와 동일 정의) ──────────────────────────────
    walk = load_walk(rd)
    stay = [dict(uid=r["no"], label=r["label"], g=float(r["g"]), v=float(r["v"]))
            for r in rd("ma_staying_input.csv")]
    soc = [dict(uid=r["no"], label=r["label"], g=float(r["g"]), v=float(r["v"]))
           for r in rd("ma_social_input.csv")]
    soc.append(dict(uid="CT0025", label=new["CT0025"]["label"],
                    g=float(new["CT0025"]["g"]), v=float(new["CT0025"]["v"])))
    # ★ 종전 코드는 여기서 모든 행의 src 를 "Pearson/latent" 로 **덮어썼다**.
    #   그 뒤 `"Spearman" not in src` 로 거르니 ④ rho 제외 민감도가 아무것도 제외하지
    #   못했고, "주분석과 동일"이라는 보고는 버그 산물이었다(2026-08-06 적발).
    #   입력표의 src 를 그대로 보존한다.
    corr = load_corr(rd, r_to_z)
    for u in ("CT0126", "CT0184"):
        corr.append(dict(uid=u, label=new[u]["label"], g=float(new[u]["g"]),
                         v=float(new[u]["v"]), src="Pearson"))
    for lst in (walk, stay, soc, corr):
        for r in lst:
            r["tier"] = qual.get(r["uid"], "?")
            r["route"] = "citation-tracking" if r["uid"].startswith("CT") else "db-search"

    CL = [("MA1 보행속도", walk, False), ("MA2 체류", stay, False),
          ("MA3 사회적 상호작용", soc, False), ("MA4 지각-행태 상관", corr, True)]

    out, L = [], []
    raw = []   # 반올림 전 값 — 원고 Table 3 은 이것을 읽는다(CSV 4자리 → 2·3자리 이중 반올림 방지)

    def rec(cluster, name, rows, note, br=False):
        o = pool(rows, br)
        raw.append(dict(cluster=cluster, analysis=name, pooled=bool(o and o["pooled"]),
                        studies=[r["uid"] for r in rows],
                        **({k: o[k] for k in ("k", "est", "lo", "hi", "p", "I2", "r", "r_lo", "r_hi")
                            if k in o} if o else {"k": 0})))
        out.append(dict(cluster=cluster, analysis=name, k=(o["k"] if o else 0),
                        est=(round(o["est"], 4) if o else ""),
                        lo=(round(o["lo"], 4) if o else ""), hi=(round(o["hi"], 4) if o else ""),
                        p=(round(o["p"], 4) if o else ""), I2=(round(o["I2"], 1) if o else ""),
                        r_back=(round(o["r"], 4) if o and "r" in o else ""),
                        pooled=("yes" if o and o["pooled"] else "no"),
                        studies="; ".join(r["uid"] for r in rows), note=note))
        return o

    def line(tag, o, br=False):
        if o is None:
            return f"| {tag} | 0 | — | — | — | — | 풀링 불가 |\n"
        s = (f"| {tag} | {o['k']} | {o['est']:+.3f} | {o['lo']:+.3f} ~ {o['hi']:+.3f} | "
             f"{o['p']:.3f} | {o['I2']:.1f}% |")
        s += " 단일연구 |" if not o["pooled"] else (f" r={o['r']:+.3f} |" if br and "r" in o else " |")
        return s + "\n"

    n_inc = sum(1 for r in csv.DictReader(open(os.path.join(FT, "corpus_v4_verdicts.csv"),
                                               encoding="utf-8-sig"))
                if r["final_verdict"] == "FINAL_INCLUDE")
    L.append(f"# Paper32 — 민감도 분석 v2 ({n_inc}편 코퍼스)\n")
    L.append("\n규약 `analysis_rules.md §6` + 갈래 축 추가. 품질 등급은 `quality_v2.csv`(전문 재평가 정본).\n")

    for name, rows, br in CL:
        L.append(f"\n---\n\n## {name}\n\n")
        L.append("기여: " + " · ".join(f"{r['uid']}({r['tier']}"
                                      + (", 인용추적" if r["route"] != "db-search" else "") + ")"
                                      for r in rows) + "\n\n")
        L.append("| 분석 | k | 추정치 | 95% CI | p | I² | 비고 |\n|---|---|---|---|---|---|---|\n")
        base = rec(name, "주분석", rows, "primary", br)
        L.append(line("**주분석**", base, br))

        if len(rows) > 2:
            for i in range(len(rows)):
                sub = rows[:i] + rows[i + 1:]
                o = rec(name, f"LOO −{rows[i]['uid']}", sub, rows[i]["label"], br)
                L.append(line(f"① LOO −{rows[i]['uid']} ({rows[i]['tier']})", o, br))

        hi = [r for r in rows if r["tier"] in ("high", "moderate")]
        L.append(line("② low 제외", rec(name, "저품질 제외", hi,
                                       f"제외 {[r['uid'] for r in rows if r['tier']=='low']}", br), br))
        L.append(line("② high만", rec(name, "high만", [r for r in rows if r["tier"] == "high"], "", br), br))

        db = [r for r in rows if r["route"] == "db-search"]
        if len(db) != len(rows):
            L.append(line("⑦ 인용추적 제외(구 코퍼스)",
                          rec(name, "인용추적 제외", db, "갈래 축", br), br))

    # ── ③ 관측 n · ④ rho · ⑤ 가정 · ⑥ 풀 정의 ────────────────────
    L.append("\n---\n\n## ③④⑤⑥ 그 밖의 축\n\n")
    L.append("| 분석 | k | 추정치 | 95% CI | p | I² | 비고 |\n|---|---|---|---|---|---|---|\n")
    g617o, v617o = hedges_g(0.84, 0.13, 189, 0.81, 0.13, 108)
    w3 = [dict(r) for r in walk]
    for r in w3:
        if r["uid"] == "617":
            r["g"], r["v"] = g617o, v617o
    L.append(line("③ MA1 617 관측 n(189/108)", rec("MA1 보행속도", "③ 관측 n", w3, "유사반복 가정 반전")))

    pear = [r for r in corr if "Spearman" not in r.get("src", "")]
    L.append(line("④ MA4 rho 제외", rec("MA4 지각-행태 상관", "④ rho 제외", pear, "규약 §1", True), True))

    L.append(line("⑤ MA1 −532(균등분할 가정)",
                  rec("MA1 보행속도", "⑤ 532 제외", [r for r in walk if r["uid"] != "532"], "군당 n 미보고")))
    L.append(line("⑤ MA2 −665(p→t 역산)",
                  rec("MA2 체류", "⑤ 665 제외", [r for r in stay if r["uid"] != "665"], "SD 미보고")))

    # ⚠️ 2026-09-16 정정: 종전 코드는 원자료를 직접 읽어 532 의 Exp1·Exp2 를 따로 넣었다(k=5).
    #    D7-2 에서 주분석만 고치고 이 축은 같은 규약 위반이 남아 있었다(독립 검토 중 적발).
    #    주분석과 같은 로더로 논문 내 합성을 적용한다 → k=4.
    walk_all = load_walk(rd, exclude_music=False)
    for r in walk_all:
        r["tier"] = qual.get(r["uid"], "?"); r["route"] = "db-search"
    L.append(line("⑥ MA1 +481(음악 포함)", rec("MA1 보행속도", "⑥ 481 포함", walk_all, "노출·행태 이질")))

    with open(os.path.join(MA, "ma_sensitivity_v2.csv"), "w", newline="",
              encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["cluster", "analysis", "k", "est", "lo", "hi", "p",
                                          "I2", "r_back", "pooled", "studies", "note"])
        w.writeheader(); w.writerows(out)
    import json
    with open(os.path.join(MA, "ma_sensitivity_v2_raw.json"), "w", encoding="utf-8") as f:
        json.dump(raw, f, ensure_ascii=False, indent=1)

    # ── 핵심 판정 ──────────────────────────────────────────────────
    def get(cl, an):
        for r in out:
            if r["cluster"] == cl and r["analysis"] == an:
                return r
        return None

    L.append("\n---\n\n## 판정\n\n")
    m1 = get("MA1 보행속도", "주분석")
    m1low = get("MA1 보행속도", "저품질 제외")
    m3 = get("MA3 사회적 상호작용", "주분석")
    m3db = get("MA3 사회적 상호작용", "인용추적 제외")
    m4 = get("MA4 지각-행태 상관", "주분석")
    m4db = get("MA4 지각-행태 상관", "인용추적 제외")
    L.append(f"1. **MA1은 저품질을 빼면 k={m1low['k']}** — 기여 {m1['k']}효과가 전부 MMAT low라 "
             "풀링이 성립하지 않는다. 이 리뷰의 가장 큰 한계이며 은폐하지 않고 보고한다.\n")
    if m3 and m3db:
        L.append(f"2. **MA3의 유의성은 인용추적에 의존한다** — 인용추적분(CT0025)을 빼면 "
                 f"k {m3['k']}→{m3db['k']}, p {m3['p']}→{m3db['p']}로 유의성이 사라진다. "
                 "즉 등록된 보조 검색 경로가 결론을 바꿨다.\n")
    if m4 and m4db:
        L.append(f"3. **MA4는 갈래에 무관하게 안정적** — 인용추적분을 빼도 "
                 f"r {m4['r_back']}→{m4db['r_back']}, p {m4['p']}→{m4db['p']}.\n")
    L.append("\n⚠️ k = 2 행의 신뢰구간은 해석하지 말 것(t(1) = 12.7).\n")

    open(os.path.join(MA, "ma_sensitivity_v2.md"), "w", encoding="utf-8").write("".join(L))
    print("".join(L[-6:]))
    print(f"[저장] ma/ma_sensitivity_v2.csv ({len(out)}행) · ma_sensitivity_v2.md")


if __name__ == "__main__":
    main()
