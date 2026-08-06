# -*- coding: utf-8 -*-
"""
Paper32 — 민감도 분석 v2 (96편 코퍼스 · MA3 k=4 · MA4 k=7 반영)
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


def reml_tau2(y, v, iters=500):
    t2 = max(float(np.var(y, ddof=1) - np.mean(v)), 0.0) if len(y) > 1 else 0.0
    for _ in range(iters):
        w = 1 / (v + t2)
        mu = np.sum(w * y) / np.sum(w)
        new = max(float(np.sum(w ** 2 * ((y - mu) ** 2 - v)) / np.sum(w ** 2)), 0.0)
        if abs(new - t2) < 1e-12:
            break
        t2 = new
    return t2


def pool(rows, back_r=False):
    if not rows:
        return None
    y = np.array([r["g"] for r in rows], float)
    v = np.array([r["v"] for r in rows], float)
    k = len(y)
    if k == 1:
        est, se = float(y[0]), math.sqrt(float(v[0]))
        o = dict(k=1, est=est, lo=est - 1.96 * se, hi=est + 1.96 * se,
                 p=float(2 * (1 - stats.norm.cdf(abs(est / se)))), I2=0.0, tau2=0.0, pooled=False)
    else:
        t2 = reml_tau2(y, v); w = 1 / (v + t2)
        mu = float(np.sum(w * y) / np.sum(w)); se = math.sqrt(1 / np.sum(w))
        qhk = float(np.sum(w * (y - mu) ** 2) / (k - 1)); se_hk = se * math.sqrt(qhk)
        tc = stats.t.ppf(0.975, k - 1)
        wf = 1 / v; muf = float(np.sum(wf * y) / np.sum(wf))
        Q = float(np.sum(wf * (y - muf) ** 2))
        o = dict(k=k, est=mu, lo=mu - tc * se_hk, hi=mu + tc * se_hk,
                 p=float(2 * (1 - stats.t.cdf(abs(mu / se_hk), k - 1))), tau2=t2,
                 I2=max(0.0, (Q - (k - 1)) / Q) * 100 if Q > 0 else 0.0, pooled=True)
    if back_r:
        o["r"] = math.tanh(o["est"])
    return o


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


def main():
    qual = {r["uid"]: r["quality_tier"] for r in
            csv.DictReader(open(os.path.join(FT, "quality_v2.csv"), encoding="utf-8-sig"))}
    new = {r["uid"]: r for r in rd("ma_v2_new_inputs.csv")}

    # ── 풀 구성 (ma_v2.py와 동일 정의) ──────────────────────────────
    walk = [dict(uid=r["no"], label=r["label"], g=float(r["g"]), v=float(r["var"]),
                 contrast=r["contrast"])
            for r in rd("ma_walking_input.csv")
            if "control" not in r["contrast"] and r["no"] != "481"]
    stay = [dict(uid=r["no"], label=r["label"], g=float(r["g"]), v=float(r["v"]))
            for r in rd("ma_staying_input.csv")]
    soc = [dict(uid=r["no"], label=r["label"], g=float(r["g"]), v=float(r["v"]))
           for r in rd("ma_social_input.csv")]
    soc.append(dict(uid="CT0025", label=new["CT0025"]["label"],
                    g=float(new["CT0025"]["g"]), v=float(new["CT0025"]["v"])))
    corr = []
    for r in rd("ma_correlation_input.csv"):
        z, v = r_to_z(float(r["r"]), int(r["n"]))
        corr.append(dict(uid=r["no"], label=r["label"], g=z, v=v, r=float(r["r"]),
                         n=int(r["n"]), src="Pearson/latent"))
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

    def rec(cluster, name, rows, note, br=False):
        o = pool(rows, br)
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

    L.append("# Paper32 — 민감도 분석 v2 (96편 코퍼스)\n")
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

    walk_all = [dict(uid=r["no"], label=r["label"], g=float(r["g"]), v=float(r["var"]))
                for r in rd("ma_walking_input.csv") if "control" not in r["contrast"]]
    for r in walk_all:
        r["tier"] = qual.get(r["uid"], "?"); r["route"] = "db-search"
    L.append(line("⑥ MA1 +481(음악 포함)", rec("MA1 보행속도", "⑥ 481 포함", walk_all, "노출·행태 이질")))

    with open(os.path.join(MA, "ma_sensitivity_v2.csv"), "w", newline="",
              encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["cluster", "analysis", "k", "est", "lo", "hi", "p",
                                          "I2", "r_back", "pooled", "studies", "note"])
        w.writeheader(); w.writerows(out)

    # ── 핵심 판정 ──────────────────────────────────────────────────
    def get(cl, an):
        for r in out:
            if r["cluster"] == cl and r["analysis"] == an:
                return r
        return None

    L.append("\n---\n\n## 판정\n\n")
    m1low = get("MA1 보행속도", "저품질 제외")
    m3 = get("MA3 사회적 상호작용", "주분석")
    m3db = get("MA3 사회적 상호작용", "인용추적 제외")
    m4 = get("MA4 지각-행태 상관", "주분석")
    m4db = get("MA4 지각-행태 상관", "인용추적 제외")
    L.append(f"1. **MA1은 저품질을 빼면 k={m1low['k']}** — 기여 4효과가 전부 MMAT low라 "
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
