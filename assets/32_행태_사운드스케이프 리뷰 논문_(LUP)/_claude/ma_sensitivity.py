# -*- coding: utf-8 -*-
"""
Paper32 — 등록 프로토콜 민감도 분석 5종 (analysis_rules.md §6)
  ① leave-one-out
  ② MMAT 저품질(low) 제외        ← quality_all.csv 필요(품질평가 완료로 이제 실행 가능)
  ③ 관측 n 기반 재계산(유사반복 가정 반전)
  ④ rho(Spearman) 제외
  ⑤ 가정 의존 입력 제외(균등분할 가정·p→t 역산)
입력은 전부 ma/*_input.csv (effect_sizes_all.csv verbatim에서 유래).
출력: fulltext/ma/ma_sensitivity.csv · ma_sensitivity.md
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


def pool(y, v, back_r=False):
    """REML + Hartung-Knapp. k<2면 단일 연구 CI(정규근사)를 반환하되 pooled 아님을 표시."""
    y = np.asarray(y, float); v = np.asarray(v, float); k = len(y)
    if k == 0:
        return None
    if k == 1:
        est, se = float(y[0]), math.sqrt(float(v[0]))
        o = dict(k=1, est=est, lo=est - 1.96 * se, hi=est + 1.96 * se,
                 p=float(2 * (1 - stats.norm.cdf(abs(est / se)))), tau2=0.0, I2=0.0, pooled=False)
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
                 p=float(2 * (1 - stats.t.cdf(abs(mu / se_hk), k - 1))),
                 tau2=t2, I2=max(0.0, (Q - (k - 1)) / Q) * 100 if Q > 0 else 0.0, pooled=True)
    if back_r:
        o["r"] = math.tanh(o["est"]); o["r_lo"] = math.tanh(o["lo"]); o["r_hi"] = math.tanh(o["hi"])
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
    qual = {int(r["no"]): r["quality_tier"] for r in
            csv.DictReader(open(os.path.join(FT, "quality_all.csv"), encoding="utf-8-sig"))}

    # ── 클러스터별 주분석 입력 재구성 ────────────────────────────────
    # MA1 주분석 정의(ma_summary.md 정본): 대조군 대비 행 제외 + 481(음악→배회행태)은
    # 노출·행태 이질로 주분석에서 제외 → k=4. 481 포함형은 아래 ⑥에서 민감도로 제시.
    walk_all = [dict(no=int(r["no"]), label=r["label"], g=float(r["g"]), v=float(r["var"]),
                     contrast=r["contrast"])
                for r in rd("ma_walking_input.csv") if "control" not in r["contrast"]]
    walk = [r for r in walk_all if r["no"] != 481]
    stay = [dict(no=int(r["no"]), label=r["label"], g=float(r["g"]), v=float(r["v"]))
            for r in rd("ma_staying_input.csv")]
    soc = [dict(no=int(r["no"]), label=r["label"], g=float(r["g"]), v=float(r["v"]))
           for r in rd("ma_social_input.csv")]
    corr = []
    for r in rd("ma_correlation_input.csv"):
        z, v = r_to_z(float(r["r"]), int(r["n"]))
        corr.append(dict(no=int(r["no"]), label=r["label"], g=z, v=v,
                         r=float(r["r"]), n=int(r["n"]), src=r["src"]))
    for lst in (walk_all, walk, stay, soc, corr):
        for r in lst:
            r["tier"] = qual.get(r["no"], "?")

    CL = [("MA1 보행속도", walk, False), ("MA2 체류", stay, False),
          ("MA3 사회적 상호작용", soc, False), ("MA4 지각-행태 상관", corr, True)]

    out_rows, L = [], []

    def rec(cluster, sens, rows, note, back_r=False):
        o = pool([r["g"] for r in rows], [r["v"] for r in rows], back_r=back_r) if rows else None
        d = dict(cluster=cluster, analysis=sens, k=(o["k"] if o else 0),
                 est=(round(o["est"], 4) if o else ""), lo=(round(o["lo"], 4) if o else ""),
                 hi=(round(o["hi"], 4) if o else ""), p=(round(o["p"], 4) if o else ""),
                 I2=(round(o["I2"], 1) if o else ""), tau2=(round(o["tau2"], 4) if o else ""),
                 pooled=("yes" if o and o["pooled"] else "no"),
                 studies="; ".join(str(r["no"]) for r in rows), note=note)
        if o and back_r:
            d["r_back"] = round(o["r"], 4)
        out_rows.append(d)
        return o

    def line(o, tag, back_r=False):
        if o is None:
            return f"| {tag} | 0 | — | — | — | — | 풀링 불가 |"
        s = (f"| {tag} | {o['k']} | {o['est']:+.3f} | {o['lo']:+.3f} ~ {o['hi']:+.3f} | "
             f"{o['p']:.3f} | {o['I2']:.1f}% |")
        s += (" 단일연구(풀링 아님) |" if not o["pooled"] else " |")
        if back_r and "r" in o:
            s = s.rstrip("|").rstrip() + f" r={o['r']:+.3f} |"
        return s

    L.append("# Paper32 — 민감도 분석 (등록 프로토콜 §Synthesis plan 이행)\n")
    L.append("\n작성 2026-08-03 · 규약 `analysis_rules.md §6`. 주분석 결과의 강건성을 5개 축으로 점검.\n")
    L.append("품질 등급은 MMAT 2018 판정(`quality_all.csv`) — high(4~5) / moderate(3) / low(0~2).\n")

    for name, rows, back_r in CL:
        L.append(f"\n---\n\n## {name}\n\n")
        L.append("기여 연구 품질: " + " · ".join(f"{r['no']}({r['tier']})" for r in rows) + "\n\n")
        L.append("| 분석 | k | 추정치 | 95% CI | p | I² | 비고 |\n|---|---|---|---|---|---|---|\n")
        base = rec(name, "주분석", rows, "primary", back_r)
        L.append(line(base, "**주분석**", back_r) + "\n")

        # ① leave-one-out
        if len(rows) > 2:
            for i in range(len(rows)):
                sub = rows[:i] + rows[i + 1:]
                o = rec(name, f"LOO 제외 {rows[i]['no']}", sub, rows[i]["label"], back_r)
                L.append(line(o, f"① LOO −{rows[i]['no']} ({rows[i]['tier']})", back_r) + "\n")

        # ② 저품질 제외
        hi = [r for r in rows if r["tier"] in ("high", "moderate")]
        o = rec(name, "저품질(low) 제외", hi,
                f"제외 {[r['no'] for r in rows if r['tier']=='low']}", back_r)
        L.append(line(o, "② low 제외", back_r) + "\n")
        hionly = [r for r in rows if r["tier"] == "high"]
        o = rec(name, "high만", hionly, "", back_r)
        L.append(line(o, "② high만", back_r) + "\n")

    # ── ③ 관측 n 기반 재계산 (MA1 617 유사반복 가정 반전) ──────────────
    L.append("\n---\n\n## ③ 관측 n 기반 재계산 (유사반복 가정 반전)\n\n")
    L.append("규약 §4는 참가자 n을 채택한다. 이 가정을 뒤집어 원문의 관측 단위 n을 쓰면 어떻게 되는지 확인한다.\n\n")
    g617o, v617o = hedges_g(0.84, 0.13, 189, 0.81, 0.13, 108)   # 관측 n(구간 반복)
    walk3 = [dict(r) for r in walk]
    for r in walk3:
        if r["no"] == 617:
            r["g"], r["v"] = g617o, v617o
    o = rec("MA1 보행속도", "③ 617 관측 n(189/108)", walk3, "참가자 n=27 → 관측 n")
    L.append("| 분석 | k | 추정치 | 95% CI | p | I² | 비고 |\n|---|---|---|---|---|---|---|\n")
    L.append(line(base if False else o, "③ MA1 617 관측 n") + "\n")
    L.append(f"\n- 617 개별 g: 참가자 n {walk[[r['no'] for r in walk].index(617)]['g']:+.3f} "
             f"(v={walk[[r['no'] for r in walk].index(617)]['v']:.4f}) → 관측 n {g617o:+.3f} (v={v617o:.4f})\n")
    L.append("- 방향은 불변, 분산만 축소되어 가중치가 커진다. 즉 유사반복을 인정하면 617의 영향력이 과대평가된다.\n")
    L.append("- MA3의 931·1069는 발췌(excerpt)·그룹 단위 n을 쓰므로 같은 위험을 공유하나, 참가자 n이 원문에\n"
             "  보고되지 않아 재계산 불가 — 서술로만 표기한다.\n")

    # ── ④ rho 제외 (MA4) ────────────────────────────────────────────
    L.append("\n---\n\n## ④ Spearman rho 제외 (MA4)\n\n")
    L.append("규약 §1은 rho를 r 근사로 취급하되 민감도에서 분리 검토하도록 한다.\n\n")
    pear = [r for r in corr if "Spearman" not in r["src"]]
    o4 = rec("MA4 지각-행태 상관", "④ rho 제외", pear,
             f"제외 {[r['no'] for r in corr if 'Spearman' in r['src']]}", True)
    L.append("| 분석 | k | 추정치 | 95% CI | p | I² | 비고 |\n|---|---|---|---|---|---|---|\n")
    L.append(line(o4, "④ Pearson·잠재변수만", True) + "\n")

    # ── ⑤ 가정 의존 입력 제외 ────────────────────────────────────────
    L.append("\n---\n\n## ⑤ 가정 의존 입력 제외\n\n")
    L.append("원문이 직접 보고하지 않아 **내가 가정을 넣어 복원한** 입력을 뺀다.\n")
    L.append("- MA1 532(Franek 2019 Exp1·Exp2): 군당 n 미보고 → 총 N 균등분할 가정\n")
    L.append("- MA2 665(Ba & Kang 2020): SD 미보고 → 사후 p=0.003에서 t 역산\n\n")
    w5 = [r for r in walk if r["no"] != 532]
    o5a = rec("MA1 보행속도", "⑤ 균등분할 가정(532) 제외", w5, "군당 n 미보고")
    s5 = [r for r in stay if r["no"] != 665]
    o5b = rec("MA2 체류", "⑤ p→t 역산(665) 제외", s5, "SD 미보고")
    L.append("| 분석 | k | 추정치 | 95% CI | p | I² | 비고 |\n|---|---|---|---|---|---|---|\n")
    L.append(line(o5a, "⑤ MA1 −532") + "\n")
    L.append(line(o5b, "⑤ MA2 −665") + "\n")

    # ── ⑥ MA1 풀 정의 대안 ───────────────────────────────────────────
    L.append("\n---\n\n## ⑥ MA1 풀 정의 대안 — 음악 자극 연구(481) 포함\n\n")
    L.append("주분석은 '자연음 vs 교통·도시소음' 대비만 묶는다. 481(Meng 2018)은 음악 vs 무음악이고\n"
             "행태도 통과보행이 아니라 배회(walking-around)여서 노출·행태가 모두 이질적이라 제외했다.\n"
             "이 결정이 결과를 만드는지 확인한다.\n\n")
    o6 = rec("MA1 보행속도", "⑥ 481(음악) 포함", walk_all, "노출·행태 이질 결합")
    L.append("| 분석 | k | 추정치 | 95% CI | p | I² | 비고 |\n|---|---|---|---|---|---|---|\n")
    L.append(line(o6, "⑥ MA1 +481") + "\n")
    L.append("\n- 481은 g = −1.661로 절대값이 가장 크고 방향은 주분석과 같다. 포함하면 추정치가 커지고\n"
             "  CI가 0에서 더 멀어지므로, **제외 결정은 보수적인 쪽**이다(유리하게 고른 것이 아니다).\n")

    # ── 저장 ────────────────────────────────────────────────────────
    keys = ["cluster", "analysis", "k", "est", "lo", "hi", "p", "I2", "tau2", "pooled",
            "r_back", "studies", "note"]
    with open(os.path.join(MA, "ma_sensitivity.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=keys, extrasaction="ignore")
        w.writeheader(); w.writerows(out_rows)
    open(os.path.join(MA, "ma_sensitivity.md"), "w", encoding="utf-8").write("".join(L))
    print("".join(L))
    print(f"\n[저장] ma_sensitivity.csv ({len(out_rows)}행) · ma_sensitivity.md")


if __name__ == "__main__":
    main()
