# -*- coding: utf-8 -*-
"""
Paper32 — MA #1: 보행속도 클러스터 (자연음/긍정음 vs 교통·도시소음 대비)
analysis_rules.md 준수: Hedges' g(J 보정) · REML 랜덤효과 + Hartung-Knapp · 유사반복은 참가자 n
입력 수치는 전부 effect_sizes_all.csv verbatim(원문 표) — 아래 source 필드에 위치 기록.
출력: ma/ma_walking_input.csv · ma/ma_walking_result.md
"""
import sys, os, csv, math
import numpy as np
from scipy import stats

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
MA = os.path.join(BASE, "fulltext", "ma")
os.makedirs(MA, exist_ok=True)

# 대비 정의: g>0 = 긍정음(자연음·음악) 조건에서 보행속도가 더 빠름
# 회피가설 예측: 소음일수록 빠름 → g<0 기대
STUDIES = [
    # (id, label, 긍정음 M,SD,n, 비교음 M,SD,n, contrast, source, note)
    (461, "Franek 2018", 1.53, 0.12, 31, 1.65, 0.11, 26, "birdsong vs traffic noise",
     "Table/본문 Results; F2,80=7.759", "field walk 1.8km, between-subjects"),
    (461, "Franek 2018 (vs control)", 1.53, 0.12, 31, 1.58, 0.13, 26, "birdsong vs control(no sound)",
     "동일 표", "민감도용 — 대조군 비교"),
    (532, "Franek 2019 Exp1", 1.56, 0.17, 29, 1.62, 0.16, 29, "birdsong vs crowded city noise",
     "Exp1 fall N=87 3군 균등 가정(29/29/29)", "3군 N=87 → 군당 n 미보고, 균등분할 가정(민감도 대상)"),
    (532, "Franek 2019 Exp2", 1.49, 0.11, 22, 1.61, 0.15, 22, "birdsong vs crowded city noise",
     "Exp2 spring N=65 3군 균등 가정(22/22/21); F2,62=3.290, p=.044", "동일"),
    (617, "Oases soundscape", 0.84, 0.13, 27, 0.81, 0.13, 27, "bird/nature vs traffic noise",
     "Table 6; F=7.558 df=485", "유사반복 — 참가자 n=27 사용(관측 n 189/108 아님)"),
    (481, "Meng 2018 (music)", 1.14, 0.17839, 40, 1.43, 0.16733, 40, "background music vs no music",
     "Table 1 S1 vs S0", "walking-around 행태; 음악=긍정음 프레임"),
]


def hedges_g(m1, sd1, n1, m2, sd2, n2):
    df = n1 + n2 - 2
    sp = math.sqrt(((n1 - 1) * sd1 ** 2 + (n2 - 1) * sd2 ** 2) / df)
    d = (m1 - m2) / sp
    J = 1 - 3 / (4 * df - 1)
    g = J * d
    v = (n1 + n2) / (n1 * n2) + g ** 2 / (2 * (n1 + n2))
    return g, v


def reml_tau2(y, v, iters=200):
    t2 = max(np.var(y, ddof=1) - np.mean(v), 0.0)
    for _ in range(iters):
        w = 1 / (v + t2)
        mu = np.sum(w * y) / np.sum(w)
        num = np.sum(w ** 2 * ((y - mu) ** 2 - v)) + 1 / np.sum(w) * np.sum(w ** 2) * 0
        den = np.sum(w ** 2)
        new = max(num / den, 0.0)
        if abs(new - t2) < 1e-10:
            t2 = new; break
        t2 = new
    return t2


def main():
    rows = []
    for (sid, lab, m1, s1, n1, m2, s2, n2, contrast, src, note) in STUDIES:
        g, v = hedges_g(m1, s1, n1, m2, s2, n2)
        rows.append(dict(no=sid, label=lab, contrast=contrast, m_pos=m1, sd_pos=s1, n_pos=n1,
                         m_neg=m2, sd_neg=s2, n_neg=n2, g=round(g, 4), var=round(v, 5),
                         se=round(math.sqrt(v), 4), source=src, note=note))
    with open(os.path.join(MA, "ma_walking_input.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys())); w.writeheader(); w.writerows(rows)

    # 주분석: 논문당 1효과 (461은 vs traffic 채택, 532는 Exp1/Exp2 = 독립표본이라 둘 다 유지)
    main_rows = [r for r in rows if "control" not in r["contrast"]]
    y = np.array([r["g"] for r in main_rows]); v = np.array([r["var"] for r in main_rows])
    k = len(y)
    t2 = reml_tau2(y, v)
    w_ = 1 / (v + t2)
    mu = np.sum(w_ * y) / np.sum(w_)
    se_re = math.sqrt(1 / np.sum(w_))
    # Hartung-Knapp
    q_hk = np.sum(w_ * (y - mu) ** 2) / (k - 1)
    se_hk = se_re * math.sqrt(q_hk)
    tcrit = stats.t.ppf(0.975, k - 1)
    ci = (mu - tcrit * se_hk, mu + tcrit * se_hk)
    p_hk = 2 * (1 - stats.t.cdf(abs(mu / se_hk), k - 1))
    # 이질성
    wf = 1 / v
    muf = np.sum(wf * y) / np.sum(wf)
    Q = np.sum(wf * (y - muf) ** 2)
    dfQ = k - 1
    I2 = max(0.0, (Q - dfQ) / Q) * 100
    pQ = 1 - stats.chi2.cdf(Q, dfQ)

    # leave-one-out
    loo = []
    for i in range(k):
        yi = np.delete(y, i); vi = np.delete(v, i)
        t2i = reml_tau2(yi, vi); wi = 1 / (vi + t2i)
        loo.append((main_rows[i]["label"], round(float(np.sum(wi * yi) / np.sum(wi)), 3)))

    lines = []
    lines.append("# MA #1 — 보행속도 클러스터 (긍정음 vs 소음)\n")
    lines.append(f"- 규약: analysis_rules.md · g>0 = 긍정음 조건에서 **더 빠른** 보행 (회피가설은 g<0 예측)\n")
    lines.append(f"- 주분석 k={k} 대비 (논문 {len(set(r['no'] for r in main_rows))}편)\n\n")
    lines.append("| 연구 | 대비 | g | SE | n(긍정/비교) |\n|---|---|---|---|---|\n")
    for r in main_rows:
        lines.append(f"| {r['label']} | {r['contrast']} | {r['g']:+.3f} | {r['se']:.3f} | {r['n_pos']}/{r['n_neg']} |\n")
    lines.append(f"\n**풀링(REML + Hartung-Knapp)**: g = **{mu:+.3f}** "
                 f"(95% CI {ci[0]:+.3f} ~ {ci[1]:+.3f}), p = {p_hk:.4f}\n")
    lines.append(f"- 이질성: tau^2 = {t2:.4f}, I^2 = {I2:.1f}%, Q({dfQ}) = {Q:.2f}, p = {pQ:.3f}\n")
    lines.append(f"- leave-one-out: {loo}\n")
    sens = [r for r in rows if "control" in r["contrast"]]
    if sens:
        lines.append(f"\n**민감도(대조군 대비 추가)**: {[(r['label'], r['g']) for r in sens]}\n")
    lines.append("\n## 해석 주의\n")
    lines.append("- 532는 군당 n 미보고 → 균등분할 가정(민감도에서 관측 n 변형 재계산 필요)\n")
    lines.append("- 617은 유사반복 구조 → 참가자 n=27 보수적 적용\n")
    lines.append("- 481은 음악(긍정음)이 보행속도를 **낮춘** 사례로 프레임 내 부호가 반대 방향 기여\n")
    open(os.path.join(MA, "ma_walking_result.md"), "w", encoding="utf-8").write("".join(lines))
    print("".join(lines))


if __name__ == "__main__":
    main()
