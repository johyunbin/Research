# -*- coding: utf-8 -*-
"""
Paper32 — MA #2 체류 · #3 사회적 상호작용 · #4 역방향(r 계열)
analysis_rules.md 준수. 모든 입력은 effect_sizes_all.csv verbatim에서 유래(source 필드 기록).
출력: fulltext/ma/ma_staying_*.csv|md · ma_social_*.csv|md · ma_reverse_*.md
"""
import sys, os, csv, math
import numpy as np
from scipy import stats

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
MA = os.path.join(BASE, "fulltext", "ma")
os.makedirs(MA, exist_ok=True)


def hedges_g(m1, sd1, n1, m2, sd2, n2):
    df = n1 + n2 - 2
    sp = math.sqrt(((n1 - 1) * sd1 ** 2 + (n2 - 1) * sd2 ** 2) / df)
    d = (m1 - m2) / sp
    J = 1 - 3 / (4 * df - 1)
    g = J * d
    v = (n1 + n2) / (n1 * n2) + g ** 2 / (2 * (n1 + n2))
    return g, v


def or_to_g(lnOR, se_lnOR):
    """로그오즈 → Cohen d (Chinn) → Hedges g 근사(대표본이라 J≈1)"""
    k = math.sqrt(3) / math.pi
    return lnOR * k, (se_lnOR * k) ** 2


def r_to_z(r, n):
    z = 0.5 * math.log((1 + r) / (1 - r))
    return z, 1 / (n - 3)


def reml_tau2(y, v, iters=500):
    t2 = max(float(np.var(y, ddof=1) - np.mean(v)), 0.0)
    for _ in range(iters):
        w = 1 / (v + t2)
        mu = np.sum(w * y) / np.sum(w)
        new = max(float(np.sum(w ** 2 * ((y - mu) ** 2 - v)) / np.sum(w ** 2)), 0.0)
        if abs(new - t2) < 1e-12:
            break
        t2 = new
    return t2


def pool(y, v, label, back_r=False):
    y = np.asarray(y, float); v = np.asarray(v, float); k = len(y)
    t2 = reml_tau2(y, v)
    w = 1 / (v + t2)
    mu = float(np.sum(w * y) / np.sum(w))
    se = math.sqrt(1 / np.sum(w))
    if k > 1:
        qhk = float(np.sum(w * (y - mu) ** 2) / (k - 1))
        se_hk = se * math.sqrt(qhk)
        tc = stats.t.ppf(0.975, k - 1)
        lo, hi = mu - tc * se_hk, mu + tc * se_hk
        p = float(2 * (1 - stats.t.cdf(abs(mu / se_hk), k - 1)))
    else:
        lo, hi, p = mu - 1.96 * se, mu + 1.96 * se, float("nan")
    wf = 1 / v
    muf = float(np.sum(wf * y) / np.sum(wf))
    Q = float(np.sum(wf * (y - muf) ** 2))
    I2 = max(0.0, (Q - (k - 1)) / Q) * 100 if k > 1 and Q > 0 else 0.0
    out = dict(label=label, k=k, est=mu, lo=lo, hi=hi, p=p, tau2=t2, I2=I2, Q=Q)
    if back_r:
        out["r"] = math.tanh(mu); out["r_lo"] = math.tanh(lo); out["r_hi"] = math.tanh(hi)
    return out


def fmt(o, unit="g"):
    s = f"| {o['label']} | {o['k']} | {o['est']:+.3f} | {o['lo']:+.3f} ~ {o['hi']:+.3f} | {o['p']:.3f} | {o['tau2']:.3f} | {o['I2']:.1f}% |"
    if "r" in o:
        s += f"\n(역변환 r = {o['r']:+.3f} [{o['r_lo']:+.3f}, {o['r_hi']:+.3f}])"
    return s


def main():
    report = []

    # ================= MA #2 체류(dwell) — 긍정음 vs 대조/소음, g =================
    # 323 Aletta: log-duration EMM ± SE, n=stoppers. 음악 3조건 통합 vs 무음악 대조
    #   SD = SE*sqrt(n)
    n_cl, n_am, n_jz, n_ct = 181, 161, 135, 119
    sd_cl, sd_am, sd_jz, sd_ct = 0.072*math.sqrt(n_cl), 0.079*math.sqrt(n_am), 0.087*math.sqrt(n_jz), 0.120*math.sqrt(n_ct)
    n_mu = n_cl + n_am + n_jz
    m_mu = (5.29*n_cl + 5.25*n_am + 5.17*n_jz) / n_mu
    sd_mu = math.sqrt(((n_cl-1)*sd_cl**2 + (n_am-1)*sd_am**2 + (n_jz-1)*sd_jz**2) / (n_mu-3))
    g323, v323 = hedges_g(m_mu, sd_mu, n_mu, 4.83, sd_ct, n_ct)

    # 1069 OSI-P(P) 지속적 상호작용 발생(체류 성격 지표는 아님) → 사회로. 여기선 제외.
    # 1280 로지스틱 OR: NSI(자연음지수)↑ → 장기체류 OR=1.475 (B=0.389, SE=0.132)
    g1280, v1280 = or_to_g(0.389, 0.132)

    # 931 PTSSI(G)/체류성 지표 대신 사회로. 체류에는 미사용.
    # 665: 음악 vs 팬소음 EMM 차이(약 70s)이나 SD 미보고 → 제외(서술)
    # 1178: 옴니버스 F(df=2) → 규약상 미변환. 단 Scheffe music vs fan p=0.031 서술.

    # 665: 음악 vs 무음(사후비교 p=0.003, 군 n≈50/47) → p→t→d 역산(규약 §1 test_stat 경로)
    df665 = 50 + 47 - 2
    t665 = stats.t.ppf(1 - 0.003 / 2, df665)
    d665 = t665 * math.sqrt(1 / 50 + 1 / 47)
    J665 = 1 - 3 / (4 * df665 - 1)
    g665 = J665 * d665
    v665 = (50 + 47) / (50 * 47) + g665 ** 2 / (2 * (50 + 47))

    staying_rows = [
        dict(no=323, label="Aletta 2016 (음악 vs 무음악)", g=g323, v=v323,
             note="log-duration EMM±SE→SD 환산, n=477/119", source="Table EMM; F(3,574)=3.781"),
        dict(no=665, label="Ba & Kang 2020 (음악 vs 무음, 체류시간)", g=g665, v=v665,
             note="사후 p=0.003·군n≈50/47에서 t 역산(보수적)", source="EMM +30s, Sound F=26.879"),
        dict(no=1280, label="Bao 2026 (자연음지수↑, 장기체류 OR)", g=g1280, v=v1280,
             note="OR=1.475→Chinn 변환, N=241", source="binary logistic"),
    ]
    o_stay = pool([r["g"] for r in staying_rows], [r["v"] for r in staying_rows], "체류(긍정음→체류 증가)")

    # ================= MA #3 사회적 상호작용 — 자연음 vs 소음/대조, g =================
    # 931: PPSI(G) 그룹 상호작용 비율 — bird/water(자연) vs traffic/construction(소음)
    #   자연 통합: bird 29.64(15.79,n=18) + water 22.95(14.11,n=19)
    n_b, n_w, n_t, n_c, n_ctl = 18, 19, 18, 18, 17
    n_nat = n_b + n_w
    m_nat = (29.64*n_b + 22.95*n_w) / n_nat
    sd_nat = math.sqrt(((n_b-1)*15.79**2 + (n_w-1)*14.11**2) / (n_nat-2))
    n_noise = n_t + n_c
    m_noise = (13.88*n_t + 10.57*n_c) / n_noise
    sd_noise = math.sqrt(((n_t-1)*15.49**2 + (n_c-1)*10.64**2) / (n_noise-2))
    g931, v931 = hedges_g(m_nat, sd_nat, n_nat, m_noise, sd_noise, n_noise)
    # 931 대조 대비(민감도)
    g931c, v931c = hedges_g(m_nat, sd_nat, n_nat, 7.35, 8.69, n_ctl)

    # 1069: SP(P) 쌍 상호작용 참여 — 자연(birdsong 51.21/23.58 n=37, water 43.60/18.85 n=36)
    #   vs 소음(traffic 34.94/16.67 n=39, construction 39.02/17.41 n=34)
    n_b2, n_w2, n_t2, n_c2, n_ctl2 = 37, 36, 39, 34, 36
    n_nat2 = n_b2 + n_w2
    m_nat2 = (51.21*n_b2 + 43.60*n_w2) / n_nat2
    sd_nat2 = math.sqrt(((n_b2-1)*23.58**2 + (n_w2-1)*18.85**2) / (n_nat2-2))
    n_noise2 = n_t2 + n_c2
    m_noise2 = (34.94*n_t2 + 39.02*n_c2) / n_noise2
    sd_noise2 = math.sqrt(((n_t2-1)*16.67**2 + (n_c2-1)*17.41**2) / (n_noise2-2))
    g1069, v1069 = hedges_g(m_nat2, sd_nat2, n_nat2, m_noise2, sd_noise2, n_noise2)

    # 14 Moser 1988: 소음(도로공사 R+) vs 무소음(Ro) — 원조행동(열쇠 줍기/가리키기 = 도움)
    #   셀: Ignoring Ro=8+20=28, R+=13+56=69 / 도움(pointing+picking) Ro=(31+27)+(11+3)=72, R+=(32+43)+(5+1)=81
    a, b_, c_, d_ = 72, 28, 81, 69   # help/ignore for Ro / R+
    orv = (a * d_) / (b_ * c_)
    se_lnor = math.sqrt(1/a + 1/b_ + 1/c_ + 1/d_)
    g14, v14 = or_to_g(math.log(orv), se_lnor)
    # 프레임: 긍정(조용) 조건 vs 소음 조건 · g>0 = 조용할 때 도움행동 더 많음
    # OR=odds(help|Ro)/odds(help|R+)=2.19 → 조용 조건에서 도움 ↑ → g는 양(+)이 프레임 정합

    social_rows = [
        dict(no=931, label="Chen 2023 (자연음 vs 소음, 그룹상호작용%)", g=g931, v=v931,
             note="bird+water vs traffic+construction, 발췌 n=37/36", source="Table 2 PPSI(G)"),
        dict(no=1069, label="Chen 2024 (자연음 vs 소음, 쌍상호작용%)", g=g1069, v=v1069,
             note="excerpt 단위 n=73/73", source="SP(P)"),
        dict(no=14, label="Moser 1988 (도로공사소음→원조행동)", g=g14, v=v14,
             note="2x2 OR→Chinn, Ro vs R+ 열쇠 도움행동", source="Table 4 cell counts"),
    ]
    o_soc = pool([r["g"] for r in social_rows], [r["v"] for r in social_rows], "사회적 상호작용(자연음→증가)")
    o_soc_nat = pool([g931, g1069], [v931, v1069], "  └ 관찰연구만(931·1069)")

    # ================= MA #4 역방향/지각-행태 상관 (Fisher-z) =================
    corr_rows = [
        dict(no=1076, r=0.564, n=419, label="Xu 2024 (사운드스케이프 쾌적성 ↔ 정적행태)", src="잠재변수 상관"),
        dict(no=1221, r=0.209, n=315, label="Wang 2025 (자연음 사건빈도 ↔ 대기시간)", src="Spearman"),
        dict(no=1177, r=0.400, n=58, label="Béjaïa (음향쾌적 ↔ 보행쾌적, 오전)", src="Spearman, CI 보고"),
        dict(no=980, r=0.551, n=180, label="Bao 2023 (체류시간 ↔ 회복지각)", src="PRSS 상관"),
        dict(no=1018, r=0.360, n=310, label="Sitting/walking 군 상관(대표치)", src="Pearson"),
    ]
    zs, vz = [], []
    for r in corr_rows:
        z, v = r_to_z(r["r"], r["n"]); zs.append(z); vz.append(v)
    o_rev = pool(zs, vz, "지각-행태 상관(양방향 혼재)", back_r=True)

    # ---------- 저장 ----------
    with open(os.path.join(MA, "ma_staying_input.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["no", "label", "g", "v", "note", "source"]); w.writeheader()
        for r in staying_rows: w.writerow({**r, "g": round(r["g"], 4), "v": round(r["v"], 5)})
    with open(os.path.join(MA, "ma_social_input.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["no", "label", "g", "v", "note", "source"]); w.writeheader()
        for r in social_rows: w.writerow({**r, "g": round(r["g"], 4), "v": round(r["v"], 5)})
    with open(os.path.join(MA, "ma_correlation_input.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["no", "label", "r", "n", "src"]); w.writeheader()
        for r in corr_rows: w.writerow(r)

    print("=== MA #2 체류 ===")
    for r in staying_rows: print(f"  [{r['no']}] g={r['g']:+.3f} (v={r['v']:.4f}) {r['label']}")
    print(fmt(o_stay))
    print("\n=== MA #3 사회적 상호작용 ===")
    for r in social_rows: print(f"  [{r['no']}] g={r['g']:+.3f} (v={r['v']:.4f}) {r['label']}")
    print(fmt(o_soc)); print(fmt(o_soc_nat))
    print(f"  민감도: 931 자연음 vs 대조 g={g931c:+.3f}")
    print("\n=== MA #4 지각-행태 상관 (Fisher-z) ===")
    for r in corr_rows: print(f"  [{r['no']}] r={r['r']:+.3f} n={r['n']} {r['label']}")
    print(fmt(o_rev))

    return o_stay, o_soc, o_soc_nat, o_rev, staying_rows, social_rows, corr_rows, g931c


if __name__ == "__main__":
    main()
