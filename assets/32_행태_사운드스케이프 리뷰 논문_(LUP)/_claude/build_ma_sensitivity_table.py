# -*- coding: utf-8 -*-
"""
Paper32 — Table 3: 네 클러스터의 합성 추정치와 민감도 분석 (원고 3.5절)

사용자 요청(2026-09-17): LUP 리뷰 예시(Zhang et al. 2025)처럼 결과 절의 기술 방식과 표 구성을
맞출 것. 예시는 합성 결과와 강건성 분석(Robustness analysis)을 따로 보고한다. 우리 원고는
보충자료를 두지 않으므로, 본문에 흩어져 있던 민감도 수치를 이 표 하나로 모으고 본문에는
요점만 남긴다.

수치는 전부 반올림 전 값에서 읽는다(이중 반올림 금지 — 초록 r 에서 실제로 사고가 있었다).
  주분석            = fulltext/ma/ma_forest_data.json   (ma_v2.py)
  예측구간          = fulltext/ma/ma_supplementary.csv  (ma_supplementary_analyses.py)
  LOO·품질·갈래·③④⑤⑥ = fulltext/ma/ma_sensitivity_v2_raw.json (ma_sensitivity_v2.py)
  방문빈도·비독립 연구 추가 = fulltext/ma/ma_v2_raw.json (ma_v2.py)
출력: fulltext/ma_sensitivity_table.md
"""
import sys, os, csv, json, math

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
MA = os.path.join(FT, "ma")

FOREST = json.load(open(os.path.join(MA, "ma_forest_data.json"), encoding="utf-8"))
SENS = json.load(open(os.path.join(MA, "ma_sensitivity_v2_raw.json"), encoding="utf-8"))
V2 = json.load(open(os.path.join(MA, "ma_v2_raw.json"), encoding="utf-8"))
PI = {r["cluster"].split()[0]: r for r in
      csv.DictReader(open(os.path.join(MA, "ma_supplementary.csv"), encoding="utf-8-sig"))
      if r["analysis"] == "prediction_interval"}

MINUS = "−"


def num(x, dp=2, sign=True):
    s = f"{x:+.{dp}f}" if sign else f"{x:.{dp}f}"
    return s.replace("-", MINUS)


def pval(p):
    return "< 0.001" if p < 0.001 else f"{p:.3f}"


def sens(cluster_prefix, analysis):
    hit = [r for r in SENS if r["cluster"].startswith(cluster_prefix) and r["analysis"] == analysis]
    if len(hit) != 1:
        raise SystemExit(f"⚠️ 민감도 결과를 하나로 특정하지 못했다: {cluster_prefix} / {analysis} ({len(hit)}건)")
    return hit[0]


def v2(name):
    hit = [v for k, v in V2.items() if k.startswith(name)]
    if len(hit) != 1:
        raise SystemExit(f"⚠️ ma_v2_raw.json 에서 '{name}' 를 하나로 특정하지 못했다")
    return hit[0]


def est_ci(o, is_r):
    """k ≥ 3 이면 추정치 [95% CI]. k = 2 이면 HK 구간이 t(1) 로 발산하므로 추정치만."""
    e, lo, hi = (o["r"], o["r_lo"], o["r_hi"]) if is_r else (o["est"], o["lo"], o["hi"])
    if o["k"] < 3:
        return num(e)
    return f"{num(e)} [{num(lo)}, {num(hi)}]"


def row(label, o, is_r, pi=""):
    return (f"| {label} | {o['k']} | {est_ci(o, is_r)} | {pi or '—'} | {pval(o['p'])} "
            f"| {o['I2']:.1f}% |")


def loo_row(prefix, is_r):
    loo = [r for r in SENS if r["cluster"].startswith(prefix) and r["analysis"].startswith("LOO")]
    ks = {r["k"] for r in loo}
    es = [r["r"] if is_r else r["est"] for r in loo]
    ps = [r["p"] for r in loo]
    return (f"| Leave-one-out (range) | {'/'.join(str(k) for k in sorted(ks))} "
            f"| {num(min(es))} to {num(max(es))} | — | {pval(min(ps))} to {pval(max(ps))} | — |")


def main():
    out = ["| Analysis | *k* | Estimate [95% CI] | 95% PI | *p* | *I*² |",
           "|---|---|---|---|---|---|"]

    def head(text):
        out.append(f"| **{text}** | | | | | |")

    def main_row(key, pi_key, is_r):
        p = FOREST[key]["pooled"]
        o = dict(p)
        if is_r:
            o.update(r=math.tanh(p["est"]), r_lo=math.tanh(p["lo"]), r_hi=math.tanh(p["hi"]))
        pr = PI[pi_key]
        lo, hi = float(pr["pi_lo"]), float(pr["pi_hi"])
        if is_r:
            lo, hi = math.tanh(lo), math.tanh(hi)
        out.append(row("Main analysis", o, is_r, f"{num(lo)}, {num(hi)}"))

    # ── 보행속도 ────────────────────────────────────────────────────
    head("Walking speed (Hedges' g)")
    main_row("walking", "MA1", False)
    out.append(loo_row("MA1", False))
    if sens("MA1", "저품질 제외")["k"] != 0:
        raise SystemExit("⚠️ 보행속도 low 제외 분석이 추정 가능해졌다 — 표 행 구성을 다시 볼 것")
    out.append("| Excluding low-quality studies | 0 | Not estimable | — | — | — |")
    out.append(row("Sample size from observations", sens("MA1", "③ 관측 n"), False))
    out.append(row("Excluding imputed input", sens("MA1", "⑤ 532 제외"), False))
    out.append(row("Including plaza background-music study", sens("MA1", "⑥ 481 포함"), False))

    # ── 체류 ────────────────────────────────────────────────────────
    head("Staying / dwell time (Hedges' g)")
    main_row("staying", "MA2", False)
    out.append(loo_row("MA2", False))
    out.append(row("Excluding imputed input", sens("MA2", "⑤ 665 제외"), False))
    out.append(row("Including visit-frequency outcome", v2("민감도: 방문빈도 포함"), False))

    # ── 사회적 상호작용 ────────────────────────────────────────────
    head("Social interaction (Hedges' g)")
    main_row("social", "MA3", False)
    out.append(loo_row("MA3", False))
    out.append(row("Excluding study from citation searching", sens("MA3", "인용추적 제외"), False))
    out.append(row("Including study with non-independent observations", v2("민감도: CT0414 추가"), False))

    # ── 소리–행태 상관 ─────────────────────────────────────────────
    head("Sound–behaviour correlation (r)")
    main_row("correlation", "MA4", True)
    out.append(loo_row("MA4", True))
    out.append(row("Excluding low-quality studies", sens("MA4", "저품질 제외"), True))
    out.append(row("Excluding studies from citation searching", sens("MA4", "인용추적 제외"), True))
    out.append(row("Excluding rank correlations", sens("MA4", "④ rho 제외"), True))

    # 표에서 뺀 분석이 정말 주분석과 같은지 확인한다(주석에 "identical" 이라고 쓰므로)
    same = [("MA2", "저품질 제외"), ("MA3", "저품질 제외")]
    for c, a in same:
        m, s = sens(c, "주분석"), sens(c, a)
        if (m["k"], round(m["est"], 10)) != (s["k"], round(s["est"], 10)):
            raise SystemExit(f"⚠️ {c} {a} 가 주분석과 다르다 — 표 주석을 고칠 것")
    for key in ("walking", "staying"):
        if any(e["route"] == "citation-tracking" for e in FOREST[key]["effects"]):
            raise SystemExit(f"⚠️ {key} 에 인용추적 연구가 생겼다 — 표 주석을 고칠 것")
    t1 = {r["uid"]: r for r in csv.DictReader(open(os.path.join(FT, "table1_v2.csv"), encoding="utf-8-sig"))}
    lab = [e["uid"] for k in FOREST for e in FOREST[k]["effects"]
           if "lab" in (t1[str(e["uid"])]["setting"] + t1[str(e["uid"])]["design"]).lower()]
    if lab:
        raise SystemExit(f"⚠️ 실험실 재현 연구가 합성에 기여한다 {lab} — 표 주석을 고칠 것")

    text = "\n".join(out) + "\n"
    open(os.path.join(FT, "ma_sensitivity_table.md"), "w", encoding="utf-8").write(text)
    print(text)
    print("[저장] fulltext/ma_sensitivity_table.md")


if __name__ == "__main__":
    main()
