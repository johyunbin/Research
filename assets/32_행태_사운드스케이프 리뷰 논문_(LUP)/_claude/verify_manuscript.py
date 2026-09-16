# -*- coding: utf-8 -*-
"""
Paper32 — 원고 수치 ↔ 정본 데이터 대조
집필 중 기억으로 쓴 숫자가 섞이는 것을 막는다. 원고에 등장하는 핵심 수치를 정본에서 재산출해 대조.
사용: python verify_manuscript.py [원고경로]
"""
import math
import sys, os, csv, re
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
MA = os.path.join(FT, "ma")
DEFAULT_MS = os.path.join(os.path.dirname(BASE), "01_논문작업",
                          "Manuscript_KO.md")


def rd(p):
    return list(csv.DictReader(open(p, encoding="utf-8-sig")))


def main():
    ms_path = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_MS
    if not os.path.exists(ms_path):
        print(f"⚠️ 원고 없음: {ms_path}"); return 1
    ms = open(ms_path, encoding="utf-8").read()

    cv = rd(os.path.join(FT, "corpus_v4_verdicts.csv"))
    t1 = rd(os.path.join(FT, "table1_v2.csv"))
    q = rd(os.path.join(FT, "quality_v2.csv"))
    ec = rd(os.path.join(FT, "evidence_counts_v2.csv"))

    inc = [r for r in t1 if r["verdict"] == "FINAL_INCLUDE"]
    facts = {}

    facts["최종 포함 96"] = sum(1 for r in cv if r["final_verdict"] == "FINAL_INCLUDE")
    facts["민감도 4"] = sum(1 for r in cv if r["final_verdict"] == "SENS_ONLY")
    facts["DB 갈래 81"] = sum(1 for r in cv if r["final_verdict"] == "FINAL_INCLUDE"
                             and r["source"] == "db-search")
    facts["인용추적 15"] = sum(1 for r in cv if r["final_verdict"] == "FINAL_INCLUDE"
                              and r["source"] == "citation-tracking")

    ys = [int(r["year"]) for r in inc if r["year"].isdigit()]
    facts["2020년 이후 68"] = sum(1 for y in ys if y >= 2020)
    facts["최소 연도 1975"] = min(ys)
    facts["중앙 연도 2022"] = sorted(ys)[len(ys) // 2]

    st = Counter(r["setting"] for r in inc)
    facts["street 36"] = st["street"]
    facts["park 28"] = st["park"]
    de = Counter(r["design"] for r in inc)
    facts["field experiment 17"] = de["field experiment"]

    qt = Counter(r["quality"] for r in inc)
    facts["high 22"] = qt["high"]
    facts["moderate 40"] = qt["moderate"]
    facts["low 34"] = qt["low"]

    # 국가
    import importlib.util
    spec = importlib.util.spec_from_file_location("g7", os.path.join(BASE, "make_fig7_geo_time.py"))
    g7 = importlib.util.module_from_spec(spec)
    try:
        spec.loader.exec_module(g7)
    except SystemExit:
        pass
    ext = rd(os.path.join(FT, "corpus_v4_extraction.csv"))
    cc = Counter()
    for r in ext:
        if r["uid"] not in {x["uid"] for x in inc}:
            continue
        for c in g7.norm_countries(r["country"]):
            cc[c] += 1
    facts["China 40"] = cc["China"]

    dirn = Counter()
    for r in ec:
        if r["kind"] == "direction_x_domain":
            dirn[r["col"]] += int(r["n"])
    facts["forward 110"] = dirn["forward"]
    facts["reverse 74"] = dirn["reverse"]
    facts["both 22"] = dirn["both"]

    gen = Counter()
    for r in ec:
        if r["kind"] == "generation_x_band":
            gen[(r["row"], r["col"])] = int(r["n"])
    facts["G3 2010s 3"] = gen[("2010–2019", "G3")]
    facts["G3 2020s 21"] = gen[("2020–", "G3")]
    facts["G1 2020s 38"] = gen[("2020–", "G1")]
    facts["G2 2020s 28"] = gen[("2020–", "G2")]

    multi = sum(1 for r in inc if len([g for g in r["measure_gen"].split(";")
                                       if g.strip() in ("G1", "G2", "G3")]) >= 2)
    facts["2세대 이상 17"] = multi

    air = sum(int(r["n"]) for r in ec
              if r["kind"] == "domain_x_source" and r["col"] == "항공기소음")
    facts["항공기 4"] = air

    # ── 대조 ───────────────────────────────────────────────────────
    # ⚠️ 라벨에 기대값을 박아두면 정본이 갱신될 때 라벨이 낡는다(실제로 그랬다).
    #    데이터에서 산출한 값이 **원고 본문에 실제로 등장하는지**를 본다.
    ok, bad = [], []
    for label, val in facts.items():
        name = re.sub(r"\s*\d+$", "", label)
        # 숫자가 원고 어딘가에 있는가(천단위 콤마 허용).
        # ⚠️ 한글본에서는 "113건"처럼 숫자 뒤에 한글이 붙어 `\b`가 성립하지 않는다.
        #    앞뒤 경계를 "숫자가 아닌 것"으로 완화한다.
        pats = [rf"(?<!\d){val}(?!\d)", rf"(?<!\d){val:,}(?!\d)"]
        found = any(re.search(p, ms) for p in pats)
        (ok if found else bad).append((name, val))

    must = {"코퍼스 98": r"\b98\b", "DB 갈래 81": r"\b81\b", "인용추적 15": r"\b15\b",
            "보조검색 2": r"\btwo\b", "현장조작 비율 19%": r"19%", "역방향 비율 40%": r"40%"}
    missing = [k for k, pat in must.items() if not re.search(pat, ms)]

    # 메타분석 수치가 원고와 일치하는가
    # ⚠️ 기대값을 여기 박아두면 정본이 바뀔 때 검증기만 낡는다(실제로 낡았다).
    #    ma_forest_data.json(= ma_v2.py 산출)에서 읽어 원고에 그 값이 있는지 본다.
    import json
    fj = json.load(open(os.path.join(MA, "ma_forest_data.json"), encoding="utf-8"))
    NAME = {"walking": "MA1 walking", "staying": "MA2 staying",
            "social": "MA3 social", "correlation": "MA4 correlation"}
    ma_bad = []
    for key, nm in NAME.items():
        p = fj[key]["pooled"]
        # 추정치는 부호 표기가 원고마다 다르므로(−/-) 절대값 문자열로 찾는다.
        # 상관 클러스터는 원고가 역변환 r 로 보고하므로 둘 중 하나만 있으면 통과.
        cands = [f"{abs(p['est']):.3f}"]
        if fj[key].get("back_r"):
            cands += [f"{abs(math.tanh(p['est'])):.3f}", f"{abs(math.tanh(p['est'])):.2f}"]
        if not any(re.search(re.escape(c), ms) for c in cands):
            ma_bad.append(f"{nm} est={cands[0]} — 원고에 없음")
        if not re.search(rf"\bk\b[^\n]{{0,12}}=\s*{p['k']}\b|\|\s*{p['k']}\s*\|", ms):
            ma_bad.append(f"{nm} k={p['k']} — 원고에 없음")

    # leave-one-out 값도 본문에 그대로 인용되므로 S8 산출과 대조한다.
    # (D7 재계산 뒤 §3.4 의 LOO 수치가 구값으로 남아 있던 것을 게이트가 적발 — 2026-09-16)
    sens = rd(os.path.join(MA, "ma_sensitivity_v2.csv"))
    loo = [r for r in sens if r["analysis"].startswith("LOO")]
    for r in loo:
        if r["cluster"].startswith("MA3"):
            g, pv = f"{abs(float(r['est'])):.3f}", f"{float(r['p']):.3f}".lstrip("0")
            if not re.search(rf"{re.escape(g)}\(\*p\* = {re.escape(pv)}\)", ms):
                ma_bad.append(f"MA3 {r['analysis']} g={g} p={pv} — 원고에 없음")
    rs = [float(r["r_back"]) for r in loo if r["cluster"].startswith("MA4") and r["r_back"]]
    if rs:
        lo, hi = f"{min(rs):.3f}".lstrip("0"), f"{max(rs):.3f}".lstrip("0")
        if not re.search(rf"\*r\* = {re.escape(lo)} ~ {re.escape(hi)}", ms):
            ma_bad.append(f"MA4 LOO r 범위 {lo} ~ {hi} — 원고에 없음")

    print(f"=== 원고 수치 검증 — {os.path.basename(ms_path)} ===\n")
    print(f"데이터 산출값이 원고에 존재: {len(ok)} / 누락 {len(bad)}")
    for name, v in bad:
        print(f"  ❌ {name} = {v} → 원고에서 못 찾음")
    if not bad:
        print("  ✅ 전 항목 일치")
    print(f"\n원고 필수 수치 누락: {missing if missing else '없음'}")
    print(f"메타분석 수치 대조: {ma_bad if ma_bad else '✅ 일치'}")

    # 원고에 남은 플레이스홀더
    todo = re.findall(r"\*\[.*?\]\*", ms)
    print(f"\n미작성 구간 {len(todo)}개: {[t[:40] for t in todo]}")
    return 1 if (bad or ma_bad) else 0


if __name__ == "__main__":
    sys.exit(main())
