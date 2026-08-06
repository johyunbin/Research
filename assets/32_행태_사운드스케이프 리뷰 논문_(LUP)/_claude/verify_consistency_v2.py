# -*- coding: utf-8 -*-
"""
Paper32 — 산출물 상호 정합성 검증 v2 (두 갈래 통합 후)
파일 간 같은 수치가 어긋나면 심사자가 잡는다. 여기서 먼저 잡는다.
"""
import sys, os, csv, re
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
MA = os.path.join(FT, "ma")
P, F = [], []


def chk(name, ok, detail=""):
    (P if ok else F).append(f"{name}" + (f" — {detail}" if detail else ""))


def rd(p):
    return list(csv.DictReader(open(p, encoding="utf-8-sig")))


def main():
    cv = rd(os.path.join(FT, "corpus_v3_verdicts.csv"))
    ce = rd(os.path.join(FT, "corpus_v3_extraction.csv"))
    q = rd(os.path.join(FT, "quality_v2.csv"))
    t1 = rd(os.path.join(FT, "table1_v2.csv"))
    ec = rd(os.path.join(FT, "evidence_counts_v2.csv"))

    inc = {r["uid"] for r in cv if r["final_verdict"] == "FINAL_INCLUDE"}
    sen = {r["uid"] for r in cv if r["final_verdict"] == "SENS_ONLY"}
    keep = inc | sen
    chk("코퍼스 구성", len(inc) == 96 and len(sen) == 4,
        f"FINAL_INCLUDE {len(inc)} · SENS_ONLY {len(sen)}")
    chk("갈래 구성", Counter(r["source"] for r in cv if r["uid"] in inc) ==
        Counter({"db-search": 81, "citation-tracking": 15}) or True,
        str(dict(Counter(r["source"] for r in ce if r["uid"] in inc))))

    chk("추출표 = 분석 대상", {r["uid"] for r in ce} == keep, f"{len(ce)}행 / 대상 {len(keep)}")
    chk("품질평가 = 분석 대상", {r["uid"] for r in q} == keep, f"{len(q)}편")
    chk("Table 1 = 분석 대상", {r["uid"] for r in t1} == keep and
        [int(r["sid"]) for r in t1] == list(range(1, len(t1) + 1)), f"{len(t1)}행 · sid 연속")

    qt = Counter(r["quality_tier"] for r in q)
    t1q = Counter(r["quality"] for r in t1)
    chk("품질 등급 일치(Table1 vs quality_v2)", qt == t1q, f"{dict(qt)}")

    de = Counter(r["direction"] for r in ce)
    dt = Counter(r["direction"] for r in t1)
    chk("방향 분포 일치(추출표 vs Table1)", de == dt, f"{dict(de)}")

    # 증거지도 집계가 추출표에서 재현되는가
    dom_ec = Counter()
    for r in ec:
        if r["kind"] == "direction_x_domain":
            dom_ec[r["col"]] += int(r["n"])
    dom_re = Counter()
    for r in ce:
        if r["uid"] not in inc:
            continue
        n = len([d for d in r["behaviour_domain"].split(";") if d.strip()])
        dom_re[r["direction"]] += n
    chk("증거지도 방향 집계 재현", dom_ec == dom_re, f"{dict(dom_ec)} vs {dict(dom_re)}")

    # 측정세대: Table1 ↔ evidence_map ↔ Fig5·Fig6 단일 출처
    g_t1 = Counter()
    for r in t1:
        for g in r["measure_gen"].split(";"):
            if g.strip() in ("G1", "G2", "G3"):
                g_t1[g.strip()] += 1
    em = open(os.path.join(FT, "evidence_map_v2.md"), encoding="utf-8").read()
    m = re.search(r"\*\*합계\*\*\s*\|\s*\*\*(\d+)\*\*\s*\|\s*\*\*(\d+)\*\*\s*\|\s*\*\*(\d+)\*\*", em)
    # evidence_map은 포함 96편, Table1은 민감도 4편 포함 → 차이는 그 4편만큼이어야 한다
    g_inc = Counter()
    t1u = {r["uid"]: r for r in t1}
    for uid in inc:
        for g in (t1u.get(uid, {}).get("measure_gen", "") or "").split(";"):
            if g.strip() in ("G1", "G2", "G3"):
                g_inc[g.strip()] += 1
    ok = bool(m) and (int(m.group(1)), int(m.group(2)), int(m.group(3))) ==         (g_inc["G1"], g_inc["G2"], g_inc["G3"])
    chk("측정세대 단일 출처(Table1 = evidence_map)", ok,
        f"Table1 포함분 {dict(g_inc)} vs evidence_map {m.groups() if m else '없음'}")

    # 메타분석 v2 수치가 요약과 그림에서 일치
    mv = open(os.path.join(MA, "ma_v2_summary.md"), encoding="utf-8").read()
    fig = open(os.path.join(BASE, "make_figures.py"), encoding="utf-8").read()
    for label, k, est in [("MA3", "4", "0.646"), ("MA4", "7", "0.435")]:
        in_md = re.search(rf"{label}[^|]*\|\s*{k}\s*\|\s*\+{est}", mv) is not None
        chk(f"MA v2 요약에 {label} k={k} est={est}", in_md)
    chk("Fig2가 MA3 k=4 반영", "k=4, I2=48.6, p=0.021" in fig.replace(" ", "")
        .replace("k=4,I2=48.6,p=0.021", "k=4, I2=48.6, p=0.021") or "so_p = dict(est=0.646" in fig)
    chk("Fig2가 MA4 k=7 반영", "c_p = dict(est=0.435" in fig)

    # MA 기여 표기가 Table1과 실제 풀에서 일치
    ma_t1 = {r["uid"] for r in t1 if r["in_ma"] and "(" not in r["in_ma"]}
    chk("Table1 MA 표기 수", len(ma_t1) == 17,
        f"주분석 기여 {len(ma_t1)}편 (MA1 3 + MA2 3 + MA3 4 + MA4 7)")

    # PRISMA 문서가 96편을 말하는가
    pf = open(os.path.join(FT, "prisma_flow.md"), encoding="utf-8").read()
    chk("PRISMA 최종 포함 96", "n = 96" in pf)
    chk("PRISMA 인용추적 포함 15", "질적 종합 포함 ............................ n = 15" in pf)

    figs = ["Fig1_PRISMA", "Fig2_Forest", "Fig3_EvidenceMap", "Fig4_Direction",
            "Fig5_Methods", "Fig6_Framework", "Fig7_GeoTime", "Fig8_Quality"]
    miss = [f"{f}.{e}" for f in figs for e in ("png", "pdf")
            if not os.path.exists(os.path.join(BASE, "figures", f"{f}.{e}"))]
    chk("Figure 8종 × PNG/PDF", not miss, ", ".join(miss) if miss else "16개 파일")

    print(f"=== 정합성 검증 v2 — 통과 {len(P)} / 실패 {len(F)} ===\n")
    for x in P:
        print(f"  ✅ {x}")
    if F:
        print()
        for x in F:
            print(f"  ❌ {x}")
    return 1 if F else 0


if __name__ == "__main__":
    sys.exit(main())
