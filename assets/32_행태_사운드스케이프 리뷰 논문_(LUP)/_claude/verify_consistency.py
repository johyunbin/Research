# -*- coding: utf-8 -*-
"""
Paper32 — 산출물 상호 정합성 검증 (투고 전 자체 게이트)
파일 간 같은 수치가 어긋나면 심사자가 잡는다. 여기서 먼저 잡는다.
"""
import sys, os, csv, re
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
MA = os.path.join(FT, "ma")
P, F = [], []          # pass / fail


def chk(name, ok, detail=""):
    (P if ok else F).append(f"{name}" + (f" — {detail}" if detail else ""))


def rd(p):
    return list(csv.DictReader(open(p, encoding="utf-8-sig")))


def main():
    verd = rd(os.path.join(FT, "ft_verdicts_v2.csv"))
    ext = rd(os.path.join(FT, "ft_extraction_v2.csv"))
    qual = rd(os.path.join(FT, "quality_all.csv"))
    t1 = rd(os.path.join(FT, "table1_study_characteristics.csv"))
    dm = rd(os.path.join(FT, "design_matrix.csv"))
    sens = rd(os.path.join(MA, "ma_sensitivity.csv"))

    vc = Counter(r["final_verdict"] for r in verd)
    inc = vc["FINAL_INCLUDE"]; sen = vc["SENS_ONLY"]
    chk("판정 분포", inc == 81 and sen == 3 and vc["UNCERTAIN"] == 0,
        f"INCLUDE {inc} · SENS {sen} · EXCLUDE {vc['FINAL_EXCLUDE']} · UNCERTAIN {vc['UNCERTAIN']}")

    tgt = {int(r["no"]) for r in verd if r["final_verdict"] in ("FINAL_INCLUDE", "SENS_ONLY")}
    chk("추출표 = 판정 대상", {int(r["no"]) for r in ext} == tgt, f"{len(ext)}행 / 대상 {len(tgt)}")
    chk("품질평가 = 판정 대상", {int(r["no"]) for r in qual} == tgt, f"{len(qual)}편")
    chk("Table 1 = 판정 대상", {int(r["sid"]) for r in t1} == set(range(1, len(t1) + 1)) and len(t1) == len(tgt),
        f"{len(t1)}행 · sid 연속")

    # 품질 등급이 Table 1과 quality_all에서 일치하는가
    qmap = {int(r["no"]): r["quality_tier"] for r in qual}
    t1q = Counter(r["quality"] for r in t1)
    qq = Counter(qmap.values())
    chk("품질 등급 분포 일치(Table1 vs quality_all)", t1q == qq, f"{dict(t1q)} vs {dict(qq)}")

    # 방향 분포가 추출표와 Table 1에서 일치
    de = Counter(r["direction"].strip() for r in ext)
    dt = Counter(r["direction"].strip() for r in t1)
    chk("방향 분포 일치(추출표 vs Table1)", de == dt, f"{dict(de)} vs {dict(dt)}")

    # 민감도 주분석이 ma_summary 정본과 일치
    summ = open(os.path.join(MA, "ma_summary.md"), encoding="utf-8").read()
    prim = {r["cluster"]: r for r in sens if r["analysis"] == "주분석"}
    for cl, k, est in [("MA1 보행속도", "4", -0.500), ("MA2 체류", "3", 0.313),
                       ("MA3 사회적 상호작용", "3", 0.569)]:
        r = prim.get(cl)
        ok = r and r["k"] == k and abs(float(r["est"]) - est) < 0.002
        chk(f"민감도 주분석 = 정본 ({cl})", bool(ok),
            f"k={r['k'] if r else '?'} est={r['est'] if r else '?'} (정본 k={k}, {est:+.3f})")
    r4 = prim.get("MA4 지각-행태 상관")
    chk("민감도 주분석 = 정본 (MA4)", r4 and r4["k"] == "5" and abs(float(r4["r_back"]) - 0.426) < 0.002,
        f"k={r4['k'] if r4 else '?'} r={r4.get('r_back') if r4 else '?'} (정본 k=5, r=+0.426)")

    # MA1 저품질 제외가 k=0인가 (논문 핵심 한계)
    lowx = [r for r in sens if r["cluster"] == "MA1 보행속도" and r["analysis"] == "저품질(low) 제외"]
    chk("MA1 저품질 제외 = 풀링 불가", bool(lowx) and lowx[0]["k"] == "0",
        f"k={lowx[0]['k'] if lowx else '?'}")

    # 설계 매트릭스 근거 논문이 전부 실재·FINAL_INCLUDE인가
    vmap = {int(r["no"]): r["final_verdict"] for r in verd}
    bad = []
    for r in dm:
        ids = [int(i) for i in r["evidence_studies"].replace(" ", "").split(";") if i]
        if len(ids) != int(r["n_studies"]):
            bad.append(f"{r['lever_id']} n 불일치")
        for i in ids:
            if vmap.get(i) != "FINAL_INCLUDE":
                bad.append(f"{r['lever_id']}:{i}={vmap.get(i)}")
    chk("설계 매트릭스 근거 = 포함 코퍼스", not bad, "; ".join(bad[:5]) if bad else f"{len(dm)}개 레버")

    # 측정세대 카운트가 그림 소스와 evidence_map에서 일치
    gc = Counter()
    for r in t1:
        for g in (x.strip() for x in r["measure_gen"].split(";")):
            if g in ("G1", "G2", "G3"):
                gc[g] += 1
    em = open(os.path.join(FT, "evidence_map.md"), encoding="utf-8").read()
    m = re.search(r"\*\*합계\*\*\s*\|\s*\*\*(\d+)\*\*\s*\|\s*\*\*(\d+)\*\*\s*\|\s*\*\*(\d+)\*\*", em)
    chk("측정세대 합계(Table1 vs evidence_map)",
        bool(m) and (int(m.group(1)), int(m.group(2)), int(m.group(3))) == (gc["G1"], gc["G2"], gc["G3"]),
        f"Table1 {gc['G1']}/{gc['G2']}/{gc['G3']} vs evidence_map {m.groups() if m else '미검출'}")

    # 인용추적 수치가 PRISMA 서술과 일치
    ctf = rd(os.path.join(FT, "ct_screen_final.csv"))
    ctr = rd(os.path.join(FT, "ct_retrieval_status.csv"))
    fin = Counter(r["final"] for r in ctf)
    got = sum(1 for r in ctr if r["result"] in ("ok", "already"))
    pf = open(os.path.join(FT, "prisma_flow.md"), encoding="utf-8").read()
    seek = fin["RETRIEVE"] + fin["UNCERTAIN"]
    chk("인용추적 스크리닝 총계", len(ctf) == 428, f"{len(ctf)}건")
    chk("인용추적 전문 확보 대상", seek == 79 and f"n = {seek}" in pf, f"{seek}건 · 회수 {got}")
    chk("인용추적 대기 건수 = PRISMA 표기", f"확보 대기 ............................ {len(ctr)-got}" in pf,
        f"대기 {len(ctr)-got}")

    # 그림 전부 존재
    figs = ["Fig1_PRISMA", "Fig2_Forest", "Fig3_EvidenceMap", "Fig4_Direction",
            "Fig5_Methods", "Fig6_Framework", "Fig7_GeoTime", "Fig8_Quality"]
    miss = [f for f in figs for e in ("png", "pdf")
            if not os.path.exists(os.path.join(BASE, "figures", f"{f}.{e}"))]
    chk("Figure 8종 × PNG/PDF", not miss, ", ".join(sorted(set(miss))) if miss else "16개 파일")

    print(f"=== 정합성 검증 — 통과 {len(P)} / 실패 {len(F)} ===\n")
    for x in P:
        print(f"  ✅ {x}")
    if F:
        print()
        for x in F:
            print(f"  ❌ {x}")
    return 1 if F else 0


if __name__ == "__main__":
    sys.exit(main())
