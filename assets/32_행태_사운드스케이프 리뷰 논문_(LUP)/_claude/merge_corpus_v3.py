# -*- coding: utf-8 -*-
"""
Paper32 — 두 갈래 코퍼스 통합 (DB 검색 + 인용추적)
본검색 ft_verdicts_v2/ft_extraction_v2 + 인용추적 ct_verdicts_final/ct_extraction_final
→ 최종 코퍼스 v3. REC 번호와 본검색 no가 충돌하지 않도록 인용추적분은 CT 접두어를 붙인다.
출력: fulltext/corpus_v3_verdicts.csv · corpus_v3_extraction.csv · corpus_v3_summary.md
"""
import sys, os, csv
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")

EXT_COLS = ["no", "final_verdict", "year", "journal", "title", "country", "setting", "design",
            "sample_n", "exposure", "behaviour_domain", "behaviour_measure",
            "measurement_method", "direction", "key_finding", "effect_stats"]


def rd(name):
    p = os.path.join(FT, name)
    return list(csv.DictReader(open(p, encoding="utf-8-sig"))) if os.path.exists(p) else []


def main():
    mv = rd("ft_verdicts_v2.csv")
    me = rd("ft_extraction_v2.csv")
    cv = rd("ct_verdicts_final.csv")
    ce = rd("ct_extraction_final.csv")
    if not cv:
        print("⚠️ ct_verdicts_final.csv 없음 — merge_ct_fulltext.py 먼저 실행"); return

    problems = []
    # ── 판정 통합 ──────────────────────────────────────────────────
    out_v = []
    for r in mv:
        out_v.append({"uid": r["no"], "source": "db-search", "final_verdict": r["final_verdict"],
                      "orig_id": r["no"]})
    ct_keep = 0
    for r in cv:
        uid = f"CT{int(r['rec']):04d}"
        out_v.append({"uid": uid, "source": "citation-tracking",
                      "final_verdict": r["verdict"], "orig_id": r["rec"]})
        if r["verdict"] in ("FINAL_INCLUDE", "SENS_ONLY"):
            ct_keep += 1

    ids = [r["uid"] for r in out_v]
    if len(ids) != len(set(ids)):
        problems.append("uid 충돌 발생")

    # ── 추출 통합 ──────────────────────────────────────────────────
    # 두 갈래가 multi-value 구분자를 다르게 썼다(본검색=세미콜론, 인용추적=쉼표 혼용).
    # 집계·그림이 도메인을 세는 기준이므로 세미콜론으로 통일한다.
    MULTI = ("behaviour_domain", "measurement_method")
    VALID_DOM = {"movement", "staying", "space-use", "social", "activity"}

    def norm_multi(v):
        parts = [p.strip() for p in v.replace(",", ";").split(";") if p.strip()]
        seen, out = set(), []
        for p in parts:
            if p.lower() not in seen:
                seen.add(p.lower()); out.append(p)
        return "; ".join(out)

    out_e = []
    for r in me:
        row = {c: (r.get(c) or "") for c in EXT_COLS}
        for c in MULTI:
            row[c] = norm_multi(row[c])
        out_e.append({**row, "uid": r["no"], "source": "db-search"})
    for r in ce:
        row = {c: (r.get(c) or "") for c in EXT_COLS}
        for c in MULTI:
            row[c] = norm_multi(row[c])
        out_e.append({**row, "uid": f"CT{int(r['no']):04d}", "source": "citation-tracking"})

    for r in out_e:
        bad = [d.strip() for d in r["behaviour_domain"].split(";")
               if d.strip() and d.strip().lower() not in VALID_DOM]
        if bad:
            problems.append(f"{r['uid']}: 도메인 통제어휘 밖 {bad}")

    keep = {r["uid"] for r in out_v if r["final_verdict"] in ("FINAL_INCLUDE", "SENS_ONLY")}
    have = {r["uid"] for r in out_e}
    if keep != have:
        if keep - have:
            problems.append(f"추출 누락: {sorted(keep - have)}")
        if have - keep:
            problems.append(f"추출 과잉: {sorted(have - keep)}")

    cols = ["uid", "source"] + EXT_COLS
    with open(os.path.join(FT, "corpus_v3_verdicts.csv"), "w", newline="",
              encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["uid", "source", "final_verdict", "orig_id"])
        w.writeheader(); w.writerows(out_v)
    with open(os.path.join(FT, "corpus_v3_extraction.csv"), "w", newline="",
              encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
        w.writeheader(); w.writerows(out_e)

    # ── 요약 ───────────────────────────────────────────────────────
    vc = Counter(r["final_verdict"] for r in out_v)
    src = Counter(r["source"] for r in out_e)
    inc = sum(1 for r in out_v if r["final_verdict"] == "FINAL_INCLUDE")
    sen = sum(1 for r in out_v if r["final_verdict"] == "SENS_ONLY")
    db_inc = sum(1 for r in out_v if r["source"] == "db-search" and r["final_verdict"] == "FINAL_INCLUDE")
    ct_inc = inc - db_inc
    dirn = Counter(r["direction"] for r in out_e if r.get("direction"))
    dom = Counter()
    for r in out_e:
        for d in (r.get("behaviour_domain") or "").split(";"):
            if d.strip():
                dom[d.strip()] += 1
    es = sum(1 for r in out_e if (r.get("effect_stats") or "").upper() not in ("", "NR"))

    print("=== 검증 ===")
    print("문제 없음" if not problems else "\n".join("⚠️ " + x for x in problems))
    print(f"\n최종 코퍼스: FINAL_INCLUDE {inc} (DB {db_inc} + 인용추적 {ct_inc}) · SENS_ONLY {sen}")
    print(f"판정 전체: {dict(vc)}")
    print(f"추출 {len(out_e)}행 (출처 {dict(src)}) · 효과크기 보유 {es}행")
    print(f"방향 {dict(dirn)} · 도메인 {dict(dom)}")

    L = ["# Paper32 — 최종 코퍼스 (두 갈래 통합)\n",
         f"\n등록 프로토콜(osf.io/7ew8q)의 두 식별 경로를 모두 이행한 결과.\n",
         f"\n| 갈래 | 포함 | 민감도 전용 |\n|---|---|---|\n",
         f"| 데이터베이스 검색 | {db_inc} | {sum(1 for r in out_v if r['source']=='db-search' and r['final_verdict']=='SENS_ONLY')} |\n",
         f"| 인용 추적 | {ct_inc} | {sum(1 for r in out_v if r['source']=='citation-tracking' and r['final_verdict']=='SENS_ONLY')} |\n",
         f"| **합계** | **{inc}** | **{sen}** |\n",
         f"\n분석 대상 {inc + sen}편(포함 {inc} + 민감도 {sen}) · 효과크기 verbatim 보유 {es}편\n",
         f"\n- 방향: {dict(dirn)}\n- 행태 도메인: {dict(dom)}\n",
         "\n## 식별자 규약\n\n"
         "`uid` — DB 검색분은 기존 `no` 그대로, 인용추적분은 `CT####`(REC 번호 4자리). "
         "두 갈래의 번호 체계가 독립이라 충돌을 막기 위한 접두어이며, `orig_id`에 원 번호를 보존한다.\n"]
    if problems:
        L.append("\n## ⚠️ 검증 문제\n\n" + "\n".join(f"- {x}" for x in problems) + "\n")
    open(os.path.join(FT, "corpus_v3_summary.md"), "w", encoding="utf-8").write("".join(L))
    print("\n[저장] corpus_v3_verdicts.csv · corpus_v3_extraction.csv · corpus_v3_summary.md")


if __name__ == "__main__":
    main()
