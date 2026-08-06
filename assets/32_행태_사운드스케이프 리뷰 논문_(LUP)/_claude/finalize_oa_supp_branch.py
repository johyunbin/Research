# -*- coding: utf-8 -*-
"""
Paper32 — OpenAlex 보조검색 갈래 마감 + 3갈래 코퍼스 통합
등록 프로토콜의 세 번째 식별원(OpenAlex 보조검색)을 뒤늦게 이행한 결과를 확정하고,
corpus_v3 → corpus_v4(3갈래)로 확장한다.
출력: fulltext/oa_supp_summary.md · corpus_v4_verdicts.csv · corpus_v4_extraction.csv
"""
import sys, os, csv
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")

EXT_COLS = ["no", "final_verdict", "year", "journal", "title", "country", "setting", "design",
            "sample_n", "exposure", "behaviour_domain", "behaviour_measure",
            "measurement_method", "direction", "key_finding", "effect_stats"]
MULTI = ("behaviour_domain", "measurement_method")


def rd(p):
    return list(csv.DictReader(open(p, encoding="utf-8-sig"))) if os.path.exists(p) else []


def norm_multi(v):
    parts = [p.strip() for p in (v or "").replace(",", ";").split(";") if p.strip()]
    seen, out = set(), []
    for p in parts:
        if p.lower() not in seen:
            seen.add(p.lower()); out.append(p)
    return "; ".join(out)


def main():
    problems = []
    scr = rd(os.path.join(FT, "oa_supp_screen_final.csv"))
    ret = rd(os.path.join(FT, "oa_supp_retrieval.csv"))
    vd = rd(os.path.join(FT, "oas_results_ft", "oasft_01_verdict.csv"))
    ex = rd(os.path.join(FT, "oas_results_ft", "oasft_01_extract.csv"))

    sc = Counter(r["verdict"] for r in scr)
    seek = sc["RETRIEVE"] + sc["UNCERTAIN"]
    pre = [r for r in ret if r["result"] == "prescreen-exclude"]
    got = [r for r in ret if r["result"] in ("ok", "already")]
    notret = [r for r in ret if r["result"] not in ("ok", "already", "prescreen-exclude")]
    vc = Counter(r["verdict"] for r in vd)
    inc = [r for r in vd if r["verdict"] == "FINAL_INCLUDE"]

    if len(vd) != len(got):
        problems.append(f"전문심사 {len(vd)}건 ≠ 확보 {len(got)}건")
    if {r["no"] for r in ex} != {f"OAS{int(r['sid']):04d}" for r in inc}:
        problems.append("추출표 ≠ 포함 집합")

    # ── 3갈래 코퍼스 통합 ─────────────────────────────────────────
    cv = rd(os.path.join(FT, "corpus_v3_verdicts.csv"))
    ce = rd(os.path.join(FT, "corpus_v3_extraction.csv"))
    out_v = [dict(r) for r in cv]
    out_e = [dict(r) for r in ce]
    for r in vd:
        uid = f"OAS{int(r['sid']):04d}"
        out_v.append({"uid": uid, "source": "openalex-supplementary",
                      "final_verdict": r["verdict"], "orig_id": r["sid"]})
    for r in ex:
        row = {c: (r.get(c) or "") for c in EXT_COLS}
        for c in MULTI:
            row[c] = norm_multi(row[c])
        out_e.append({**row, "uid": r["no"], "source": "openalex-supplementary"})

    keep = {r["uid"] for r in out_v if r["final_verdict"] in ("FINAL_INCLUDE", "SENS_ONLY")}
    have = {r["uid"] for r in out_e}
    if keep != have:
        problems.append(f"추출 누락 {sorted(keep-have)} · 과잉 {sorted(have-keep)}")
    if len({r["uid"] for r in out_v}) != len(out_v):
        problems.append("uid 충돌")

    with open(os.path.join(FT, "corpus_v4_verdicts.csv"), "w", newline="",
              encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["uid", "source", "final_verdict", "orig_id"])
        w.writeheader(); w.writerows(out_v)
    with open(os.path.join(FT, "corpus_v4_extraction.csv"), "w", newline="",
              encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["uid", "source"] + EXT_COLS, extrasaction="ignore")
        w.writeheader(); w.writerows(out_e)

    finc = [r for r in out_v if r["final_verdict"] == "FINAL_INCLUDE"]
    by_src = Counter(r["source"] for r in finc)
    sens = sum(1 for r in out_v if r["final_verdict"] == "SENS_ONLY")

    L = [f"# Paper32 — OpenAlex 보조검색 갈래 (등록 세 번째 식별원)\n",
         "\n등록 프로토콜은 OpenAlex를 보조 식별원으로 명시했으나, 3-DB 풀에 없는 고유분이 "
         "**스크리닝되지 않은 채 남아 있었다**(자체 감사에서 발견). 뒤늦게 전수 이행한 결과다.\n",
         f"\n## 흐름\n\n```\n"
         f"  OpenAlex 고유 레코드 ...................... 352\n"
         f"  기존 풀·인용추적·코퍼스와 중복 제거 ....... −23\n"
         f"                                              ↓\n"
         f"  제목·초록 스크리닝 ........................ n = {len(scr)}\n"
         f"  배제 ...................................... −{sc['EXCLUDE']}\n"]
    NAMES = {"E5": "비실증(서평·사설·논평·에세이)", "E4": "관찰가능 행태 아웃컴 없음",
             "E1": "동물·생물음향", "E6": "주제 무관", "E2": "세팅 부적합(실내)",
             "E3": "음환경 변수 없음", "E7": "언어·문헌유형 부적합"}
    for k, v in Counter(r["reason_code"] for r in scr if r["reason_code"]).most_common():
        L.append(f"    {NAMES.get(k, k)} {'.' * max(2, 30 - len(NAMES.get(k, k)))} {v}\n")
    L.append(f"                                              ↓\n"
             f"  전문 확보 대상 ............................ n = {seek}\n"
             f"  문헌유형 사전 배제 ........................ −{len(pre)}  (업계지·코칭 뉴스레터)\n"
             f"  미확보 .................................... −{len(notret)}\n"
             f"                                              ↓\n"
             f"  전문 평가 ................................. n = {len(vd)}\n"
             f"  배제 ...................................... −{vc['FINAL_EXCLUDE']}\n"
             f"                                              ↓\n"
             f"  포함 ...................................... n = {len(inc)}\n```\n")

    L.append(f"\n## 결과 해석\n\n"
             f"**이 갈래의 산출은 {len(inc)}편이며, 메타분석에는 한 편도 기여하지 않는다**(둘 다 질적·서술 "
             f"연구이고 검정통계가 없다). 그럼에도 이행한 이유는 등록한 식별원이기 때문이다.\n\n"
             f"배제 사유 구성이 이 풀의 성격을 보여준다. **비실증이 {Counter(r['reason_code'] for r in scr)['E5']}건"
             f"({Counter(r['reason_code'] for r in scr)['E5']/len(scr)*100:.0f}%)** 으로 압도적인데 서평·사설·"
             f"논평·부고까지 포함된다. OpenAlex가 WoS·Scopus·PubMed보다 훨씬 넓게 색인하므로, "
             f"3-DB에 없다는 것이 곧 '누락'을 뜻하지 않는다는 경험적 근거다.\n\n"
             f"**문헌유형·언어 확인이 결정적이었다.** 전문 확보 5편 중 2편은 OpenAlex 메타데이터가 "
             f"`language=en`이었으나 실제 본문이 한국어·일본어였다(영문은 제목·초록뿐). "
             f"메타데이터만 믿으면 등록 언어 기준을 어길 뻔했다.\n")

    L.append(f"\n## 3갈래 통합 코퍼스\n\n| 갈래 | 포함 |\n|---|---|\n")
    for s, lab in [("db-search", "데이터베이스 검색"), ("citation-tracking", "인용 추적"),
                   ("openalex-supplementary", "OpenAlex 보조검색")]:
        L.append(f"| {lab} | {by_src[s]} |\n")
    L.append(f"| **합계** | **{len(finc)}** |\n\n민감도 분석 전용 {sens}편 별도 · "
             f"분석 대상 총계 {len(finc)+sens}편\n")
    if problems:
        L.append("\n## ⚠️ 검증 문제\n\n" + "\n".join(f"- {x}" for x in problems) + "\n")
    else:
        L.append("\n검증: 전문심사 = 확보 건수 일치 · 추출 = 포함 집합 일치 · uid 충돌 없음.\n")

    open(os.path.join(FT, "oa_supp_summary.md"), "w", encoding="utf-8").write("".join(L))
    print("검증:", "문제 없음" if not problems else problems)
    print(f"스크리닝 {len(scr)} → 확보대상 {seek} → 사전배제 {len(pre)} · 미확보 {len(notret)} → 평가 {len(vd)} → 포함 {len(inc)}")
    print(f"3갈래 코퍼스: {dict(by_src)} · 합계 {len(finc)} · 민감도 {sens}")
    print("[저장] oa_supp_summary.md · corpus_v4_verdicts.csv · corpus_v4_extraction.csv")


if __name__ == "__main__":
    main()
