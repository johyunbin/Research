# -*- coding: utf-8 -*-
"""
Paper32 — PRISMA 2020 흐름 문서(fulltext/prisma_flow.md)를 정본 수치(figures/fig_data.json)에서 생성 (2026-09-18)

종전 prisma_flow.md 는 손으로 쓴 문서라 코퍼스가 바뀔 때마다 옛 수치가 남았다(검토 보고서 §4 #11).
이제 Fig. 1 과 같은 원천(fig_data.json → make_fig1_prisma_v2.py 의 합산 규칙)에서 만든다.
"""
import json, os, sys

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
D = json.load(open(os.path.join(BASE, "figures", "fig_data.json"), encoding="utf-8"))
P = D["prisma"]
DB, CT, SP = P["db"], P["ct"], P["supp"]
W = 44


def line(label, n, sign=""):
    dots = "." * max(3, W - len(label))
    return f"    {label} {dots} {sign}{n:,}" if isinstance(n, int) else f"    {label} {dots} {n}"


def reasons(pairs, indent=6):
    return [" " * indent + f"{k} (n = {v:,})" for k, v in pairs]


L = ["# Paper32 — PRISMA 2020 흐름 (정본 수치 · 자동 생성)", "",
     "생성: `build_prisma_flow_md.py` ← `figures/fig_data.json`(= Fig. 1 과 같은 원천). 손으로 고치지 않는다.",
     "검색 실행일 2026-08-02 · 사전등록 osf.io/7ew8q · 인용 추적 2026-08-03 · OpenAlex 보조 검색 2026-08-06 ·",
     "8월에 확보하지 못한 전문의 추가 확보·평가 2026-09-17–18(무료 원문 + 사용자 확보분).", "",
     "## 1. 데이터베이스 경로", "", "```",
     "IDENTIFICATION",
     line("Web of Science Core Collection", DB["wos"]), line("Scopus", DB["scopus"]), line("PubMed", DB["pubmed"]),
     line("합계", f"n = {DB['identified']:,}"), line("중복 제거", DB["duplicates"], "−"),
     "SCREENING",
     line("제목·초록 스크리닝", f"n = {DB['screened']:,}"), line("배제", DB["excluded"], "−"), *reasons(DB["excl"]),
     "RETRIEVAL / ELIGIBILITY",
     line("전문 확보 대상", f"n = {DB['sought']:,}"), line("전문 미확보", DB["not_retrieved"], "−"), *reasons(DB["nr"]),
     line("전문 평가", f"n = {DB['assessed']:,}"), line("전문 단계 배제", DB["ft_excluded"], "−"), *reasons(DB["ftx"]),
     line("민감도 분석 전용", DB["sens"], "−"),
     "INCLUDED",
     line("질적 종합 포함", f"n = {DB['included']:,}"), "```", "",
     "## 2. 인용 추적 경로", "", "```",
     line("시드(8월 DB 경로 포함·민감도 전용)", f"{CT['seeds']}편"),
     line("후보 레코드(기존 풀에 없던 것)", f"n = {CT['identified']:,} (backward {CT['backward']:,} · forward {CT['forward']:,})"),
     line("제목 기준 자동 선별로 제외", CT["deprioritised"], "−"), *reasons(CT["dep"]),
     line("제목 스크리닝", f"n = {CT['screened_title']:,}"), line("제목 단계 배제", CT["excl_title"], "−"), *reasons(CT["et"]),
     line("초록 스크리닝", f"n = {CT['screened_abs']:,}"), line("초록 단계 배제", CT["excl_abs"], "−"), *reasons(CT["ea"]),
     line("전문 확보 대상", f"n = {CT['sought']:,}"), line("전문 미확보", CT["not_retrieved"], "−"), *reasons(CT["nr"]),
     line("전문 평가", f"n = {CT['assessed']:,}"), line("전문 단계 배제", CT["ft_excluded"], "−"), *reasons(CT["ftx"]),
     line("민감도 분석 전용", CT["sens"], "−"),
     line("질적 종합 포함", f"n = {CT['included']:,}"), "```", "",
     "## 3. OpenAlex 보조 검색 경로", "", "```",
     line("OpenAlex 고유 레코드(3개 DB 에 없는 것)", f"n = {SP['identified']:,}"), line("중복 제거", SP["duplicates"], "−"),
     line("제목·초록 스크리닝", f"n = {SP['screened']:,}"), line("배제", SP["excluded"], "−"), *reasons(SP["excl"]),
     line("전문 확보 대상", f"n = {SP['sought']:,}"), line("확보 전 배제(학술지 논문 아님)", SP["prescreen"], "−"),
     line("전문 미확보", SP["not_retrieved"], "−"), *reasons(SP["nr"]),
     line("전문 평가", f"n = {SP['assessed']:,}"), line("전문 단계 배제", SP["ft_excluded"], "−"), *reasons(SP["ftx"]),
     line("민감도 분석 전용", SP["sens"], "−"),
     line("질적 종합 포함", f"n = {SP['included']:,}"), "```", "",
     "## 4. 최종", "", "```",
     line("리뷰 포함", f"n = {D['n_included']}"),
     line("민감도 분석 전용", f"n = {D['n_sens']}"),
     line("분석 대상", f"n = {D['n_included'] + D['n_sens']}"), "```", "",
     "- Fig. 1 의 '기타 방법' 열은 인용 추적과 보조 검색을 합친 값이다. 보조 검색의 확보 전 배제 2건은 그 열의 "
     "'Reports not retrieved' 에 합산한다(make_fig1_prisma_v2.py 규칙).",
     "- 데이터베이스 경로와 인용 추적 경로의 식별 레코드가 모두 2,073건인 것은 우연의 일치다.",
     "- 전문 배제 사유 코드: X2 세팅 부적격 · X3 음환경 노출(역방향: 행태→음환경 경로) 없음 · X4 관찰 가능한 행태 결과 없음 · "
     "X5 비실증 · X6 학술지 논문 아님 · X7 영어 아님. 레코드별 코드 = `fulltext/ft_exclusion_reasons.csv`, 확보 상태 = `fulltext/retrieval_status_all.csv`."]
assert DB["sought"] - DB["not_retrieved"] == DB["assessed"] and DB["assessed"] - DB["ft_excluded"] - DB["sens"] == DB["included"]
assert CT["sought"] - CT["not_retrieved"] == CT["assessed"] and CT["assessed"] - CT["ft_excluded"] - CT["sens"] == CT["included"]
assert SP["sought"] - SP["prescreen"] - SP["not_retrieved"] == SP["assessed"]
out = os.path.join(BASE, "fulltext", "prisma_flow.md")
open(out, "w", encoding="utf-8", newline="").write("\n".join(L) + "\n")
print("저장:", out, "· 포함", D["n_included"], "· 민감도", D["n_sens"])
