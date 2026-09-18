# -*- coding: utf-8 -*-
"""
Paper32 — 메타분석 갱신(9월 편입: MA4 에 Yu & Kang 2008 추가, 민감도 행 4개 추가) 원고 반영 (2026-09-18)

- Table 2·3 본문 = fulltext/ma_char_table.md · ma_sensitivity_table.md 생성본 그대로(verify_manuscript 가 대조)
- 수치 = fulltext/ma/ma_forest_data.json · ma_v2_raw.json · ma_sensitivity_v2_raw.json 에서 읽어 서식화
- Yu & Kang (2008) 인용은 {{yu2008}} 자리표시 → insert_citations.py 가 Crossref 확인 후 번호로 바꾼다
- 모든 치환은 원문이 정확히 1회 있을 때만 수행
"""
import json, math, os, sys

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
P = os.path.join(os.path.dirname(BASE), "01_논문작업", "Manuscript_KO.md")
t = open(P, encoding="utf-8").read()


def rep(a, b):
    global t
    n = t.count(a)
    if n != 1:
        sys.exit(f"STOP count={n}: {a[:80]}")
    t = t.replace(a, b)


fd = json.load(open(os.path.join(FT, "ma", "ma_forest_data.json"), encoding="utf-8"))
c = fd["correlation"]["pooled"]
r = lambda z: math.tanh(z)
sgn = lambda x: ("+" if x >= 0 else "−") + f"{abs(x):.2f}"
est, lo, hi = r(c["est"]), r(c["lo"]), r(c["hi"])
assert c["k"] == 7
corr_txt = f"*r* = {sgn(est)}, 95% CI [{sgn(lo)}, {sgn(hi)}], *p* = {c['p']:.3f}"
i2 = f"{c['I2']:.1f}"
print("MA4:", corr_txt, "I2", i2)

# ── Table 2·3: 생성본으로 교체(캡션·주석은 유지) ─────────────────────
def swap_table(caption_start, gen_file):
    global t
    a = t.index(caption_start)
    a = t.index("\n| ", a) + 1                     # 표 첫 행
    b = t.index("\n\n", a)                          # 표 끝
    gen = open(os.path.join(FT, gen_file), encoding="utf-8").read().replace("\r\n", "\n").strip("\n")
    assert gen.startswith("| "), gen_file
    t = t[:a] + gen + t[b:]


swap_table("**Table 2.** Characteristics and effect sizes", "ma_char_table.md")
swap_table("**Table 3.** Pooled estimates and sensitivity analyses", "ma_sensitivity_table.md")
rep("Rows beginning with *Adding* include a study that was excluded from the main analysis (Section 2.9).",
    "Rows beginning with *Adding* include a study that was excluded from the main analysis or used only in sensitivity analyses (Sections 2.7 and 2.9).")

# ── 초록 ──────────────────────────────────────────────────
rep("소리–행태 상관(sound–behaviour correlation; *k* = 6, *r* = +0.42, 95% CI [+0.19, +0.61])",
    f"소리–행태 상관(sound–behaviour correlation; *k* = 7, *r* = {sgn(est)}, 95% CI [{sgn(lo)}, {sgn(hi)}])")

# ── 3.4 ───────────────────────────────────────────────────
rep("네 개의 행태 클러스터가 메타분석 요건을 충족했으며, 16편의 연구가 합성에 기여했다.",
    "네 개의 행태 클러스터가 메타분석 요건을 충족했으며, 17편의 연구가 합성에 기여했다.")
rep("소리–행태 상관 클러스터(*k* = 6)에서는 소리 지표와 행태 사이에 정적 상관이 나타났고(*r* = +0.42, 95% CI [+0.19, +0.61], *p* = 0.006), 이질성이 높았다(*I*² = 92.0%). 개별 상관계수는 +0.16에서 +0.65 사이에 분포했다.",
    f"소리–행태 상관 클러스터(*k* = 7)에서는 소리 지표와 행태 사이에 정적 상관이 나타났고({corr_txt}), 이질성이 높았다(*I*² = {i2}%). "
    "개별 상관계수는 −0.01에서 +0.65 사이에 분포했다.")
rep("가로의 음향 쾌적성과 보행 쾌적성의 상관(*r* = +0.40)[46]은 신뢰구간이 가장 넓었다.",
    "가로의 음향 쾌적성과 보행 쾌적성의 상관(*r* = +0.40)[46]은 신뢰구간이 가장 넓었다. "
    "유럽과 중국의 19개 도시 옥외공간에서 이용자 10,031명을 조사한 연구에서는 조사원이 관찰한 이동 활동과 음량 평가 사이에 "
    "상관이 거의 없었다(*r* = −0.01){{yu2008}}.")

# ── 3.5 ───────────────────────────────────────────────────
s = json.load(open(os.path.join(FT, "ma", "ma_sensitivity_v2_raw.json"), encoding="utf-8"))
rep("보행속도 클러스터는 기여 연구가 모두 MMAT low 등급이어서 low 등급을 제외한 분석을 수행할 수 없었다.",
    "보행속도 클러스터는 기여 연구가 모두 MMAT low 등급이어서 low 등급을 제외한 분석을 수행할 수 없었다. "
    "음환경 조건이 아니라 정숙 안내판 개입을 비교한 탐방로 연구를 보행속도 클러스터에 더하면 *g* = −0.34(*p* = 0.305)였다.")
rep("관찰 기록 간 독립성 문제로 주분석에서 제외한 연구를 포함하면 *g* = +0.81(*p* = 0.005)이었다.",
    "관찰 기록 간 독립성 문제로 주분석에서 제외한 연구를 포함하면 *g* = +0.81(*p* = 0.005)이었다. "
    "결과가 행동 기대로 측정돼 민감도 분석 전용으로 분류한 실험 연구를 더하면 *g* = +0.67(*p* = 0.004, *I*² = 34.1%)이었다.")
rep("소리–행태 상관 클러스터에서는 한 편씩 제외한 분석의 추정치가 *r* = +0.39에서 +0.48 사이였고, 모든 민감도 분석에서 신뢰구간이 0을 배제했다. MMAT low 등급 연구를 제외하면 *r* = +0.47, 인용 추적으로 확보한 연구를 제외하면 *r* = +0.44, 순위상관을 제외하면 *r* = +0.48이었다.",
    "소리–행태 상관 클러스터에서는 한 편씩 제외한 분석의 추정치가 *r* = +0.32에서 +0.42 사이였고, 모든 민감도 분석에서 신뢰구간이 0을 배제했다. "
    "MMAT low 등급 연구를 제외하면 *r* = +0.39, 인용 추적으로 확보한 연구를 제외하면 *r* = +0.35, 순위상관을 제외하면 *r* = +0.48이었다. "
    "측정점 평균 군중밀도를 응답자 수로 분석한 연구를 더하면 *r* = +0.37, 장소 음환경이 아니라 소리의 중요도나 선호를 평정한 두 연구를 더하면 *r* = +0.31이었다.")
rep("기여 효과가 가장 많은 클러스터도 *k* = 6으로 사전에 정한 기준(*k* ≥ 10)에 미달해 출판편향은 평가하지 않았다.",
    "기여 효과가 가장 많은 클러스터도 *k* = 7로 사전에 정한 기준(*k* ≥ 10)에 미달해 출판편향은 평가하지 않았다.")

# ── 4.2·결론 ──────────────────────────────────────────────
rep("여러 행태 결과가 함께 포함돼 이질성이 높았다(*I*² = 92.0%).", f"여러 행태 결과가 함께 포함돼 이질성이 높았다(*I*² = {i2}%).")
rep("98편 가운데 16편이 기여한 네 메타분석 클러스터는", "113편 가운데 17편이 기여한 네 메타분석 클러스터는")

open(P, "w", encoding="utf-8", newline="").write(t)
print("OK — 메타분석 부분 갱신(인용 자리표시 {{yu2008}} 1개)")
