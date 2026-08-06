# -*- coding: utf-8 -*-
"""
Paper32 — 인용추적 전문심사 결과에 사후 보정 적용
에이전트가 올린 판단 필요 사항 3건을 등록 프로토콜·기존 선례에 맞춰 처리한다.
merge_ct_fulltext.py 실행 뒤에 돌린다.
출력: ct_verdicts_final.csv·ct_extraction_final.csv 갱신 + fulltext/ct_corrections.md
"""
import sys, os, csv

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")

# ── 보정 1: 언어 필터 ────────────────────────────────────────────
# 등록 프로토콜 + 사용자 스코프 결정(2026-08-02): 중국어 문헌 제외.
# REC 410 = 风景园林(Landscape Architecture, 중국어 게재). OpenAlex 서지가 영문 제목을
# 제공해 스크리닝 단계에서 걸러지지 않았고, 전문 심사에서 비로소 드러났다.
LANG_EXCLUDE = {
    "410": "중국어 학술지(风景园林) 게재 — 등록 언어 기준(영어) 미충족. "
           "OpenAlex 영문 제목 때문에 스크리닝에서 미검출, 전문에서 확인",
}


def main():
    vp = os.path.join(FT, "ct_verdicts_final.csv")
    ep = os.path.join(FT, "ct_extraction_final.csv")
    if not os.path.exists(vp):
        print("⚠️ ct_verdicts_final.csv 없음 — merge_ct_fulltext.py 먼저 실행"); return

    verd = list(csv.DictReader(open(vp, encoding="utf-8-sig")))
    ext = list(csv.DictReader(open(ep, encoding="utf-8-sig")))
    log = []

    for r in verd:
        if r["rec"] in LANG_EXCLUDE and r["verdict"] in ("FINAL_INCLUDE", "SENS_ONLY"):
            log.append({"rec": r["rec"], "from": r["verdict"], "to": "FINAL_EXCLUDE",
                        "code": "X7", "why": LANG_EXCLUDE[r["rec"]], "title": r["title"]})
            r["verdict"] = "FINAL_EXCLUDE"
            r["reason_code"] = "X7"
            r["rationale"] = LANG_EXCLUDE[r["rec"]][:120]
    ext = [r for r in ext if r["no"] not in LANG_EXCLUDE]

    with open(vp, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(verd[0].keys())); w.writeheader(); w.writerows(verd)
    with open(ep, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(ext[0].keys())); w.writeheader(); w.writerows(ext)

    from collections import Counter
    vc = Counter(r["verdict"] for r in verd)

    L = ["# Paper32 — 인용추적 전문심사 사후 보정\n",
         "\n전문심사 에이전트가 판단을 유보하고 올린 사항을 등록 프로토콜·기존 선례에 따라 처리한 기록.\n",
         "\n## 보정 1 — 언어 필터 적용 (판정 변경)\n\n"]
    if log:
        L.append("| REC | 변경 | 사유 | 제목 |\n|---|---|---|---|\n")
        for x in log:
            L.append(f"| {x['rec']} | {x['from']} → **{x['to']}** (X7) | {x['why'][:60]} | "
                     f"{(x['title'] or '')[:52]} |\n")
        L.append("\n등록 프로토콜과 사용자 스코프 결정(2026-08-02)에 따라 **중국어 문헌은 제외**한다. "
                 "OpenAlex가 영문 제목·초록 메타데이터를 제공하는 중국어 게재지가 있어 제목·초록 "
                 "스크리닝에서는 걸러지지 않고 전문 단계에서 드러났다. 배제 코드 **X7(언어 기준 미충족)**를 "
                 "신설해 기록한다 — 기존 X1~X6에는 언어 항목이 없었다.\n")
    else:
        L.append("해당 없음.\n")

    L.append("\n## 보정 2 — 판정 변경 없음, 플래그만 (표본 중복 의심)\n\n"
             "**REC 7**(Fang et al. 2023, *J. Outdoor Recreation and Tourism*, n = 2,034)과 "
             "**기존 ID 761**(2021, *Forests*, 중국 도시휴양림, n = 2,034)은 **표본 크기가 정확히 같다.** "
             "제목·저널·연도가 다른 별개 논문이지만 동일 설문 데이터셋을 두 편으로 나눠 발표했을 "
             "가능성이 있다.\n\n"
             "- 두 편 모두 포함 자격 자체는 충족하므로 **배제하지 않는다**(중복 출판이 아니라 "
             "  동일 데이터의 분할 보고일 뿐이며, 서술 종합에는 둘 다 기여한다).\n"
             "- 다만 **같은 메타분석 클러스터에 둘 다 효과크기를 내면 독립성 가정이 깨진다.** "
             "  풀링 단계에서 한 편만 채택하거나 논문 내 평균으로 합성해야 한다 — "
             "  `analysis_rules.md §3`(논문 내 다중 효과 처리)의 확장 적용.\n"
             "- `source_flags.md`에 이 건을 등재해 MA 실행 시 반드시 확인하도록 한다.\n")

    L.append("\n## 보정 3 — 규칙 정합 확인 (변경 없음)\n\n"
             "인용추적 심사 지시의 **P2(행태 아웃컴이 자기보고 의도에 그침 → 배제)**는 "
             "본검색 갈래의 **R2(행동의향 → SENS_ONLY)**와 처리 방향이 달랐다. "
             "지시문 작성 시의 불일치이며, 갈래 간 기준이 어긋나면 안 된다.\n\n"
             "**확인 결과 실제 판정에는 영향이 없었다.** P2가 언급된 4건(REC 21·63·296·328)은 "
             "모두 다른 기준이 먼저 결정적이었다 — 21·296·328은 X4(관찰가능 행태 아웃컴 자체가 없음), "
             "63은 X3(음환경 변수 부재). 즉 '의도만 있어서' 배제된 건은 하나도 없으므로 "
             "R2를 적용했더라도 결과가 같다.\n\n"
             "향후 재현 시에는 **R2(SENS_ONLY)를 정본**으로 삼는다. 프로토콜 이탈 로그에 기록.\n")

    L.append(f"\n---\n\n보정 후 판정: {dict(vc)}\n")
    open(os.path.join(FT, "ct_corrections.md"), "w", encoding="utf-8").write("".join(L))
    print(f"[보정] 언어 필터 {len(log)}건 적용 · 판정 {dict(vc)}")
    print(f"       추출표 {len(ext)}행")
    print("[저장] ct_corrections.md")


if __name__ == "__main__":
    main()
