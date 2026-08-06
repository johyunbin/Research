# -*- coding: utf-8 -*-
"""
Paper32 — 인용추적 전문 확보 단계 마감
ct_pdf 실제 파일을 기준으로 회수 상태를 확정하고 PRISMA 수치를 산출한다.
미확보 사유는 사용자 보고(유료 구매 필요)와 접근 실패로 구분해 기록.
출력: fulltext/ct_retrieval_final.csv · ct_retrieval_summary.md
"""
import sys, os, csv, re
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
DEST = os.path.join(BASE, "ct_pdf")

# 사용자가 직접 확인한 미확보 사유(2026-08-03 세션 보고)
PAYWALL = {"399", "391", "190"}


def main():
    done = {str(int(m.group(1))) for f in os.listdir(DEST) if (m := re.match(r"CT(\d{4})_", f))}
    fin = {r["rec"]: r for r in csv.DictReader(open(os.path.join(FT, "ct_screen_final.csv"),
                                                    encoding="utf-8-sig"))}
    st = {r["rec"]: r for r in csv.DictReader(open(os.path.join(FT, "ct_retrieval_status.csv"),
                                                   encoding="utf-8-sig"))}
    tg = [r for r in fin.values() if r["final"] in ("RETRIEVE", "UNCERTAIN")]

    out = []
    for r in tg:
        rec = r["rec"]
        if rec in done:
            state, why = "retrieved", ""
        elif rec in PAYWALL:
            state, why = "not-retrieved", "pay-per-view only (개별 구매 필요)"
        else:
            state, why = "not-retrieved", "기관 구독 범위 밖 또는 접근 차단"
        out.append({"rec": rec, "screen": r["final"], "state": state, "reason": why,
                    "oa_status": st.get(rec, {}).get("oa_status", ""),
                    "year": st.get(rec, {}).get("year", ""),
                    "journal": st.get(rec, {}).get("journal", ""),
                    "doi": r["doi"], "title": r["title"]})

    with open(os.path.join(FT, "ct_retrieval_final.csv"), "w", newline="",
              encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["rec", "screen", "state", "reason", "oa_status",
                                          "year", "journal", "doi", "title"])
        w.writeheader(); w.writerows(sorted(out, key=lambda r: int(r["rec"])))

    A = [r for r in out if r["screen"] == "RETRIEVE"]
    B = [r for r in out if r["screen"] == "UNCERTAIN"]
    gA = sum(1 for r in A if r["state"] == "retrieved")
    gB = sum(1 for r in B if r["state"] == "retrieved")
    nr = [r for r in out if r["state"] == "not-retrieved"]

    L = ["# Paper32 — 인용추적 전문 확보 결과 (마감)\n",
         "\n등록 프로토콜(osf.io/7ew8q)의 backward/forward citation tracking 이행분. "
         "전문 확보 단계를 여기서 마감하고 전문심사로 넘어간다.\n",
         f"\n## 회수율\n\n| 구분 | 확보 | 대상 | 비율 |\n|---|---|---|---|\n",
         f"| **A군** — 초록에서 행태 아웃컴 확인 | **{gA}** | {len(A)} | **{gA/len(A)*100:.0f}%** |\n",
         f"| B군 — 초록 미확보, 제목만 판단 | {gB} | {len(B)} | {gB/len(B)*100:.0f}% |\n",
         f"| 합계 | {gA+gB} | {len(out)} | {(gA+gB)/len(out)*100:.0f}% |\n",
         "\n**A군 회수율이 판단 기준이다.** B군은 초록을 기계적으로 확보할 수 없어("
         "Elsevier 등이 OpenAlex·Crossref·Semantic Scholar 어디에도 초록을 싣지 않음) "
         "제목만으로 '배제할 근거 없음'으로 남긴 집합이며, 적격률이 낮을 것으로 예상된다. "
         "따라서 B군 미확보가 결론을 좌우할 위험은 낮다.\n",
         f"\n## 미확보 {len(nr)}건의 사유\n\n"]
    for k, v in Counter(r["reason"] for r in nr).most_common():
        L.append(f"- {k} — {v}건\n")
    L.append("\n### 유료 구매가 필요해 확보하지 못한 건 (A군)\n\n")
    L.append("| REC | 연도 | 저널 | 제목 |\n|---|---|---|---|\n")
    for r in sorted((x for x in nr if x["rec"] in PAYWALL), key=lambda r: -int(r["year"] or 0)):
        L.append(f"| {r['rec']} | {r['year']} | {(r['journal'] or '')[:32]} | "
                 f"{(r['title'] or '')[:70]} |\n")
    L.append("\n이 3건은 PRISMA 흐름도에 **'reports not retrieved — pay-per-view only'**로 기재한다."
             " 본검색 갈래에서도 같은 사유로 3건을 미확보 처리한 전례가 있다.\n")

    L.append(f"\n## PRISMA 'other methods' 갈래 확정 수치\n\n```\n"
             f"  전문 확보 대상 ............................ n = {len(out)}\n"
             f"  전문 확보 완료 ............................ n = {gA+gB}\n"
             f"  미확보 .................................... n = {len(nr)}\n"
             f"    유료 개별구매 필요 ...................... {sum(1 for r in nr if r['rec'] in PAYWALL)}\n"
             f"    기관 구독 밖·접근 차단 .................. {len(nr)-sum(1 for r in nr if r['rec'] in PAYWALL)}\n"
             f"                                              ↓\n"
             f"  전문 평가 대상 ............................ n = {gA+gB}   ← 다음 단계\n```\n")

    open(os.path.join(FT, "ct_retrieval_summary.md"), "w", encoding="utf-8").write("".join(L))
    print(f"[저장] ct_retrieval_final.csv · ct_retrieval_summary.md")
    print(f"  A군 {gA}/{len(A)} ({gA/len(A)*100:.0f}%) · B군 {gB}/{len(B)} ({gB/len(B)*100:.0f}%) "
          f"· 합계 {gA+gB}/{len(out)}")
    print(f"  미확보 {len(nr)}건: {dict(Counter(r['reason'] for r in nr))}")


if __name__ == "__main__":
    main()
