# -*- coding: utf-8 -*-
"""
Paper32 — 사용자 직접 다운로드 목록(구독 필요분만) + 내가 처리할 OA 목록 분리
출력: fulltext/DOWNLOAD_사용자용.md · fulltext/oa_targets.csv
"""
import sys, os, csv

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
PROXY = "https://doi-org-ssl.access.yonsei.ac.kr/"

PUB = {"10.1016": "Elsevier", "10.3390": "MDPI", "10.1080": "Taylor & Francis",
       "10.1177": "SAGE", "10.1007": "Springer", "10.3389": "Frontiers",
       "10.1111": "Wiley", "10.1121": "AIP/ASA", "10.1051": "EDP",
       "10.1038": "Nature", "10.1093": "Oxford", "10.1037": "APA",
       "10.1057": "Palgrave", "10.1515": "De Gruyter", "10.1162": "MIT Press",
       "10.3724": "FJYL", "10.15243": "기타", "10.5814": "기타"}


def pub(doi):
    return PUB.get((doi or "").split("/")[0], (doi or "").split("/")[0])


def main():
    rows = [r for r in csv.DictReader(open(os.path.join(FT, "ct_retrieval_status.csv"),
                                           encoding="utf-8-sig"))
            if r["result"] not in ("ok", "already")]
    oa = [r for r in rows if r["oa_status"] != "closed"]
    closed = [r for r in rows if r["oa_status"] == "closed"]
    A = sorted([r for r in closed if r["screen"] == "RETRIEVE"], key=lambda r: -int(r["year"] or 0))
    B = sorted([r for r in closed if r["screen"] == "UNCERTAIN"], key=lambda r: -int(r["year"] or 0))

    L = ["# Paper32 — 다운로드 부탁드릴 목록 (구독 필요분)\n",
         "\n오픈액세스 32편은 제가 따로 받고 있습니다. **여기 있는 건 기관 구독이 필요해 "
         "로그인된 브라우저가 있어야 하는 것들**입니다.\n",
         "\n## 어떻게 하면 되나\n",
         "\n1. 링크(연세대 프록시 DOI)를 클릭하면 출판사 페이지로 바로 갑니다.\n"
         "2. PDF를 받아 **아무 폴더에나** 모아 주세요. **파일명은 그대로 두셔도 됩니다** — "
         "본문 텍스트로 대조해 제가 자동으로 REC 번호에 붙이고 이름도 정리합니다.\n"
         "3. 다 받으시면 폴더 위치만 알려 주세요.\n"
         "4. 접근이 안 되는 건 건너뛰셔도 됩니다. 못 받은 건 PRISMA에 "
         "'not retrieved'로 정직하게 적습니다.\n",
         f"\n## 우선순위 A — {len(A)}편 (이것만 받아주셔도 충분합니다)\n\n"
         "초록에서 관찰가능한 행태 아웃컴을 확인한 건입니다. **여기서 적격이 많이 나옵니다.**\n\n",
         "| REC | 연도 | 저널 | 출판사 | 행태 아웃컴 | 제목 | 링크 |\n|---|---|---|---|---|---|---|\n"]
    for r in A:
        L.append(f"| {r['rec']} | {r['year']} | {(r['journal'] or '')[:30]} | {pub(r['doi'])} | "
                 f"{(r['behavior_hint'] or '—')[:16]} | {(r['title'] or '').replace('|','/')[:66]} | "
                 f"[열기]({PROXY}{r['doi']}) |\n")

    L.append(f"\n## 우선순위 B — {len(B)}편 (여력 되실 때, 안 하셔도 됩니다)\n\n"
             "출판사가 초록을 공개 API에 싣지 않아 **제목만으로 판단**해야 했던 건입니다. "
             "제목상 배제 근거가 없어 남겼을 뿐이라 적격률은 낮을 것으로 봅니다.\n\n")
    L.append("| REC | 연도 | 저널 | 출판사 | 제목 | 링크 |\n|---|---|---|---|---|---|\n")
    for r in B:
        L.append(f"| {r['rec']} | {r['year']} | {(r['journal'] or '')[:30]} | {pub(r['doi'])} | "
                 f"{(r['title'] or '').replace('|','/')[:72]} | [열기]({PROXY}{r['doi']}) |\n")

    from collections import Counter
    L.append(f"\n---\n\n## 참고\n\n"
             f"- 구독 필요 {len(closed)}편의 출판사 분포: "
             f"{dict(Counter(pub(r['doi']) for r in closed).most_common())}\n"
             f"- Elsevier(ScienceDirect)가 대부분입니다. 앞서 자동 접근을 시도했다가 "
             f"프록시 계정 차단 위험이 보여 중단했던 곳이라, 직접 받으시는 것이 안전합니다.\n"
             f"- 제가 처리 중인 오픈액세스 {len(oa)}편은 `oa_targets.csv`에 있습니다.\n")

    out = os.path.join(FT, "DOWNLOAD_사용자용.md")
    open(out, "w", encoding="utf-8").write("".join(L))

    with open(os.path.join(FT, "oa_targets.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["rec", "screen", "oa_status", "year", "journal",
                                          "doi", "landing", "pdf_url", "title"],
                           extrasaction="ignore")
        w.writeheader()
        w.writerows(sorted(oa, key=lambda r: (r["screen"] != "RETRIEVE", -int(r["year"] or 0))))

    print(f"[저장] {out}")
    print(f"  사용자 몫 {len(closed)}편 = A {len(A)} · B {len(B)}")
    print(f"  내 몫(OA) {len(oa)}편 → oa_targets.csv "
          f"(A {sum(1 for r in oa if r['screen']=='RETRIEVE')} · "
          f"B {sum(1 for r in oa if r['screen']=='UNCERTAIN')})")
    print(f"  OA 출판사: {dict(Counter(pub(r['doi']) for r in oa).most_common())}")


if __name__ == "__main__":
    main()
