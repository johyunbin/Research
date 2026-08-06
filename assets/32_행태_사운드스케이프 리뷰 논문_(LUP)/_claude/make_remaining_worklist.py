# -*- coding: utf-8 -*-
"""
Paper32 — 남은 확보 대상 목록(우선순위 재계산). ct_pdf 편입분은 자동 제외되므로 반복 실행 가능.
출력: fulltext/DOWNLOAD_사용자용.md (덮어씀)
"""
import sys, os, csv, re

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
DEST = os.path.join(BASE, "ct_pdf")
PROXY = "https://doi-org-ssl.access.yonsei.ac.kr/"

CORE = ["landscape and urban", "applied acoustics", "building and environment", "urban forestry",
        "cities", "environmental psychology", "environment and behavior", "sustainable cities",
        "science of the total", "transportation research", "health & place", "health and place",
        "urban design", "noise", "acoustics", "soundscape", "landscape research",
        "journal of transport"]
KW = [("behaviou", 3), ("behavior", 3), ("walk", 3), ("pedestrian", 3), ("activity", 2),
      ("visit", 2), ("use ", 2), ("social", 2), ("stay", 3), ("dwell", 3), ("cycl", 2),
      ("physical activity", 3), ("crowd", 2), ("movement", 2)]


def score(r):
    s, y = 0, int(r["year"] or 0)
    s += 3 if y >= 2025 else (2 if y >= 2023 else (1 if y >= 2020 else 0))
    j = (r["journal"] or "").lower()
    if any(c in j for c in CORE):
        s += 3
    t = (r["title"] or "").lower()
    for kw, w in KW:
        if kw in t:
            s += w
    return s


def main():
    done = {str(int(m.group(1))) for f in os.listdir(DEST)
            if (m := re.match(r"CT(\d{4})_", f))}
    st = {r["rec"]: r for r in csv.DictReader(open(os.path.join(FT, "ct_retrieval_status.csv"),
                                                   encoding="utf-8-sig"))}
    fin = {r["rec"]: r for r in csv.DictReader(open(os.path.join(FT, "ct_screen_final.csv"),
                                                    encoding="utf-8-sig"))}
    rest = [st[k] for k in st if k not in done]
    closed = [r for r in rest if r["oa_status"] == "closed"]
    oa = [r for r in rest if r["oa_status"] != "closed"]
    A = sorted([r for r in closed if r["screen"] == "RETRIEVE"], key=lambda r: -score(r))
    B = sorted([r for r in closed if r["screen"] == "UNCERTAIN"], key=lambda r: -score(r))

    L = ["# Paper32 — 남은 다운로드 목록\n",
         f"\n갱신 시각 기준 `ct_pdf/`에 **{len(done)}편** 편입됨. 아래는 아직 없는 것만.\n",
         "\n파일명은 무엇이든 상관없습니다. 아무 폴더에나 모아두고 위치만 알려 주세요 — "
         "본문 텍스트로 대조해 자동 편입합니다.\n"]

    if A:
        L.append(f"\n## A군 잔여 — {len(A)}편 (초록에서 행태 아웃컴 확인됨)\n\n")
        L.append("| REC | 연도 | 저널 | 제목 | 링크 |\n|---|---|---|---|---|\n")
        for r in A:
            L.append(f"| {r['rec']} | {r['year']} | {(r['journal'] or '')[:30]} | "
                     f"{(r['title'] or '').replace('|','/')[:62]} | [열기]({PROXY}{r['doi']}) |\n")

    top = B[:12]
    L.append(f"\n## B군 우선 12편 — 여기까지만 받아주시면 충분합니다\n\n"
             "초록을 기계적으로 못 구해 제목만으로 남긴 31편 중, **연도·저널·제목 신호로 상위 12편**을 "
             "골랐습니다. 나머지 19편은 신호가 약해 생략해도 결론에 영향이 없을 것으로 봅니다.\n\n")
    L.append("| REC | 연도 | 저널 | 제목 | 링크 |\n|---|---|---|---|---|\n")
    for r in top:
        L.append(f"| {r['rec']} | {r['year']} | {(r['journal'] or '')[:30]} | "
                 f"{(r['title'] or '').replace('|','/')[:62]} | [열기]({PROXY}{r['doi']}) |\n")

    L.append(f"\n<details><summary>B군 나머지 {len(B)-len(top)}편 (생략 가능)</summary>\n\n")
    L.append("| REC | 연도 | 저널 | 제목 | 링크 |\n|---|---|---|---|---|\n")
    for r in B[len(top):]:
        L.append(f"| {r['rec']} | {r['year']} | {(r['journal'] or '')[:30]} | "
                 f"{(r['title'] or '').replace('|','/')[:62]} | [열기]({PROXY}{r['doi']}) |\n")
    L.append("\n</details>\n")
    L.append(f"\n---\n\n오픈액세스 {len(oa)}편은 제가 브라우저로 받습니다(목록 `oa_targets_resolved.csv`).\n")

    open(os.path.join(FT, "DOWNLOAD_사용자용.md"), "w", encoding="utf-8").write("".join(L))
    print(f"[저장] fulltext/DOWNLOAD_사용자용.md")
    print(f"  편입 완료 {len(done)} · 남은 사용자 몫 A {len(A)} + B {len(B)}(우선 {len(top)}) · OA(내 몫) {len(oa)}")


if __name__ == "__main__":
    main()
