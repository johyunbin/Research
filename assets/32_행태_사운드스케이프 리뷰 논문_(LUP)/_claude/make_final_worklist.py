# -*- coding: utf-8 -*-
"""
Paper32 — 남은 전문 확보 대상 최종 목록(사용자 직접 다운로드용)
ct_pdf 편입분은 자동 제외 → 반복 실행 가능. 확보 경로가 다른 것을 우선순위로 묶는다.
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
        "journal of transport", "travel behaviour", "ecological informatics"]
KW = [("behaviou", 3), ("behavior", 3), ("walk", 3), ("pedestrian", 3), ("activity", 2),
      ("visit", 2), ("social", 2), ("stay", 3), ("dwell", 3), ("cycl", 2), ("crowd", 2),
      ("physical activity", 3), ("movement", 2), ("soundscape", 2), ("noise", 1)]


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


def row(r, direct=""):
    link = f"[프록시]({PROXY}{r['doi']})" if r["doi"] else "—"
    if direct:
        link = f"[직접]({direct}) · " + link
    return (f"| {r['rec']} | {r['year']} | {(r['journal'] or '')[:28]} | "
            f"{(r.get('behavior_hint') or '—')[:16]} | "
            f"{(r['title'] or '').replace('|','/')[:58]} | {link} |\n")


def table(L, title, note, rows, direct_map=None):
    if not rows:
        return
    L.append(f"\n## {title} — {len(rows)}편\n\n{note}\n\n")
    L.append("| REC | 연도 | 저널 | 행태 아웃컴 | 제목 | 링크 |\n|---|---|---|---|---|---|\n")
    for r in rows:
        L.append(row(r, (direct_map or {}).get(r["rec"], "")))


def main():
    done = {str(int(m.group(1))) for f in os.listdir(DEST) if (m := re.match(r"CT(\d{4})_", f))}
    st = {r["rec"]: r for r in csv.DictReader(open(os.path.join(FT, "ct_retrieval_status.csv"),
                                                   encoding="utf-8-sig"))}
    res = {r["rec"]: r for r in csv.DictReader(open(os.path.join(FT, "oa_targets_resolved.csv"),
                                                    encoding="utf-8-sig"))}
    rest = sorted([st[k] for k in st if k not in done], key=lambda r: -score(r))
    oa = [r for r in rest if r["oa_status"] != "closed"]
    cl = [r for r in rest if r["oa_status"] == "closed"]
    A = [r for r in cl if r["screen"] == "RETRIEVE"]
    B = [r for r in cl if r["screen"] == "UNCERTAIN"]
    direct = {r["rec"]: (r.get("pdf_try1") or "") for r in res.values()
              if (r.get("pdf_try1") or "").startswith("http")
              and "doi.org" not in (r.get("pdf_try1") or "")}

    L = [f"# Paper32 — 남은 전문 확보 목록 ({len(rest)}편)\n",
         f"\n`ct_pdf/`에 **{len(done)}편** 확보 완료. 아래는 아직 없는 것만이며, "
         "우선순위 순으로 정렬했습니다.\n",
         "\n### 받으신 뒤\n\n"
         "**파일명은 그대로 두시고** 아무 폴더에나 모아 위치만 알려 주세요. "
         "본문 텍스트·DOI로 대조해 자동 편입하고, `수집논문_PDF/`에 기존과 같은 "
         "`년도_저널축약_저자_제목.pdf` 형식으로 정리합니다.\n"
         "\n접근이 막히는 건 **건너뛰셔도 됩니다.** 못 받은 건 PRISMA에 "
         "'not retrieved'로 정직하게 기재합니다.\n"
         "\n링크는 두 종류입니다 — **직접**은 출판사 PDF 바로가기(오픈액세스라 로그인 불요), "
         "**프록시**는 연세대 도서관 경유입니다.\n"]

    table(L, "① 오픈액세스", "로그인 없이 받아집니다. 제가 브라우저로 시도했지만 "
          "SAGE·Elsevier 등에서 봇 차단에 걸려 중단했습니다.", oa, direct)
    table(L, "② 구독 필요 · 초록에서 행태 확인됨", "적격 가능성이 높은 쪽입니다.", A)
    table(L, "③ 구독 필요 · 제목만 판단", "초록을 기계적으로 못 구한 건입니다. "
          "적격률이 낮으니 **여력 되실 때만** 하셔도 됩니다.", B)

    from collections import Counter
    L.append(f"\n---\n\n확보 현황: **{len(done)}편 완료 / {len(rest)}편 남음** "
             f"(오픈액세스 {len(oa)} · 구독A {len(A)} · 구독B {len(B)})\n")
    L.append(f"\n남은 건의 출판사: {dict(Counter((r['doi'] or '').split('/')[0] for r in rest).most_common())}\n")

    open(os.path.join(FT, "DOWNLOAD_사용자용.md"), "w", encoding="utf-8").write("".join(L))
    print(f"[저장] fulltext/DOWNLOAD_사용자용.md")
    print(f"  확보 {len(done)} · 남음 {len(rest)} (OA {len(oa)} · 구독A {len(A)} · 구독B {len(B)})")


if __name__ == "__main__":
    main()
