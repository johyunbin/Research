# -*- coding: utf-8 -*-
"""
Paper32 — 인용추적 전문 확보 작업목록(사용자 직접 다운로드용)
연세대 프록시 DOI 링크 포함. 우선순위 A(초록으로 행태 확인) / B(초록 미확보).
출력: fulltext/ct_download_worklist.md
"""
import sys, os, csv

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
PROXY = "https://doi-org-ssl.access.yonsei.ac.kr/"


def main():
    rows = list(csv.DictReader(open(os.path.join(FT, "ct_retrieval_status.csv"),
                                    encoding="utf-8-sig")))
    done = [r for r in rows if r["result"] in ("ok", "already")]
    todo = [r for r in rows if r["result"] not in ("ok", "already")]
    A = sorted([r for r in todo if r["screen"] == "RETRIEVE"], key=lambda r: -int(r["year"] or 0))
    B = sorted([r for r in todo if r["screen"] == "UNCERTAIN"], key=lambda r: -int(r["year"] or 0))

    L = ["# Paper32 — 인용추적 전문 확보 작업목록\n",
         "\n등록 프로토콜(osf.io/7ew8q)의 backward/forward citation tracking 이행분. "
         "제목·초록 스크리닝을 통과해 **전문 확인이 필요한 건**이다.\n",
         f"\n- 자동 확보 완료: **{len(done)}편** (`ct_pdf/`) — 오픈액세스 직접 다운로드분\n",
         f"- 직접 확보 필요: **{len(todo)}편** (A {len(A)} · B {len(B)}) — "
         "대부분 출판사 봇 차단(403)이라 자동화 불가\n",
         "\n링크는 연세대 도서관 프록시 DOI다. 로그인된 브라우저에서 열면 바로 출판사 페이지로 간다.\n",
         "\n> 받은 PDF는 파일명 그대로 두고 `ct_pdf/` 에 넣으면 된다(REC 번호로 자동 대조).\n",
         "\n---\n"]

    def block(title, rows_, note):
        L.append(f"\n## {title}\n\n{note}\n\n")
        L.append("| REC | 연도 | 저널 | 행태 아웃컴 | 제목 | 링크 |\n|---|---|---|---|---|---|\n")
        for r in rows_:
            j = (r["journal"] or "")[:34]
            t = (r["title"] or "").replace("|", "/")[:78]
            hint = (r["behavior_hint"] or "—")[:18]
            doi = r["doi"]
            link = f"[열기]({PROXY}{doi})" if doi else "—"
            L.append(f"| {r['rec']} | {r['year']} | {j} | {hint} | {t} | {link} |\n")

    block("A. 우선 확보 — 초록에서 행태 아웃컴 확인됨", A,
          "초록을 읽어 적격 가능성이 실질적이라고 판단한 건. **이쪽이 수율이 높다.**")
    block("B. 차순위 — 초록을 기계적으로 확보하지 못한 건", B,
          "Elsevier 등이 초록을 공개 API에 싣지 않아 제목만으로 판단해야 했던 건. "
          "제목상 배제 근거가 없어 남겼다.")

    if done:
        L.append("\n---\n\n## 이미 확보된 건 (자동 다운로드 완료)\n\n")
        L.append("| REC | 연도 | 저널 | 파일 |\n|---|---|---|---|\n")
        for r in sorted(done, key=lambda r: int(r["rec"])):
            L.append(f"| {r['rec']} | {r['year']} | {(r['journal'] or '')[:34]} | `{r['file']}` |\n")

    L.append("\n---\n\n## 확보 후 절차\n\n"
             "1. `ct_pdf/`에 모이면 텍스트 추출 → 전문심사 패킷 생성\n"
             "2. 본검색과 **동일한 기준·동일한 판정 양식**으로 전문심사(등록 프로토콜 준수)\n"
             "3. 포함분은 추출표·효과크기·MMAT 품질평가까지 같은 파이프라인 통과\n"
             "4. PRISMA 흐름도의 'Identification via other methods' 갈래 수치 확정\n")

    open(os.path.join(FT, "ct_download_worklist.md"), "w", encoding="utf-8").write("".join(L))
    print(f"[저장] fulltext/ct_download_worklist.md — A {len(A)} · B {len(B)} · 확보완료 {len(done)}")


if __name__ == "__main__":
    main()
