# -*- coding: utf-8 -*-
"""
Paper32 — PDF 라이브러리 단일화
① 코퍼스 100편 전부 Crossref 서지로 `년도_저널축약_저자_제목.pdf` 이름 생성(기존 49편 포함 일관 재생성)
② 프로젝트 루트 `수집논문_PDF/` 한 폴더로 통합
③ 검증(%PDF·크기·100매 전수) 통과 후에만 중복 삭제:
   Include/*.pdf 원본 49 · Include/renamed/ 전체 · _claude/fulltext/pdf/ID_*.pdf 100
④ 파이프라인용 매핑 `_claude/fulltext/pdf_map.csv` (no ↔ 파일명) 생성
"""
import sys, os, re, csv, glob, time, shutil

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
PROJ = os.path.dirname(BASE)
FT = os.path.join(BASE, "fulltext")
LIB = os.path.join(PROJ, "수집논문_PDF")
INC = os.path.join(PROJ, "Include")
os.makedirs(LIB, exist_ok=True)

sys.path.insert(0, BASE)
from match_and_rename_include import abbrev, safe, crossref  # 동일 규칙 재사용


def main():
    meta = {int(r["no"]): r for r in csv.DictReader(
        open(os.path.join(FT, "fulltext_status_20260803.csv"), encoding="utf-8-sig"))}
    id_files = sorted(glob.glob(os.path.join(FT, "pdf", "ID_*.pdf")))
    print(f"[1] 코퍼스 PDF {len(id_files)}편 — Crossref 서지 조회·이름 생성")

    plan, used = [], set()
    for p in id_files:
        no = int(os.path.basename(p)[3:7])
        m = meta.get(no, {})
        cr = crossref(m["doi"]) if m.get("doi") else None
        time.sleep(0.35)
        year = (cr or {}).get("year") or m.get("year") or "0000"
        journal = (cr or {}).get("journal") or m.get("journal") or "Journal"
        authors = (cr or {}).get("authors") or "Unknown"
        title = (cr or {}).get("title") or m.get("title") or f"record {no}"
        name = f"{year}_{abbrev(journal)}_{safe(authors,28)}_{safe(title,72)}.pdf"
        name = re.sub(r"\s+", " ", name)
        stem, k = name[:-4], 2
        while name.lower() in used:
            name = f"{stem}_{k}.pdf"; k += 1
        used.add(name.lower())
        plan.append({"no": no, "src": p, "name": name})
        if len(plan) % 25 == 0:
            print(f"    서지 {len(plan)}/{len(id_files)}")

    # 복사
    for it in plan:
        shutil.copy2(it["src"], os.path.join(LIB, it["name"]))
    print(f"[2] {LIB} 에 {len(plan)}편 복사 완료")

    # 검증: 전수·매직·크기 일치
    problems = []
    for it in plan:
        d = os.path.join(LIB, it["name"])
        if not os.path.exists(d):
            problems.append(f"누락 {it['name']}"); continue
        if os.path.getsize(d) != os.path.getsize(it["src"]):
            problems.append(f"크기 불일치 {it['name']}")
        with open(d, "rb") as f:
            if f.read(5) != b"%PDF-":
                problems.append(f"매직 불량 {it['name']}")
    lib_count = len(glob.glob(os.path.join(LIB, "*.pdf")))
    if lib_count != len(plan):
        problems.append(f"폴더 파일수 {lib_count} != {len(plan)}")
    print("[3] 검증:", "통과" if not problems else "\n".join("⚠️ " + x for x in problems))

    # 매핑 저장
    with open(os.path.join(FT, "pdf_map.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f); w.writerow(["no", "filename"])
        for it in sorted(plan, key=lambda x: x["no"]):
            w.writerow([it["no"], it["name"]])
    print("[4] pdf_map.csv 저장")

    if problems:
        print("⛔ 검증 실패 — 삭제 단계 건너뜀(중복 보존)"); return

    # 중복 삭제 (검증 통과 시에만)
    n1 = n2 = n3 = 0
    for it in plan:
        os.remove(it["src"]); n1 += 1
    for p in glob.glob(os.path.join(INC, "*.pdf")):
        os.remove(p); n2 += 1
    ren = os.path.join(INC, "renamed")
    if os.path.isdir(ren):
        n3 = len(glob.glob(os.path.join(ren, "*.pdf")))
        shutil.rmtree(ren)
    try:
        os.rmdir(os.path.join(FT, "pdf"))
    except OSError:
        pass
    try:
        os.rmdir(INC)
    except OSError:
        pass
    print(f"[5] 중복 삭제: fulltext/pdf {n1} · Include 원본 {n2} · renamed {n3}")
    print(f"최종: {LIB} 단일 폴더 {lib_count}편")


if __name__ == "__main__":
    main()
