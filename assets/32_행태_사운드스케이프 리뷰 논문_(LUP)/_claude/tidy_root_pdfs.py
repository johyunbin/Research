# -*- coding: utf-8 -*-
"""
Paper32 — 프로젝트 루트에 흩어진 PDF 를 수집논문_PDF/ 로 통합

① 루트 *.pdf 의 DOI 를 본문·메타에서 추출 (match_and_rename_include.py 방식 재사용)
② 수집논문_PDF/ 전 파일의 DOI 를 같은 방식으로 추출해 색인 (1회 캐시)
③ 판정:
   - DUP    : 같은 DOI 가 라이브러리에 이미 있다 → 루트 사본 삭제
   - MOVE   : 라이브러리에 없다 → `년도_저널축약_저자_제목.pdf` 로 리네임해 이동
   - HOLD   : DOI 미확보(스캔본 등) → 건드리지 않고 보고 (사람이 판정)
④ 로그: fulltext/root_pdf_tidy_log.csv

사용:  python tidy_root_pdfs.py           # dry-run (파일 안 건드림)
       python tidy_root_pdfs.py --apply   # 실제 이동·삭제
"""
import sys, os, re, csv, glob, json, time, shutil, argparse
import fitz

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
PROJ = os.path.dirname(BASE)
LIB = os.path.join(PROJ, "수집논문_PDF")
FT = os.path.join(BASE, "fulltext")
CACHE = os.path.join(FT, "library_doi_cache.json")

sys.path.insert(0, BASE)
from match_and_rename_include import abbrev, safe, extract_doi, crossref   # 규칙 재사용


def lib_doi_index(refresh=False):
    """라이브러리 파일명 → DOI. 파일 목록이 바뀌면 자동 재작성."""
    files = sorted(os.path.basename(p) for p in glob.glob(os.path.join(LIB, "*.pdf")))
    if os.path.exists(CACHE) and not refresh:
        c = json.load(open(CACHE, encoding="utf-8"))
        if sorted(c) == files:
            return c
    idx = {}
    for i, name in enumerate(files, 1):
        doi, _ = extract_doi(os.path.join(LIB, name))
        idx[name] = doi or ""
        if i % 20 == 0:
            print(f"  라이브러리 색인 {i}/{len(files)}")
    json.dump(idx, open(CACHE, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    return idx


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    a = ap.parse_args()

    loose = sorted(glob.glob(os.path.join(PROJ, "*.pdf")))
    print(f"루트 PDF {len(loose)}개 · 라이브러리 {len(glob.glob(os.path.join(LIB, '*.pdf')))}개")
    idx = lib_doi_index()
    by_doi = {}
    for name, d in idx.items():
        if d:
            by_doi.setdefault(d, name)

    log, plan = [], {"DUP": [], "MOVE": [], "HOLD": []}
    for p in loose:
        base = os.path.basename(p)
        doi, _head = extract_doi(p)
        if not doi:
            plan["HOLD"].append((base, "", "DOI 미확보(텍스트 레이어 없음?)"))
            log.append({"file": base, "action": "HOLD", "doi": "", "target": "",
                        "note": "DOI 미확보"})
            continue
        if doi in by_doi:
            plan["DUP"].append((base, doi, by_doi[doi]))
            log.append({"file": base, "action": "DUP-DELETE", "doi": doi,
                        "target": by_doi[doi], "note": "라이브러리에 동일 DOI 존재"})
            continue
        cr = crossref(doi)
        time.sleep(0.4)
        if not cr or not cr.get("year") or not cr.get("title"):
            plan["HOLD"].append((base, doi, "Crossref 서지 미확보"))
            log.append({"file": base, "action": "HOLD", "doi": doi, "target": "",
                        "note": "Crossref 서지 미확보"})
            continue
        newname = f"{cr['year']}_{abbrev(cr['journal'])}_{safe(cr['authors'], 28)}_{safe(cr['title'], 72)}.pdf"
        newname = re.sub(r"\s+", " ", newname)
        if os.path.exists(os.path.join(LIB, newname)):
            # DOI 는 새 것인데 파일명이 충돌 — 덮어쓰지 않고 보류
            plan["HOLD"].append((base, doi, f"파일명 충돌: {newname}"))
            log.append({"file": base, "action": "HOLD", "doi": doi, "target": newname,
                        "note": "파일명 충돌(다른 DOI)"})
            continue
        plan["MOVE"].append((base, doi, newname))
        log.append({"file": base, "action": "MOVE", "doi": doi, "target": newname, "note": ""})

    print(f"\n판정: MOVE {len(plan['MOVE'])} · DUP-DELETE {len(plan['DUP'])} · HOLD {len(plan['HOLD'])}")
    for base, doi, new in plan["MOVE"]:
        print(f"  → {base}\n     {new}")
    for base, doi, tgt in plan["DUP"]:
        print(f"  ✕ {base}  (= {tgt})")
    for base, doi, why in plan["HOLD"]:
        print(f"  ? {base}  — {why}")

    if a.apply:
        for base, doi, new in plan["MOVE"]:
            shutil.move(os.path.join(PROJ, base), os.path.join(LIB, new))
        for base, doi, tgt in plan["DUP"]:
            os.remove(os.path.join(PROJ, base))
        # 라이브러리가 바뀌었으니 캐시에 신규 반영
        idx.update({new: doi for _, doi, new in plan["MOVE"]})
        json.dump(idx, open(CACHE, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        print(f"\n[적용] 이동 {len(plan['MOVE'])} · 삭제 {len(plan['DUP'])} · 보류 {len(plan['HOLD'])}")
    else:
        print("\n[dry-run] 파일을 건드리지 않았다. 실행: --apply")

    with open(os.path.join(FT, "root_pdf_tidy_log.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["file", "action", "doi", "target", "note"])
        w.writeheader(); w.writerows(log)


if __name__ == "__main__":
    main()
