# -*- coding: utf-8 -*-
"""
Paper32 — 2026-09 추가 확보 전문 77건의 전문 텍스트 추출 + 평가 배치 구성

대상: fulltext/manual_pdf_ingest_*.csv 에서 편입된 레코드(DB·CT·OAS) + db_oa_retrieval 의 OK_XML(Europe PMC 전문 XML)
원칙: 전문을 자르지 않는다(8월 D5-3: 46,000자 상한이 MMAT 등급을 깎은 인공물). 텍스트 전체를 저장하고 평가자가 전부 읽는다.
출력: fulltext/txt/TXT_<id>.txt · fulltext/newft_manifest_<태그>.csv · fulltext/newft_batches_<태그>.json
사용: python build_newft_texts.py [태그]   — 태그 생략 시 20260918. 이미 판정이 끝난 레코드(fulltext/newft_final*/verdict.csv)는 제외한다.
"""
import csv, glob, html, json, os, re, sys, unicodedata

import fitz

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
LIB = os.path.join(os.path.dirname(BASE), "수집논문_PDF")
TXT = os.path.join(FT, "txt")
OK = ("moved", "copied", "already_in_library", "name_exists")
TAG = sys.argv[1] if len(sys.argv) > 1 else "20260918"
TARGET_CHARS = 250_000          # 배치당 대략 6만 토큰(서브에이전트 컨텍스트 여유 확보)


def pdf_to_text(path):
    doc = fitz.open(path)
    parts = [f"\n\n=== [page {i + 1}] ===\n" + doc[i].get_text() for i in range(doc.page_count)]
    return unicodedata.normalize("NFKC", "".join(parts)), doc.page_count


def xml_to_text(path):
    s = open(path, encoding="utf-8", errors="ignore").read()
    s = re.sub(r"<(xref|sup)[^>]*>(.*?)</\1>", r"\2", s, flags=re.S)
    s = re.sub(r"</(p|title|sec|caption|table-wrap|fig|tr|ref|abstract|article-title)>", "\n", s)
    s = re.sub(r"</t[dh]>", "\t", s)
    s = re.sub(r"<[^>]+>", "", s)
    s = html.unescape(s)
    s = re.sub(r"[ \t]+\n", "\n", s)
    s = re.sub(r"\n{3,}", "\n\n", s)
    return unicodedata.normalize("NFKC", s), 0


def main():
    os.makedirs(TXT, exist_ok=True)
    meta = {}
    for r in csv.DictReader(open(os.path.join(FT, "db_oa_retrieval_20260917_233746.csv"), encoding="utf-8-sig")):
        meta["DB:" + r["no"]] = r
    for r in csv.DictReader(open(os.path.join(FT, "ct_retrieval_final.csv"), encoding="utf-8-sig")):
        meta["CT:" + f"CT{int(r['rec']):04d}"] = r
    for r in csv.DictReader(open(os.path.join(FT, "oa_supp_retrieval.csv"), encoding="utf-8-sig")):
        meta["OAS:" + f"OAS{int(r['sid']):04d}"] = r

    items = {}
    for lf in sorted(glob.glob(os.path.join(FT, "manual_pdf_ingest_*.csv"))):
        for r in csv.DictReader(open(lf, encoding="utf-8-sig")):
            if r["id"] and r["action"] in OK and r["library_name"]:
                items[(r["branch"], r["id"])] = os.path.join(LIB, r["library_name"])
    for r in csv.DictReader(open(os.path.join(FT, "db_oa_retrieval_20260917_233746.csv"), encoding="utf-8-sig")):
        if r["status"] == "OK_XML" and ("DB", r["no"]) not in items:
            items[("DB", r["no"])] = os.path.join(BASE, "db_oa_pdf", r["file"])

    done = set()
    for f in glob.glob(os.path.join(FT, "newft_final*", "verdict.csv")):
        done |= {r["id"] for r in csv.DictReader(open(f, encoding="utf-8-sig"))}
    items = {k: v for k, v in items.items() if k[1] not in done}
    print(f"판정 완료 제외 {len(done)}건 → 이번 대상 {len(items)}건")

    rows = []
    for (branch, rid), path in sorted(items.items(), key=lambda kv: (kv[0][0], kv[0][1].zfill(8))):
        assert os.path.exists(path), path
        text, pages = xml_to_text(path) if path.endswith(".xml") else pdf_to_text(path)
        tid = f"{int(rid):04d}" if rid.isdigit() else rid
        out = os.path.join(TXT, f"TXT_{tid}.txt")
        if os.path.exists(out) and branch == "DB" and not open(out, encoding="utf-8").read(200).startswith("### RECORD"):
            # 8월 DB 전문 텍스트와 이름 충돌 금지 — 새 레코드는 8월 코퍼스와 겹치지 않아야 한다
            raise SystemExit(f"기존 텍스트와 충돌: {out}")
        m = meta[f"{branch}:{rid}"]
        header = (f"### RECORD {rid} | branch={branch} | year={m.get('year','')} | journal={m.get('journal','')}\n"
                  f"### TITLE: {m.get('title','')}\n### DOI: {m.get('doi','')}\n### SOURCE FILE: {os.path.basename(path)}\n")
        open(out, "w", encoding="utf-8").write(header + text)
        rows.append({"id": rid, "branch": branch, "year": m.get("year", ""), "journal": m.get("journal", ""),
                     "title": m.get("title", ""), "doi": m.get("doi", ""),
                     "screening": m.get("verdict") or m.get("screen", ""), "source_file": os.path.basename(path),
                     "txt": os.path.basename(out), "pages": pages, "chars": len(text)})
    with open(os.path.join(FT, f"newft_manifest_{TAG}.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)

    # 글자 수 균형 배치(큰 것부터 가장 가벼운 배치에 배정)
    n_batches = max(1, round(sum(r["chars"] for r in rows) / TARGET_CHARS))
    batches = [{"batch": f"nft_{i + 1:02d}", "chars": 0, "ids": []} for i in range(n_batches)]
    for r in sorted(rows, key=lambda r: -r["chars"]):
        b = min(batches, key=lambda b: b["chars"])
        b["ids"].append(r["id"])
        b["chars"] += r["chars"]
    json.dump(batches, open(os.path.join(FT, f"newft_batches_{TAG}.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    from collections import Counter
    print("레코드", len(rows), Counter(r["branch"] for r in rows))
    print("총 글자", sum(r["chars"] for r in rows), "· 최대", max(r["chars"] for r in rows), "· 최소", min(r["chars"] for r in rows))
    print("짧은 전문(<8000자):", [(r["id"], r["chars"], r["pages"]) for r in rows if r["chars"] < 8000])
    for b in batches:
        print(b["batch"], len(b["ids"]), b["chars"])


if __name__ == "__main__":
    main()
