# -*- coding: utf-8 -*-
"""
Paper32 — identification 병합·중복제거 (PRISMA 2020 대비)
입력: WoS RIS×2, Scopus RIS×1, PubMed CSV, (보조) OpenAlex CSV
키: DOI(정규화) 우선, DOI 없으면 정규화 제목
출력: merged_pool_*.csv (스크리닝 풀) + prisma_numbers_*.json
"""
import sys, csv, json, re, unicodedata

sys.stdout.reconfigure(encoding="utf-8")
TIMECODE = "20260802_211715"
BASE = r"C:\Users\wh850\Research\assets\32_행태_사운드스케이프 리뷰 논문_(LUP)\_claude"


def norm_doi(d):
    d = (d or "").strip().lower()
    d = re.sub(r"^https?://(dx\.)?doi\.org/", "", d)
    return d


def norm_title(t):
    t = unicodedata.normalize("NFKD", (t or "").lower())
    return re.sub(r"[^a-z0-9]", "", t)


def parse_ris(path):
    recs, cur = [], {}
    with open(path, encoding="utf-8-sig") as f:
        for line in f:
            line = line.rstrip("\n\r")
            m = re.match(r"^([A-Z][A-Z0-9])  - ?(.*)$", line)
            if not m:
                continue
            tag, val = m.group(1), m.group(2).strip()
            if tag == "TY":
                cur = {}
            elif tag == "ER":
                recs.append(cur)
                cur = {}
            elif tag in ("TI", "T1"):
                cur.setdefault("title", val)
            elif tag == "PY":
                cur.setdefault("year", val[:4])
            elif tag in ("T2", "JO", "JF"):
                cur.setdefault("journal", val)
            elif tag in ("DO", "DI"):
                cur.setdefault("doi", val)
            elif tag == "AB":
                cur["has_abstract"] = True
    return recs


def load_csv(path, doi_col, title_col, year_col, journal_col):
    rows = []
    with open(path, encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            rows.append({
                "title": r.get(title_col, ""), "year": str(r.get(year_col, ""))[:4],
                "journal": r.get(journal_col, ""), "doi": r.get(doi_col, ""),
            })
    return rows


def main():
    wos = parse_ris(f"{BASE}\\ris\\wos_0001-1000_20260802.ris") + \
          parse_ris(f"{BASE}\\ris\\wos_1001-1010_20260802.ris")
    scopus = parse_ris(f"{BASE}\\ris\\scopus_0001-0850_20260802.ris")
    pubmed = load_csv(f"{BASE}\\pubmed_results_20260802_211715.csv", "doi", "title", "year", "journal")
    openalex = load_csv(f"{BASE}\\openalex_full_20260802_211715.csv", "doi", "title", "year", "journal")
    print(f"파싱: WoS {len(wos)} · Scopus {len(scopus)} · PubMed {len(pubmed)} · OpenAlex(보조) {len(openalex)}")

    merged, seen = [], {}
    dup_within = 0
    for src, recs in (("wos", wos), ("scopus", scopus), ("pubmed", pubmed)):
        for r in recs:
            key = norm_doi(r.get("doi")) or ("t:" + norm_title(r.get("title")))
            if key in ("", "t:"):
                key = f"blank:{src}:{len(merged)}"
            if key in seen:
                seen[key]["sources"].add(src)
                dup_within += 1
            else:
                entry = {
                    "key": key, "title": r.get("title", ""), "year": r.get("year", ""),
                    "journal": r.get("journal", ""), "doi": norm_doi(r.get("doi")),
                    "sources": {src},
                }
                seen[key] = entry
                merged.append(entry)

    official = len(merged)
    print(f"공식 3DB 병합: {len(wos)+len(scopus)+len(pubmed)} → 중복 {dup_within} 제거 → 고유 {official}건")

    # 보조: OpenAlex에만 있는 레코드 (추가 후보)
    oa_only = []
    for r in openalex:
        key = norm_doi(r.get("doi")) or ("t:" + norm_title(r.get("title")))
        if key not in seen and key not in ("", "t:"):
            oa_only.append(r)
    print(f"OpenAlex 고유분(공식 풀 밖): {len(oa_only)}건 — 보조 식별원 후보")

    out = f"{BASE}\\merged_pool_{TIMECODE}.csv"
    with open(out, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["no", "sources", "doi", "year", "journal", "title"])
        for i, e in enumerate(sorted(merged, key=lambda x: (x["year"], x["title"])), 1):
            w.writerow([i, "+".join(sorted(e["sources"])), e["doi"], e["year"], e["journal"], e["title"]])
    print(f"스크리닝 풀 저장: {out}")

    oa_out = f"{BASE}\\openalex_supplement_{TIMECODE}.csv"
    with open(oa_out, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["doi", "year", "journal", "title"])
        for r in oa_only:
            w.writerow([norm_doi(r.get("doi")), r.get("year", ""), r.get("journal", ""), r.get("title", "")])

    src_counts = {}
    for e in merged:
        for s in e["sources"]:
            src_counts[s] = src_counts.get(s, 0) + 1
    prisma = {
        "search_date": "2026-08-02",
        "identification": {"wos_core": len(wos), "scopus": len(scopus), "pubmed": len(pubmed)},
        "total_identified": len(wos) + len(scopus) + len(pubmed),
        "duplicates_removed": dup_within,
        "records_to_screen": official,
        "openalex_supplementary_unique": len(oa_only),
        "overlap_note": {"in_"+k: v for k, v in sorted(src_counts.items())},
    }
    with open(f"{BASE}\\prisma_numbers_{TIMECODE}.json", "w", encoding="utf-8") as f:
        json.dump(prisma, f, ensure_ascii=False, indent=2)
    print("PRISMA 수치 저장:", json.dumps(prisma["identification"], ensure_ascii=False),
          f"→ 스크리닝 {official}건")


if __name__ == "__main__":
    main()
