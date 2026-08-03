# -*- coding: utf-8 -*-
"""
Paper32 — PubMed 정식 검색(E-utilities, 계정 불요) + OpenAlex 전량 수집(cursor)
+ 두 소스 DOI 중복 통계. 출력: pubmed_results_*.csv, openalex_full_*.csv, pull_summary_*.json
"""
import sys, json, csv, time, urllib.parse, urllib.request, urllib.error

sys.stdout.reconfigure(encoding="utf-8")
TIMECODE = "20260802_211715"
OUTDIR = r"C:\Users\wh850\Research\assets\32_행태_사운드스케이프 리뷰 논문_(LUP)\_claude"

# ---------- 공통 fetch (429/503 백오프) ----------
def get(url: str, tries: int = 6) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": "paper32-pilot/0.3"})
    for i in range(tries):
        try:
            with urllib.request.urlopen(req, timeout=90) as r:
                return r.read()
        except urllib.error.HTTPError as e:
            if e.code in (429, 500, 502, 503) and i < tries - 1:
                time.sleep(4 * (i + 1))
                continue
            raise
        except urllib.error.URLError:
            if i < tries - 1:
                time.sleep(4 * (i + 1))
                continue
            raise

# ---------- 1) PubMed ----------
PUBMED_TERM = (
    '(soundscape*[tiab] OR "acoustic environment"[tiab] OR "sound environment"[tiab] '
    'OR "acoustic comfort"[tiab] OR "traffic noise"[tiab] OR "environmental noise"[tiab] '
    'OR "urban noise"[tiab] OR "aircraft noise"[tiab] OR "natural sound"[tiab] '
    'OR "natural sounds"[tiab] OR birdsong*[tiab] OR "bird song"[tiab] OR "bird songs"[tiab] '
    'OR "water sound"[tiab] OR "water sounds"[tiab] OR "background music"[tiab] '
    'OR "added sound"[tiab] OR "added sounds"[tiab]) '
    'AND (behavio*[tiab] OR "walking speed"[tiab] OR "pedestrian movement"[tiab] '
    'OR "route choice"[tiab] OR wayfinding[tiab] OR "dwell time"[tiab] OR linger*[tiab] '
    'OR "time spent"[tiab] OR sitting[tiab] OR "space use"[tiab] OR "park use"[tiab] '
    'OR visitation[tiab] OR "physical activity"[tiab] OR "social interaction"[tiab] '
    'OR "social interactions"[tiab] OR prosocial[tiab] OR antisocial[tiab] '
    'OR "anti-social"[tiab] OR avoidance[tiab] OR "crowd behavior"[tiab] '
    'OR "crowd behaviour"[tiab]) '
    'AND (urban[tiab] OR city[tiab] OR cities[tiab] OR park*[tiab] OR street*[tiab] '
    'OR "public space"[tiab] OR "public spaces"[tiab] OR "open space"[tiab] '
    'OR "open spaces"[tiab] OR square*[tiab] OR plaza*[tiab] OR waterfront*[tiab] '
    'OR "green space"[tiab] OR "green spaces"[tiab] OR greenspace*[tiab] '
    'OR "recreational area"[tiab] OR "recreational areas"[tiab] OR campus[tiab]) '
    'AND english[lang] AND journal article[pt]'
)

EUTILS = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils"

def pubmed_search():
    q = urllib.parse.quote(PUBMED_TERM)
    data = json.loads(get(f"{EUTILS}/esearch.fcgi?db=pubmed&retmax=10000&retmode=json&term={q}"))
    ids = data["esearchresult"]["idlist"]
    total = int(data["esearchresult"]["count"])
    print(f"[PubMed] 총 {total}건 (PMID {len(ids)}개 수신)")
    rows = []
    for i in range(0, len(ids), 200):
        batch = ids[i:i+200]
        time.sleep(0.5)  # 3req/s 한도 준수
        s = json.loads(get(f"{EUTILS}/esummary.fcgi?db=pubmed&retmode=json&id={','.join(batch)}"))
        for pid in batch:
            d = s["result"].get(pid, {})
            doi = ""
            for aid in d.get("articleids", []):
                if aid.get("idtype") == "doi":
                    doi = aid.get("value", "").lower()
            rows.append({
                "pmid": pid,
                "doi": doi,
                "title": d.get("title", ""),
                "year": (d.get("pubdate", "") or "")[:4],
                "journal": d.get("fulljournalname", ""),
            })
    path = f"{OUTDIR}\\pubmed_results_{TIMECODE}.csv"
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["pmid", "doi", "title", "year", "journal"])
        w.writeheader(); w.writerows(rows)
    print(f"[PubMed] CSV 저장: {path}")
    return total, rows

# ---------- 2) OpenAlex 전량 (cursor) ----------
from openalex_pilot_search import QUERY, FILTERS

def openalex_full():
    rows, cursor, page = [], "*", 0
    while cursor:
        page += 1
        params = {
            "filter": f"title_and_abstract.search:{QUERY},{FILTERS}",
            "per-page": 200, "cursor": cursor,
            "select": "id,doi,title,publication_year,primary_location,cited_by_count",
        }
        data = json.loads(get("https://api.openalex.org/works?" + urllib.parse.urlencode(params)))
        for w in data["results"]:
            src = (w.get("primary_location") or {}).get("source") or {}
            rows.append({
                "openalex_id": (w.get("id") or "").replace("https://openalex.org/", ""),
                "doi": (w.get("doi") or "").replace("https://doi.org/", "").lower(),
                "title": w.get("title") or "",
                "year": w.get("publication_year"),
                "journal": src.get("display_name") or "",
                "cited_by": w.get("cited_by_count"),
            })
        cursor = data["meta"].get("next_cursor")
        print(f"[OpenAlex] page {page} 누적 {len(rows)}건")
        time.sleep(1.2)
        if not data["results"]:
            break
    path = f"{OUTDIR}\\openalex_full_{TIMECODE}.csv"
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["openalex_id", "doi", "title", "year", "journal", "cited_by"])
        w.writeheader(); w.writerows(rows)
    print(f"[OpenAlex] 전량 {len(rows)}건 CSV 저장: {path}")
    return rows

# ---------- 3) 중복 통계 ----------
def main():
    pm_total, pm_rows = pubmed_search()
    oa_rows = openalex_full()
    pm_dois = {r["doi"] for r in pm_rows if r["doi"]}
    oa_dois = {r["doi"] for r in oa_rows if r["doi"]}
    inter = pm_dois & oa_dois
    union = len(pm_dois | oa_dois) + sum(1 for r in pm_rows if not r["doi"]) + sum(1 for r in oa_rows if not r["doi"])
    print(f"\n[중복 통계] PubMed DOI {len(pm_dois)} · OpenAlex DOI {len(oa_dois)} · 교집합 {len(inter)}")
    print(f"[병합 추정] 합집합(대략) ≈ {union}건 — PubMed 고유분 {len(pm_dois - oa_dois)}건이 OpenAlex 풀에 추가됨")
    summary = {
        "timecode": TIMECODE,
        "pubmed_total": pm_total,
        "pubmed_term": PUBMED_TERM,
        "openalex_total": len(oa_rows),
        "doi_intersection": len(inter),
        "pubmed_unique_dois": len(pm_dois - oa_dois),
    }
    with open(f"{OUTDIR}\\pull_summary_{TIMECODE}.json", "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    print(f"[요약] pull_summary_{TIMECODE}.json 저장")

if __name__ == "__main__":
    main()
