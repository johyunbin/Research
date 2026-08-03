# -*- coding: utf-8 -*-
"""
Paper32 — OpenAlex 본검색 드라이런 (기관 계정 불요)
3블록 불리언(A 노출 AND B 행태 AND C 세팅)을 title/abstract 검색으로 실행,
필터: 영어·저널 논문(프로시딩 제외) — 프로토콜 확정 기준(2026-08-02) 반영.
출력: 총 카운트, 연도 분포, 상위 후보 CSV, 요약 JSON.
"""
import sys, json, csv, time, urllib.parse, urllib.request

sys.stdout.reconfigure(encoding="utf-8")

BASE = "https://api.openalex.org/works"

BLOCK_A = ('(soundscape OR soundscapes OR "acoustic environment" OR "sound environment" '
           'OR "acoustic comfort" OR "traffic noise" OR "environmental noise" OR "urban noise" '
           'OR "aircraft noise" OR "natural sound" OR "natural sounds" OR birdsong '
           'OR "water sound" OR "water sounds" OR "background music" OR "added sound")')

BLOCK_B = ('(behavior OR behaviour OR behavioral OR behavioural OR "walking speed" '
           'OR "pedestrian movement" OR "route choice" OR wayfinding OR "dwell time" '
           'OR lingering OR "time spent" OR sitting OR "space use" OR "park use" '
           'OR visitation OR "physical activity" OR "social interaction" '
           'OR "social interactions" OR prosocial OR antisocial OR avoidance '
           'OR "crowd behavior" OR "crowd behaviour")')

BLOCK_C = ('(urban OR city OR cities OR park OR parks OR street OR streets '
           'OR "public space" OR "public spaces" OR "open space" OR "open spaces" '
           'OR plaza OR square OR waterfront OR greenspace OR "green space" '
           'OR "recreational area" OR "recreational areas" OR campus)')

QUERY = f"{BLOCK_A} AND {BLOCK_B} AND {BLOCK_C}"

# 프로토콜 확정 필터: 영어 · 저널 게재 논문(프로시딩/북챕터 제외)
FILTERS = "language:en,type:article,primary_location.source.type:journal"

TIMECODE = "20260802_205110"
OUTDIR = r"C:\Users\wh850\Research\assets\32_행태_사운드스케이프 리뷰 논문_(LUP)\_claude"


def fetch(params: dict) -> dict:
    url = BASE + "?" + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers={"User-Agent": "paper32-pilot/0.1"})
    with urllib.request.urlopen(req, timeout=60) as r:
        return json.loads(r.read().decode("utf-8"))


def main():
    filt = f"title_and_abstract.search:{QUERY},{FILTERS}"

    # 1) 총 카운트
    meta = fetch({"filter": filt, "per-page": 1})
    total = meta["meta"]["count"]
    print(f"[1] 필터 적용 총 카운트 (영어·저널논문): {total}")

    # 비교용: 필터 없는 원시 카운트
    raw = fetch({"filter": f"title_and_abstract.search:{QUERY}", "per-page": 1})
    print(f"    (비교) 무필터 원시 카운트: {raw['meta']['count']}")

    # 2) 연도 분포
    yr = fetch({"filter": filt, "group_by": "publication_year"})
    years = sorted(
        ((int(g["key"]), g["count"]) for g in yr["group_by"] if g["key"].isdigit()),
        key=lambda x: x[0],
    )
    recent = [(y, c) for y, c in years if y >= 2000]
    print(f"[2] 연도 분포 (2000+): {recent}")

    # 3) 상위 후보 400건 (relevance 순, 200×2페이지)
    rows = []
    for page in (1, 2):
        res = fetch({
            "filter": filt,
            "sort": "relevance_score:desc",
            "per-page": 200,
            "page": page,
            "select": "id,doi,title,publication_year,primary_location,cited_by_count,relevance_score",
        })
        for w in res["results"]:
            src = (w.get("primary_location") or {}).get("source") or {}
            rows.append({
                "openalex_id": (w.get("id") or "").replace("https://openalex.org/", ""),
                "doi": (w.get("doi") or "").replace("https://doi.org/", ""),
                "title": w.get("title") or "",
                "year": w.get("publication_year"),
                "journal": src.get("display_name") or "",
                "cited_by": w.get("cited_by_count"),
                "relevance": round(w.get("relevance_score") or 0, 2),
            })
        time.sleep(0.5)

    csv_path = f"{OUTDIR}\\openalex_pilot_{TIMECODE}.csv"
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        wcsv = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        wcsv.writeheader()
        wcsv.writerows(rows)
    print(f"[3] 상위 {len(rows)}건 CSV 저장: {csv_path}")

    # 4) 벤치마크 리콜 체크 — 파일럿 시드 논문이 검색식에 걸리는지
    benchmarks = [
        ("Franěk 보행속도", "traffic noise relaxation sounds pedestrian walking speed"),
        ("Aletta 공공공간 실험", "experimental study influence soundscapes people behaviour open public space"),
        ("Musikiosk", "soundtracking public space musikiosk soundscape intervention"),
        ("Meng 군중음악", "influence music behaviors crowd urban open public spaces"),
        ("Song 공원행태", "influence human behavioral characteristics soundscape perception urban parks"),
    ]
    print("[4] 벤치마크 리콜 체크 (제목 검색이 본검색식 AND 조건에 걸리는지):")
    hits = 0
    for label, t in benchmarks:
        chk = fetch({"filter": f"title_and_abstract.search:({t}) AND {QUERY},{FILTERS}", "per-page": 1})
        n = chk["meta"]["count"]
        hits += 1 if n > 0 else 0
        print(f"    - {label}: {'HIT' if n > 0 else 'MISS'} ({n})")
        time.sleep(0.3)
    print(f"    리콜: {hits}/{len(benchmarks)}")

    # 5) 요약 JSON
    summary = {
        "timecode": TIMECODE,
        "query_blocks": {"A": BLOCK_A, "B": BLOCK_B, "C": BLOCK_C},
        "filters": FILTERS,
        "total_filtered": total,
        "total_raw": raw["meta"]["count"],
        "year_distribution_2000plus": recent,
        "benchmark_recall": f"{hits}/{len(benchmarks)}",
        "csv": csv_path,
    }
    js_path = f"{OUTDIR}\\openalex_pilot_summary_{TIMECODE}.json"
    with open(js_path, "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    print(f"[5] 요약 저장: {js_path}")


if __name__ == "__main__":
    main()
