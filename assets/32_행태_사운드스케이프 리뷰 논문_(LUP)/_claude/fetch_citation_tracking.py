# -*- coding: utf-8 -*-
"""
Paper32 — 인용 추적(backward + forward), 등록 프로토콜 "Other search strategies" 이행
OpenAlex API로 포함 84편의 ①참고문헌(backward) ②피인용(forward)을 수집,
기존 스크리닝 풀(1,316) + 코퍼스와 대조해 '신규 후보'만 추출.
출력: fulltext/citation_tracking.csv (신규 후보) + citation_tracking_log.json
"""
import sys, csv, os, json, time, urllib.parse, urllib.request, urllib.error, re, unicodedata
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
UA = {"User-Agent": "paper32-SR-citation-tracking/0.1 (academic systematic review)"}


def get(url, tries=5):
    req = urllib.request.Request(url, headers=UA)
    for i in range(tries):
        try:
            with urllib.request.urlopen(req, timeout=60) as r:
                return json.loads(r.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            if e.code in (429, 500, 502, 503) and i < tries - 1:
                time.sleep(4 * (i + 1)); continue
            raise
        except urllib.error.URLError:
            if i < tries - 1:
                time.sleep(4 * (i + 1)); continue
            raise


def norm_doi(d):
    return re.sub(r"^https?://(dx\.)?doi\.org/", "", (d or "").strip().lower())


def norm_title(t):
    t = unicodedata.normalize("NFKD", (t or "").lower())
    return re.sub(r"[^a-z0-9]", "", t)


def main():
    # 이미 본 것: 스크리닝 풀 1,316 + 코퍼스
    seen_doi, seen_title = set(), set()
    with open(os.path.join(BASE, "merged_pool_20260802_211715.csv"), encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            if r["doi"]: seen_doi.add(norm_doi(r["doi"]))
            seen_title.add(norm_title(r["title"]))
    print(f"[0] 기존 풀 {len(seen_doi)} DOI / {len(seen_title)} 제목")

    # 포함 84편의 DOI → OpenAlex ID
    inc = []
    with open(os.path.join(FT, "ft_verdicts_v2.csv"), encoding="utf-8-sig") as f:
        vs = {int(r["no"]): r["final_verdict"] for r in csv.DictReader(f)}
    with open(os.path.join(FT, "fulltext_status_20260803.csv"), encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            n = int(r["no"])
            if vs.get(n) in ("FINAL_INCLUDE", "SENS_ONLY") and r["doi"]:
                inc.append((n, norm_doi(r["doi"])))
    print(f"[1] 인용추적 시드 {len(inc)}편")

    oa_ids, meta = {}, {}
    for i in range(0, len(inc), 40):
        batch = inc[i:i+40]
        filt = "doi:" + "|".join(d for _, d in batch)
        url = ("https://api.openalex.org/works?filter=" + urllib.parse.quote(filt) +
               "&per-page=40&select=id,doi,title,referenced_works,cited_by_api_url,cited_by_count")
        try:
            data = get(url)
        except Exception as e:
            print(f"    시드 배치 {i//40+1} 실패: {e}"); continue
        for w in data.get("results", []):
            d = norm_doi(w.get("doi"))
            oa_ids[d] = w["id"].replace("https://openalex.org/", "")
            meta[d] = w
        print(f"    시드 배치 {i//40+1}: {len(data.get('results', []))}건")
        time.sleep(1.2)
    print(f"[2] OpenAlex 매핑 {len(oa_ids)}편")

    # ── backward: referenced_works 수집
    ref_ids = Counter()
    for d, w in meta.items():
        for rid in (w.get("referenced_works") or []):
            ref_ids[rid.replace("https://openalex.org/", "")] += 1
    print(f"[3] backward 참고문헌 고유 {len(ref_ids)}건 (2회 이상 인용 {sum(1 for v in ref_ids.values() if v>=2)})")

    # ── forward: 각 시드를 인용한 논문
    fwd_ids = Counter()
    seed_ids = list(oa_ids.values())
    for i in range(0, len(seed_ids), 25):
        chunk = seed_ids[i:i+25]
        url = ("https://api.openalex.org/works?filter=" + urllib.parse.quote("cites:" + "|".join(chunk)) +
               "&per-page=200&select=id&cursor=*")
        cursor = "*"
        got = 0
        while cursor and got < 2000:
            u = ("https://api.openalex.org/works?filter=" + urllib.parse.quote("cites:" + "|".join(chunk)) +
                 f"&per-page=200&cursor={cursor}&select=id")
            try:
                data = get(u)
            except Exception as e:
                print(f"    forward chunk {i//25+1} 실패: {e}"); break
            for w in data.get("results", []):
                fwd_ids[w["id"].replace("https://openalex.org/", "")] += 1
            got += len(data.get("results", []))
            cursor = data["meta"].get("next_cursor")
            if not data.get("results"): break
            time.sleep(0.8)
        print(f"    forward chunk {i//25+1}: {got}건")
    print(f"[4] forward 피인용 고유 {len(fwd_ids)}건")

    # ── 후보 필터: 자기 자신 제외 + 다빈도 우선(backward는 2회 이상, forward는 전부)
    seed_set = set(oa_ids.values())
    cand = {}
    for wid, c in ref_ids.items():
        if wid not in seed_set and c >= 2:
            cand[wid] = ("backward", c)
    for wid, c in fwd_ids.items():
        if wid not in seed_set:
            cand.setdefault(wid, ("forward", c))
    print(f"[5] 1차 후보 {len(cand)}건 (backward≥2회 + forward 전체)")

    # 메타 조회 → 기존 풀 대조 + 언어/유형 필터
    rows, ids = [], list(cand.keys())
    for i in range(0, len(ids), 50):
        chunk = ids[i:i+50]
        url = ("https://api.openalex.org/works?filter=" + urllib.parse.quote("openalex_id:" + "|".join(chunk)) +
               "&per-page=50&select=id,doi,title,publication_year,language,type,primary_location,cited_by_count")
        try:
            data = get(url)
        except Exception as e:
            print(f"    메타 배치 {i//50+1} 실패: {e}"); continue
        for w in data.get("results", []):
            wid = w["id"].replace("https://openalex.org/", "")
            d, t = norm_doi(w.get("doi")), norm_title(w.get("title"))
            if (d and d in seen_doi) or (t and t in seen_title):
                continue                      # 이미 스크리닝한 레코드
            if w.get("language") not in (None, "en"):
                continue                      # 영어만(등록 기준)
            if w.get("type") != "article":
                continue
            src = (w.get("primary_location") or {}).get("source") or {}
            if (src.get("type") or "") != "journal":
                continue
            way, cnt = cand[wid]
            rows.append({"openalex_id": wid, "route": way, "freq": cnt,
                         "doi": d, "year": w.get("publication_year"),
                         "journal": src.get("display_name", ""), "cited_by": w.get("cited_by_count"),
                         "title": w.get("title", "")})
        time.sleep(1.0)
        if (i // 50) % 5 == 0:
            print(f"    메타 {i}/{len(ids)} · 신규 {len(rows)}")

    rows.sort(key=lambda r: (r["route"], -(r["freq"] or 0), -(r["cited_by"] or 0)))
    out = os.path.join(FT, "citation_tracking.csv")
    with open(out, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["openalex_id", "route", "freq", "doi", "year",
                                          "journal", "cited_by", "title"])
        w.writeheader(); w.writerows(rows)
    json.dump({"seeds": len(inc), "mapped": len(oa_ids), "backward_unique": len(ref_ids),
               "forward_unique": len(fwd_ids), "candidates": len(cand), "new_after_dedup": len(rows)},
              open(os.path.join(FT, "citation_tracking_log.json"), "w"), indent=2)
    print(f"\n[완료] 스크리닝 대상 신규 후보 {len(rows)}건 → {out}")
    print("   route 분포:", dict(Counter(r["route"] for r in rows)))


if __name__ == "__main__":
    main()
