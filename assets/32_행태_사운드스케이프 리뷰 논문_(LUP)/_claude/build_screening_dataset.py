# -*- coding: utf-8 -*-
"""
Paper32 — 스크리닝 데이터셋 구축
merged_pool(1,316) + RIS 초록(WoS/Scopus, 연속행 처리) + PubMed efetch 초록 보강
출력: screening/screening_dataset_*.csv + 청크 10개(screening/chunks/chunk_NN.txt)
"""
import sys, csv, json, re, time, unicodedata, urllib.parse, urllib.request, urllib.error, os

sys.stdout.reconfigure(encoding="utf-8")
TIMECODE = "20260802_221808"
BASE = r"C:\Users\wh850\Research\assets\32_행태_사운드스케이프 리뷰 논문_(LUP)\_claude"
SCR = os.path.join(BASE, "screening")
os.makedirs(os.path.join(SCR, "chunks"), exist_ok=True)
os.makedirs(os.path.join(SCR, "results"), exist_ok=True)


def norm_doi(d):
    d = (d or "").strip().lower()
    return re.sub(r"^https?://(dx\.)?doi\.org/", "", d)


def norm_title(t):
    t = unicodedata.normalize("NFKD", (t or "").lower())
    return re.sub(r"[^a-z0-9]", "", t)


def parse_ris_full(path):
    """RIS 파서 — 연속행(태그 없는 줄)을 직전 필드에 이어붙임."""
    recs, cur, last_tag = [], {}, None
    tag_re = re.compile(r"^([A-Z][A-Z0-9])  - ?(.*)$")
    with open(path, encoding="utf-8-sig") as f:
        for raw in f:
            line = raw.rstrip("\n\r")
            m = tag_re.match(line)
            if m:
                tag, val = m.group(1), m.group(2).strip()
                last_tag = tag
                if tag == "TY":
                    cur = {}
                elif tag == "ER":
                    recs.append(cur); cur = {}; last_tag = None
                elif tag in ("TI", "T1"):
                    cur["title"] = cur.get("title", "") + (" " if "title" in cur else "") + val
                elif tag in ("DO", "DI"):
                    cur.setdefault("doi", val)
                elif tag == "AB":
                    cur["abstract"] = cur.get("abstract", "") + (" " if "abstract" in cur else "") + val
            else:
                text = line.strip()
                if text and last_tag in ("AB", "TI", "T1") and cur:
                    key = "abstract" if last_tag == "AB" else "title"
                    cur[key] = cur.get(key, "") + " " + text
    return recs


def fetch(url, tries=5):
    req = urllib.request.Request(url, headers={"User-Agent": "paper32-screening/0.1"})
    for i in range(tries):
        try:
            with urllib.request.urlopen(req, timeout=90) as r:
                return r.read()
        except (urllib.error.HTTPError, urllib.error.URLError) as e:
            code = getattr(e, "code", None)
            if i < tries - 1 and (code in (429, 500, 502, 503) or code is None):
                time.sleep(3 * (i + 1)); continue
            raise


def main():
    # 1) RIS 초록 맵
    ab_by_doi, ab_by_title = {}, {}
    for name in ["ris/wos_0001-1000_20260802.ris", "ris/wos_1001-1010_20260802.ris",
                 "ris/scopus_0001-0850_20260802.ris"]:
        for r in parse_ris_full(os.path.join(BASE, name)):
            ab = (r.get("abstract") or "").strip()
            if not ab:
                continue
            d, t = norm_doi(r.get("doi")), norm_title(r.get("title"))
            if d and (d not in ab_by_doi or len(ab) > len(ab_by_doi[d])):
                ab_by_doi[d] = ab
            if t and (t not in ab_by_title or len(ab) > len(ab_by_title[t])):
                ab_by_title[t] = ab
    print(f"[1] RIS 초록: DOI키 {len(ab_by_doi)} · 제목키 {len(ab_by_title)}")

    # 2) 풀 로드 + 초록 조인
    with open(os.path.join(BASE, "merged_pool_20260802_211715.csv"), encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    missing = []
    for r in rows:
        ab = ab_by_doi.get(norm_doi(r["doi"])) or ab_by_title.get(norm_title(r["title"])) or ""
        r["abstract"] = ab
        if not ab:
            missing.append(r)
    print(f"[2] 풀 {len(rows)}건 중 초록 미확보 {len(missing)}건 → PubMed efetch 시도")

    # 3) PubMed efetch 보강 (pubmed 소스 & DOI 있는 것)
    pm_missing = [r for r in missing if "pubmed" in r["sources"]]
    if pm_missing:
        with open(os.path.join(BASE, "pubmed_results_20260802_211715.csv"), encoding="utf-8-sig") as f:
            pmid_by_doi = {norm_doi(x["doi"]): x["pmid"] for x in csv.DictReader(f) if x.get("doi")}
        ids = [(r, pmid_by_doi.get(norm_doi(r["doi"]))) for r in pm_missing]
        ids = [(r, p) for r, p in ids if p]
        print(f"    efetch 대상 PMID {len(ids)}건")
        for i in range(0, len(ids), 100):
            batch = ids[i:i+100]
            url = ("https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi?db=pubmed"
                   f"&rettype=abstract&retmode=xml&id={','.join(p for _, p in batch)}")
            xml = fetch(url).decode("utf-8", "ignore")
            arts = re.findall(r"<PubmedArticle>.*?</PubmedArticle>", xml, re.S)
            for art in arts:
                pm = re.search(r"<PMID[^>]*>(\d+)</PMID>", art)
                abs_parts = re.findall(r"<AbstractText[^>]*>(.*?)</AbstractText>", art, re.S)
                if pm and abs_parts:
                    text = re.sub(r"<[^>]+>", " ", " ".join(abs_parts))
                    text = re.sub(r"\s+", " ", text).strip()
                    for r, p in batch:
                        if p == pm.group(1) and not r["abstract"]:
                            r["abstract"] = text
            time.sleep(0.5)
    still = sum(1 for r in rows if not r["abstract"])
    print(f"[3] efetch 후 초록 미확보 {still}건 (제목만으로 스크리닝·over-inclusive 플래그)")

    # 4) 데이터셋 저장
    out = os.path.join(SCR, f"screening_dataset_{TIMECODE}.csv")
    with open(out, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["no", "sources", "doi", "year", "journal", "title", "abstract"])
        for r in rows:
            w.writerow([r["no"], r["sources"], r["doi"], r["year"], r["journal"],
                        r["title"], r["abstract"][:2000]])
    print(f"[4] 데이터셋 저장: {out}")

    # 5) 청크 10개 (에이전트 투입용 — 압축 텍스트 포맷)
    n_chunks = 10
    per = (len(rows) + n_chunks - 1) // n_chunks
    for c in range(n_chunks):
        part = rows[c*per:(c+1)*per]
        path = os.path.join(SCR, "chunks", f"chunk_{c+1:02d}.txt")
        with open(path, "w", encoding="utf-8") as f:
            for r in part:
                ab = r["abstract"][:1400] if r["abstract"] else "(NO ABSTRACT — title-only)"
                f.write(f"### ID={r['no']} | {r['year']} | {r['journal'][:60]}\n"
                        f"TITLE: {r['title']}\nABSTRACT: {ab}\n\n")
        print(f"    chunk_{c+1:02d}.txt: {len(part)}건")
    print("[5] 청크 분할 완료")


if __name__ == "__main__":
    main()
