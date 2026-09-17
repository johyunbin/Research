# -*- coding: utf-8 -*-
"""
Paper32 — 데이터베이스 경로 전문 미확보 89건의 합법적 무료 원문 확보 (2026-09-17 사용자 지시 "추가해야지")

대상: fulltext/unretrieved89_oa_check_20260917.csv (DB 경로 전문 확보 대상 189건 중 미확보 89건)
출처(합법적 오픈액세스만): OpenAlex OA locations · Semantic Scholar openAccessPdf · Europe PMC(PMCID 있는 경우)
원칙: 봇 확인·403·429가 나오면 우회하지 않고 실패로 기록한다 → 사용자 수동 확보 목록으로 넘긴다.
검증: PDF 시그니처 · 3쪽 이상 · 앞 2쪽 본문에 제목 핵심어 60% 이상 포함(참고문헌 전용 PDF·엉뚱한 파일 배제)
출력: _claude/db_oa_pdf/<no>_<year>_<제목>.pdf(.xml) · fulltext/db_oa_retrieval_<타임코드>.csv
      01_논문작업/미확보_원문_목록_<타임코드>.md (사용자 수동 확보용, DOI 링크 포함)
"""
import csv, json, os, re, sys, time, unicodedata, urllib.parse, urllib.request
from datetime import datetime, timezone, timedelta

import fitz  # PyMuPDF

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(BASE, "fulltext", "unretrieved89_oa_check_20260917.csv")
OUTDIR = os.path.join(BASE, "db_oa_pdf")
UA = "paper32-systematic-review/1.0 (academic literature review; open-access retrieval)"
STAMP = datetime.now(timezone(timedelta(hours=9))).strftime("%Y%m%d_%H%M%S")
STOP = {"with", "from", "that", "this", "their", "into", "between", "among", "based", "study", "effects", "effect"}


def get(url, accept=None, timeout=40):
    req = urllib.request.Request(url, headers={"User-Agent": UA, **({"Accept": accept} if accept else {})})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read(), r.headers.get("Content-Type", ""), r.geturl()


def get_json(url):
    try:
        data, _, _ = get(url, "application/json")
        return json.loads(data)
    except Exception:
        return None
    finally:
        time.sleep(0.6)


def slug(s, n=60):
    s = unicodedata.normalize("NFKC", s)
    s = re.sub(r'[\\/:*?"<>|]+', " ", s)
    return re.sub(r"\s+", " ", s).strip()[:n].rstrip(" .")


def title_words(t):
    return {w for w in re.findall(r"[a-z]{4,}", unicodedata.normalize("NFKC", t).lower()) if w not in STOP}


def check_pdf(blob, title):
    if not blob.startswith(b"%PDF"):
        return False, "not_pdf", 0
    try:
        doc = fitz.open(stream=blob, filetype="pdf")
    except Exception:
        return False, "pdf_unreadable", 0
    pages = doc.page_count
    text = " ".join(doc[i].get_text() for i in range(min(2, pages))).lower()
    text = re.sub(r"-\s*\n\s*", "", text)
    words = title_words(title)
    hit = sum(1 for w in words if w in text) / max(1, len(words))
    if pages < 3:
        return False, f"too_short({pages}p)", pages
    if hit < 0.6:
        return False, f"title_mismatch({hit:.0%})", pages
    return True, f"ok(title {hit:.0%})", pages


def candidates(rec):
    doi = rec["doi"].strip()
    cands = []
    if not doi:
        return cands
    w = get_json("https://api.openalex.org/works/doi:" + urllib.parse.quote(doi))
    if w:
        for loc in w.get("locations") or []:
            if loc.get("is_oa") and loc.get("pdf_url"):
                cands.append(("openalex", loc["pdf_url"]))
        oa_url = (w.get("open_access") or {}).get("oa_url")
        if oa_url and oa_url.lower().endswith(".pdf"):
            cands.append(("openalex_oa_url", oa_url))
    s2 = get_json("https://api.semanticscholar.org/graph/v1/paper/DOI:" + urllib.parse.quote(doi)
                  + "?fields=openAccessPdf,externalIds")
    pmcid = ""
    if s2:
        u = (s2.get("openAccessPdf") or {}).get("url")
        if u:
            cands.append(("semanticscholar", u))
        pmcid = ((s2.get("externalIds") or {}).get("PubMedCentral") or "")
    if not pmcid:
        ep = get_json("https://www.ebi.ac.uk/europepmc/webservices/rest/search?format=json&resultType=lite&query="
                      + urllib.parse.quote(f'DOI:"{doi}"'))
        for r in ((ep or {}).get("resultList") or {}).get("result") or []:
            if r.get("pmcid"):
                pmcid = r["pmcid"]
                break
    if pmcid:
        pmcid = pmcid if pmcid.upper().startswith("PMC") else "PMC" + pmcid
        cands.append(("europepmc_pdf", f"https://europepmc.org/articles/{pmcid}?pdf=render"))
        cands.append(("europepmc_xml", f"https://www.ebi.ac.uk/europepmc/webservices/rest/{pmcid}/fullTextXML"))
    seen, out = set(), []
    for src, u in cands:
        if u not in seen:
            seen.add(u)
            out.append((src, u))
    return out


def main():
    os.makedirs(OUTDIR, exist_ok=True)
    rows = list(csv.DictReader(open(SRC, encoding="utf-8-sig")))
    log = []
    for i, rec in enumerate(rows, 1):
        no, title = rec["no"], rec["title"]
        base = os.path.join(OUTDIR, f"{int(no):04d}_{rec['year']}_{slug(title)}")
        result = {"no": no, "verdict": rec["verdict"], "doi": rec["doi"], "year": rec["year"], "journal": rec["journal"],
                  "title": title, "oa_status": rec["oa_status"], "status": "", "source": "", "url": "", "file": "",
                  "pages": "", "tried": ""}
        tried = []
        for src, url in candidates(rec):
            try:
                blob, ctype, final = get(url, timeout=60)
            except urllib.error.HTTPError as e:
                tried.append(f"{src}:HTTP{e.code}")
                time.sleep(1)
                continue
            except Exception as e:
                tried.append(f"{src}:{type(e).__name__}")
                time.sleep(1)
                continue
            time.sleep(1)
            if src == "europepmc_xml":
                if b"<body" in blob and len(blob) > 20000:
                    path = base + ".xml"
                    open(path, "wb").write(blob)
                    result.update(status="OK_XML", source=src, url=url, file=os.path.basename(path))
                    break
                tried.append(f"{src}:no_body")
                continue
            ok, why, pages = check_pdf(blob, title)
            if ok:
                path = base + ".pdf"
                open(path, "wb").write(blob)
                result.update(status="OK_PDF", source=src, url=url, file=os.path.basename(path), pages=pages)
                break
            tried.append(f"{src}:{why}")
        if not result["status"]:
            result["status"] = "NOT_RETRIEVED"
        result["tried"] = " | ".join(tried)
        log.append(result)
        print(f"[{i:2d}/{len(rows)}] {no:>5} {result['status']:13} {result['source']:16} {title[:60]}")

    logpath = os.path.join(BASE, "fulltext", f"db_oa_retrieval_{STAMP}.csv")
    with open(logpath, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(log[0].keys()))
        w.writeheader()
        w.writerows(log)

    # 사용자 수동 확보 목록
    miss = [r for r in log if r["status"] == "NOT_RETRIEVED"]
    miss.sort(key=lambda r: (r["oa_status"] in ("closed", ""), r["year"]))
    L = [f"# 전문 미확보 논문 목록 ({STAMP[:8]})", "",
         f"데이터베이스 검색 경로에서 전문 확보 대상이었으나 아직 전문이 없는 {len(miss)}건입니다. "
         f"자동으로 확보한 {len(log) - len(miss)}건은 목록에서 뺐습니다.", "",
         "- **OA 표시**: OpenAlex 기준 무료 원문 여부. gold·hybrid·bronze·green·diamond = 무료 원문이 있다고 표시됐지만 자동 내려받기가 막힌 건, closed = 구독 전용, 빈칸 = DOI 없음",
         "- 받은 PDF는 `수집논문_PDF/추가확보/` 폴더에 넣어 주세요. 파일 이름은 상관없습니다(제목으로 대조합니다).", "",
         "| # | ID | 연도 | 학술지 | 제목 | DOI | OA 표시 |", "|---|---|---|---|---|---|---|"]
    for k, r in enumerate(miss, 1):
        doi = r["doi"].strip()
        link = f"[{doi}](https://doi.org/{doi})" if doi else "없음"
        t = r["title"].replace("|", "/")
        L.append(f"| {k} | {r['no']} | {r['year']} | {r['journal']} | {t} | {link} | {r['oa_status'] or '-'} |")
    mdpath = os.path.join(os.path.dirname(BASE), "01_논문작업", f"미확보_원문_목록_{STAMP}.md")
    open(mdpath, "w", encoding="utf-8").write("\n".join(L) + "\n")
    from collections import Counter
    print("\n상태:", Counter(r["status"] for r in log))
    print("출처:", Counter(r["source"] for r in log if r["source"]))
    print("로그:", logpath)
    print("목록:", mdpath)


if __name__ == "__main__":
    main()
