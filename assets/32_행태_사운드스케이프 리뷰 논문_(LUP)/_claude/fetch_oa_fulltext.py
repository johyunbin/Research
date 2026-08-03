# -*- coding: utf-8 -*-
"""
Paper32 — 전문(full-text) 1단계: 오픈액세스 PDF 자동 수집
대상: include+borderline 189건. OpenAlex에서 OA PDF URL 조회(50건 배치) 후
정중한 속도(2.5s 간격)로 다운로드. %PDF 매직바이트 검증.
출력: fulltext/pdf/ID_<no>.pdf + fulltext/oa_log_*.csv
"""
import sys, csv, os, re, json, time, urllib.parse, urllib.request, urllib.error

sys.stdout.reconfigure(encoding="utf-8")
TIMECODE = "20260802_221808"
BASE = r"C:\Users\wh850\Research\assets\32_행태_사운드스케이프 리뷰 논문_(LUP)\_claude"
SCR = os.path.join(BASE, "screening")
FT = os.path.join(BASE, "fulltext")
os.makedirs(os.path.join(FT, "pdf"), exist_ok=True)
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) paper32-SR-fulltext/0.1 (academic systematic review)"


def get(url, timeout=90, tries=4, accept=None):
    req = urllib.request.Request(url, headers={"User-Agent": UA, **({"Accept": accept} if accept else {})})
    for i in range(tries):
        try:
            return urllib.request.urlopen(req, timeout=timeout)
        except urllib.error.HTTPError as e:
            if e.code in (429, 500, 502, 503) and i < tries - 1:
                time.sleep(5 * (i + 1)); continue
            raise
        except urllib.error.URLError:
            if i < tries - 1:
                time.sleep(5 * (i + 1)); continue
            raise


def main():
    # 1) 대상 로드
    targets = []
    for name in (f"include_{TIMECODE}.csv", f"borderline_{TIMECODE}.csv"):
        with open(os.path.join(SCR, name), encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                targets.append({"no": r["no"], "doi": (r["doi"] or "").strip().lower(),
                                "title": r["title"], "verdict": "include" in name and "INCLUDE" or "BORDERLINE"})
    print(f"[1] 대상 {len(targets)}건 (DOI 있는 것 {sum(1 for t in targets if t['doi'])}건)")

    # 2) OpenAlex OA URL 조회 (50건 파이프 배치)
    oa = {}
    with_doi = [t for t in targets if t["doi"]]
    for i in range(0, len(with_doi), 50):
        batch = with_doi[i:i+50]
        filt = "doi:" + "|".join(t["doi"] for t in batch)
        url = ("https://api.openalex.org/works?filter=" + urllib.parse.quote(filt)
               + "&per-page=50&select=doi,open_access,best_oa_location")
        try:
            data = json.loads(get(url).read().decode("utf-8"))
        except Exception as e:
            print(f"    배치 {i//50+1} 조회 실패: {e}"); continue
        for w in data.get("results", []):
            d = (w.get("doi") or "").replace("https://doi.org/", "").lower()
            loc = w.get("best_oa_location") or {}
            oa[d] = {
                "is_oa": (w.get("open_access") or {}).get("is_oa", False),
                "pdf_url": loc.get("pdf_url") or "",
                "landing": loc.get("landing_page_url") or "",
            }
        print(f"    배치 {i//50+1}: {len(data.get('results', []))}건 조회")
        time.sleep(1.5)

    n_oa = sum(1 for t in with_doi if oa.get(t["doi"], {}).get("is_oa"))
    n_pdfurl = sum(1 for t in with_doi if oa.get(t["doi"], {}).get("pdf_url"))
    print(f"[2] OA {n_oa}건 · 직접 PDF URL {n_pdfurl}건")

    # 3) 다운로드 (정중한 속도)
    log, ok = [], 0
    for idx, t in enumerate(targets, 1):
        rec = oa.get(t["doi"], {})
        pdf_url = rec.get("pdf_url", "")
        status, path = "NO_OA_PDF_URL" if not pdf_url else "", ""
        if pdf_url:
            dest = os.path.join(FT, "pdf", f"ID_{int(t['no']):04d}.pdf")
            if os.path.exists(dest) and os.path.getsize(dest) > 10000:
                status, path, ok = "ALREADY", dest, ok + 1
            else:
                try:
                    resp = get(pdf_url, accept="application/pdf")
                    data = resp.read()
                    if data[:5] == b"%PDF-":
                        with open(dest, "wb") as f:
                            f.write(data)
                        status, path, ok = "OK", dest, ok + 1
                    else:
                        status = "NOT_PDF_CONTENT"
                except Exception as e:
                    status = f"ERR:{type(e).__name__}"
                time.sleep(2.5)
        if not rec.get("is_oa") and not pdf_url:
            status = "PAYWALLED" if t["doi"] else "NO_DOI"
        log.append({"no": t["no"], "verdict": t["verdict"], "doi": t["doi"],
                    "status": status, "pdf_url": pdf_url, "title": t["title"][:100]})
        if idx % 20 == 0:
            print(f"    진행 {idx}/{len(targets)} (성공 {ok})")

    # 4) 로그 저장
    logpath = os.path.join(FT, f"oa_log_{TIMECODE}.csv")
    with open(logpath, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["no", "verdict", "doi", "status", "pdf_url", "title"])
        w.writeheader(); w.writerows(log)
    from collections import Counter
    dist = Counter(l["status"] for l in log)
    print(f"\n[완료] PDF 확보 {ok}/{len(targets)}")
    print("상태 분포:", dict(dist))
    print("로그:", logpath)


if __name__ == "__main__":
    main()
