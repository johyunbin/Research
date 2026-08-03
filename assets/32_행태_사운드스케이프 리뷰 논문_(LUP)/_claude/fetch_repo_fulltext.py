# -*- coding: utf-8 -*-
"""
Paper32 — 전문 수집 3단계: OpenAlex 전체 OA locations 중 '리포지터리 사본' 노리기
best_oa_location(주로 출판사 CDN=봇월)이 아니라 locations[] 전체를 훑어
기관 리포지터리/프리프린트(EPrints·DSpace·HAL·DiVA·arXiv 등)를 우선 시도.
출력: fulltext/pdf/ID_XXXX.pdf + fulltext/repo_log_*.csv
"""
import sys, csv, os, json, time, glob, urllib.parse, urllib.request, urllib.error
from urllib.parse import urlparse

sys.stdout.reconfigure(encoding="utf-8")
TIMECODE = "20260802_221808"
BASE = r"C:\Users\wh850\Research\assets\32_행태_사운드스케이프 리뷰 논문_(LUP)\_claude"
FT = os.path.join(BASE, "fulltext")
H = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36",
    "Accept": "application/pdf,text/html;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}
# 출판사 CDN(봇월) — 스킵. 그 외(리포지터리)를 우선.
PUBLISHER_HOSTS = (
    "sciencedirect.com", "els-cdn.com", "springer.com", "nature.com", "wiley.com",
    "tandfonline.com", "sagepub.com", "oup.com", "academic.oup", "aip.org",
    "degruyter.com", "iospress.com", "cell.com", "mdpi.com", "frontiersin.org",
    "biomedcentral.com", "researchgate.net", "semanticscholar.org", "blob.core.windows.net",
)


def get(url, timeout=45, tries=2):
    req = urllib.request.Request(url, headers=H)
    last = None
    for i in range(tries):
        try:
            return urllib.request.urlopen(req, timeout=timeout).read()
        except Exception as e:
            last = e
            if i < tries - 1:
                time.sleep(2)
    raise last


def main():
    have = {int(os.path.basename(p)[3:7]) for p in glob.glob(os.path.join(FT, "pdf", "ID_*.pdf"))}
    with open(os.path.join(FT, "fulltext_status_20260802.csv"), encoding="utf-8-sig") as f:
        rows = [r for r in csv.DictReader(f) if int(r["no"]) not in have and r["doi"]]
    print(f"[1] 미확보 {len(rows)}건 (DOI 보유)")

    # OpenAlex: locations 전체 조회
    loc_by_doi = {}
    for i in range(0, len(rows), 50):
        batch = rows[i:i+50]
        filt = "doi:" + "|".join(r["doi"] for r in batch)
        url = ("https://api.openalex.org/works?filter=" + urllib.parse.quote(filt)
               + "&per-page=50&select=doi,locations")
        try:
            data = json.loads(get(url, timeout=60).decode("utf-8"))
        except Exception as e:
            print(f"    배치 {i//50+1} 실패: {e}"); continue
        for w in data.get("results", []):
            d = (w.get("doi") or "").replace("https://doi.org/", "").lower()
            cands = []
            for loc in (w.get("locations") or []):
                if not loc.get("is_oa"):
                    continue
                for u in (loc.get("pdf_url"), loc.get("landing_page_url")):
                    if u and not any(h in u for h in PUBLISHER_HOSTS):
                        cands.append(u)
            loc_by_doi[d] = cands
        print(f"    배치 {i//50+1} 조회 완료")
        time.sleep(1.2)

    n_cand = sum(1 for r in rows if loc_by_doi.get(r["doi"]))
    print(f"[2] 리포지터리 후보 있는 레코드 {n_cand}건")

    ok, log = 0, []
    for idx, r in enumerate(rows, 1):
        cands = loc_by_doi.get(r["doi"], [])
        status, used = "NO_REPO_CANDIDATE" if not cands else "", ""
        for u in cands[:4]:
            try:
                data = get(u)
                if data[:5] == b"%PDF-" and len(data) > 20000:
                    dest = os.path.join(FT, "pdf", f"ID_{int(r['no']):04d}.pdf")
                    with open(dest, "wb") as f:
                        f.write(data)
                    status, used, ok = "OK_REPO", u, ok + 1
                    break
                else:
                    status = "NOT_PDF"
            except Exception as e:
                status = f"ERR:{type(e).__name__}"
            time.sleep(1.5)
        log.append({"no": r["no"], "verdict": r["verdict"], "doi": r["doi"],
                    "status": status, "url": used, "n_cand": len(cands),
                    "journal": r["journal"][:60]})
        if idx % 20 == 0:
            print(f"    진행 {idx}/{len(rows)} (성공 {ok})")

    with open(os.path.join(FT, f"repo_log_{TIMECODE}.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["no", "verdict", "doi", "status", "url", "n_cand", "journal"])
        w.writeheader(); w.writerows(log)
    from collections import Counter
    print(f"\n[완료] 리포지터리 회수 {ok}건 → 누적 {len(have)+ok}/189")
    print("상태:", dict(Counter(l["status"] for l in log)))


if __name__ == "__main__":
    main()
