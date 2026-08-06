# -*- coding: utf-8 -*-
"""
Paper32 — OA 32편 내려받기 (세션 쿠키 확보 후 PDF 요청)
이전 실패 원인: 기사 페이지를 거치지 않고 PDF URL을 바로 때려 쿠키·Referer가 없었음.
여기서는 ①기사 페이지 GET(쿠키 적립) → ②Referer 붙여 PDF GET 순으로 간다.
출력: ct_pdf/*.pdf · fulltext/oa_fetch_log.csv
"""
import sys, os, csv, time, http.cookiejar, urllib.request, urllib.error, re

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
PDF = os.path.join(BASE, "ct_pdf")
os.makedirs(PDF, exist_ok=True)

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
      "Chrome/125.0.0.0 Safari/537.36")
HDR = {
    "User-Agent": UA,
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9,ko;q=0.8",
    "Sec-Ch-Ua": '"Chromium";v="125", "Not.A/Brand";v="24"',
    "Sec-Ch-Ua-Mobile": "?0",
    "Sec-Ch-Ua-Platform": '"Windows"',
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "none",
    "Sec-Fetch-User": "?1",
    "Upgrade-Insecure-Requests": "1",
}


def opener():
    cj = http.cookiejar.CookieJar()
    return urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))


def safe(s, n=88):
    s = re.sub(r"[^\w\s\-.,()]", "", s, flags=re.UNICODE)
    return re.sub(r"\s+", " ", s).strip()[:n]


def warm(op, url):
    """기사 페이지를 먼저 열어 쿠키를 적립"""
    try:
        op.open(urllib.request.Request(url, headers=HDR), timeout=45).read(200000)
        return True
    except Exception:
        return False


def grab(op, url, referer, path):
    h = dict(HDR)
    h.update({"Accept": "application/pdf,text/html;q=0.9,*/*;q=0.8",
              "Referer": referer, "Sec-Fetch-Dest": "document",
              "Sec-Fetch-Site": "same-origin"})
    try:
        with op.open(urllib.request.Request(url, headers=h), timeout=120) as r:
            head = r.read(1024)
            if not head.startswith(b"%PDF"):
                return "not-pdf"
            with open(path, "wb") as f:
                f.write(head)
                while True:
                    b = r.read(65536)
                    if not b:
                        break
                    f.write(b)
        return "ok" if os.path.getsize(path) > 20000 else "too-small"
    except urllib.error.HTTPError as e:
        return f"http{e.code}"
    except Exception as e:
        return type(e).__name__


def main():
    rows = list(csv.DictReader(open(os.path.join(FT, "oa_targets_resolved.csv"),
                                    encoding="utf-8-sig")))
    log, ok = [], 0
    for r in rows:
        rec = int(r["rec"])
        fn = f"CT{rec:04d}_{r['year']}_{safe(r['title'])}.pdf"
        path = os.path.join(PDF, fn)
        if os.path.exists(path) and os.path.getsize(path) > 20000:
            log.append({**r, "result": "already", "file": fn}); ok += 1
            print(f"  {rec:>4} already"); continue

        op = opener()
        art = r["final_url"] or r["landing"]
        warmed = warm(op, art)
        time.sleep(1.0)

        res = "no-candidate"
        for key in ("pdf_try1", "pdf_try2", "pdf_try3"):
            u = (r.get(key) or "").strip()
            if not u:
                continue
            res = grab(op, u, art, path)
            if res == "ok":
                ok += 1
                break
            if os.path.exists(path):
                os.remove(path)
            time.sleep(1.2)
        log.append({**r, "result": res, "file": fn if res == "ok" else "", "warmed": warmed})
        print(f"  {rec:>4} warm={'Y' if warmed else 'N'}  {res}")
        time.sleep(1.5)

    keys = ["rec", "screen", "oa_status", "year", "journal", "doi", "final_url",
            "pdf_try1", "result", "file", "title"]
    with open(os.path.join(FT, "oa_fetch_log.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=keys, extrasaction="ignore")
        w.writeheader(); w.writerows(log)
    from collections import Counter
    print(f"\n[완료] 확보 {ok}/{len(rows)} · {dict(Counter(x['result'] for x in log))}")


if __name__ == "__main__":
    main()
