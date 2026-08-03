# -*- coding: utf-8 -*-
"""
Paper32 — 인용추적 전문 확보 2차 시도: OpenAlex 전체 locations + Unpaywall + PMC
1차(best_oa_location)에서 실패한 건만 재시도.
출력: ct_pdf/*.pdf · fulltext/ct_retrieval_status.csv 갱신
"""
import sys, os, csv, json, time, re, urllib.parse, urllib.request, urllib.error

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
PDF = os.path.join(BASE, "ct_pdf")
os.makedirs(PDF, exist_ok=True)
UAS = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
       "Chrome/125.0 Safari/537.36")
MAIL = "wh8502@naver.com"


def api(url, tries=3):
    req = urllib.request.Request(url, headers={"User-Agent": f"paper32-SR/0.1 (mailto:{MAIL})"})
    for i in range(tries):
        try:
            with urllib.request.urlopen(req, timeout=50) as r:
                return json.loads(r.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            if e.code in (429, 500, 502, 503) and i < tries - 1:
                time.sleep(4 * (i + 1)); continue
            return None
        except Exception:
            if i < tries - 1:
                time.sleep(4 * (i + 1)); continue
            return None


def safe(s, n=90):
    s = re.sub(r"[^\w\s\-.,()]", "", s, flags=re.UNICODE)
    return re.sub(r"\s+", " ", s).strip()[:n]


def download(url, path):
    req = urllib.request.Request(url, headers={
        "User-Agent": UAS, "Accept": "application/pdf,*/*", "Referer": url})
    try:
        with urllib.request.urlopen(req, timeout=90) as r:
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


def candidates(w, doi):
    """PDF 후보 URL을 우선순위대로"""
    urls = []
    for loc in (w.get("locations") or []):
        if loc.get("pdf_url"):
            urls.append(loc["pdf_url"])
        lp = loc.get("landing_page_url") or ""
        src = ((loc.get("source") or {}).get("display_name") or "").lower()
        if "pmc" in lp.lower() or "pubmed central" in src:
            m = re.search(r"(PMC\d+)", lp)
            if m:
                urls.append(f"https://www.ncbi.nlm.nih.gov/pmc/articles/{m.group(1)}/pdf/")
        if lp.startswith("https://www.mdpi.com/"):
            urls.append(lp.rstrip("/") + "/pdf")
    up = api(f"https://api.unpaywall.org/v2/{urllib.parse.quote(doi)}?email={MAIL}") if doi else None
    if up:
        for loc in (up.get("oa_locations") or []):
            if loc.get("url_for_pdf"):
                urls.append(loc["url_for_pdf"])
    seen, out = set(), []
    for u in urls:
        if u and u not in seen:
            seen.add(u); out.append(u)
    return out


def main():
    rows = list(csv.DictReader(open(os.path.join(FT, "ct_retrieval_status.csv"),
                                    encoding="utf-8-sig")))
    todo = [r for r in rows if r["result"] not in ("ok", "already")]
    print(f"[0] 재시도 대상 {len(todo)}/{len(rows)}")

    ids = [r["rec"] for r in todo]
    # openalex_id 재조회 (screen_final에서)
    fin = {r["rec"]: r for r in csv.DictReader(open(os.path.join(FT, "ct_screen_final.csv"),
                                                    encoding="utf-8-sig"))}
    oaids = [fin[r]["openalex_id"] for r in ids if fin.get(r, {}).get("openalex_id")]
    meta = {}
    for i in range(0, len(oaids), 40):
        d = api("https://api.openalex.org/works?filter=" +
                urllib.parse.quote("openalex_id:" + "|".join(oaids[i:i + 40])) +
                "&per-page=40&select=id,doi,locations")
        if d:
            for w in d.get("results", []):
                meta[w["id"].replace("https://openalex.org/", "")] = w
        time.sleep(1.0)
    print(f"[1] locations 메타 {len(meta)}")

    n_new = 0
    for r in todo:
        w = meta.get(fin.get(r["rec"], {}).get("openalex_id", ""), {})
        urls = candidates(w, r["doi"])
        fn = f"CT{int(r['rec']):04d}_{r.get('year','')}_{safe(r['title'])}.pdf"
        path = os.path.join(PDF, fn)
        res = r["result"]
        for u in urls[:4]:
            res = download(u, path)
            if res == "ok":
                n_new += 1
                r["file"] = fn
                r["pdf_url"] = u
                break
            if os.path.exists(path):
                os.remove(path)
            time.sleep(1.2)
        if res != r["result"]:
            r["result"] = res
        if urls:
            print(f"    REC {r['rec']:>4} 후보 {len(urls)} → {res}")
        time.sleep(0.4)

    with open(os.path.join(FT, "ct_retrieval_status.csv"), "w", newline="",
              encoding="utf-8-sig") as f:
        w_ = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w_.writeheader(); w_.writerows(rows)
    from collections import Counter
    ok = sum(1 for r in rows if r["result"] in ("ok", "already"))
    print(f"\n[2] 신규 확보 {n_new} · 누적 {ok}/{len(rows)}")
    print(f"    결과 분포 {dict(Counter(r['result'] for r in rows))}")


if __name__ == "__main__":
    main()
