# -*- coding: utf-8 -*-
"""
Paper32 — 인용추적 RETRIEVE/UNCERTAIN 79건의 오픈액세스 전문 확보 시도
OpenAlex best_oa_location → PDF 직접 다운로드(가능한 것만). 실패분은 목록으로 남긴다.
출력: ct_pdf/*.pdf · fulltext/ct_retrieval_status.csv
"""
import sys, os, csv, json, time, re, urllib.parse, urllib.request, urllib.error

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
PDF = os.path.join(BASE, "ct_pdf")
os.makedirs(PDF, exist_ok=True)
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) paper32-SR/0.1 "
      "(academic systematic review; mailto:wh8502@naver.com)"}


def api(url, tries=4):
    req = urllib.request.Request(url, headers={"User-Agent": UA["User-Agent"]})
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
    req = urllib.request.Request(url, headers=UA)
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


def main():
    rows = [r for r in csv.DictReader(open(os.path.join(FT, "ct_screen_final.csv"),
                                           encoding="utf-8-sig"))
            if r["final"] in ("RETRIEVE", "UNCERTAIN")]
    print(f"[0] 전문 확보 대상 {len(rows)}건 "
          f"(RETRIEVE {sum(1 for r in rows if r['final']=='RETRIEVE')} · "
          f"UNCERTAIN {sum(1 for r in rows if r['final']=='UNCERTAIN')})")

    ids = [r["openalex_id"] for r in rows if r.get("openalex_id")]
    oa = {}
    for i in range(0, len(ids), 50):
        d = api("https://api.openalex.org/works?filter=" +
                urllib.parse.quote("openalex_id:" + "|".join(ids[i:i + 50])) +
                "&per-page=50&select=id,doi,title,open_access,best_oa_location,"
                "primary_location,publication_year")
        if d:
            for w in d.get("results", []):
                oa[w["id"].replace("https://openalex.org/", "")] = w
        time.sleep(1.0)
    print(f"[1] OpenAlex OA 메타 {len(oa)}/{len(ids)}")

    out = []
    n_ok = 0
    for r in rows:
        w = oa.get(r.get("openalex_id", ""), {})
        oai = (w.get("open_access") or {})
        loc = (w.get("best_oa_location") or {})
        pdfurl = loc.get("pdf_url") or ""
        landing = loc.get("landing_page_url") or ""
        status = oai.get("oa_status", "unknown")
        fn = f"CT{int(r['rec']):04d}_{r.get('year','')}_{safe(r['title'])}.pdf"
        path = os.path.join(PDF, fn)
        res = "no-oa-pdf"
        if os.path.exists(path) and os.path.getsize(path) > 20000:
            res = "already"
            n_ok += 1
        elif pdfurl:
            res = download(pdfurl, path)
            if res == "ok":
                n_ok += 1
            else:
                if os.path.exists(path):
                    os.remove(path)
            time.sleep(1.4)
        out.append({"rec": r["rec"], "screen": r["final"], "oa_status": status,
                    "result": res, "year": r.get("year", ""), "journal": r.get("journal", ""),
                    "doi": r["doi"], "pdf_url": pdfurl, "landing": landing,
                    "behavior_hint": r.get("behavior_hint", ""), "title": r["title"],
                    "file": fn if res in ("ok", "already") else ""})
        print(f"    REC {r['rec']:>4} [{r['final'][:4]}] {status:<8} {res}")

    with open(os.path.join(FT, "ct_retrieval_status.csv"), "w", newline="",
              encoding="utf-8-sig") as f:
        w_ = csv.DictWriter(f, fieldnames=["rec", "screen", "oa_status", "result", "year",
                                           "journal", "doi", "pdf_url", "landing",
                                           "behavior_hint", "file", "title"])
        w_.writeheader(); w_.writerows(out)
    from collections import Counter
    print(f"\n[2] 확보 {n_ok}/{len(rows)} · 결과 분포 {dict(Counter(o['result'] for o in out))}")
    print(f"    OA 상태 {dict(Counter(o['oa_status'] for o in out))}")
    print("    → fulltext/ct_retrieval_status.csv")


if __name__ == "__main__":
    main()
