# -*- coding: utf-8 -*-
"""
Paper32 — OpenAlex 보조검색 스크리닝 통과분(15건) 전문 확보 시도
문헌유형·언어를 먼저 확인해 명백한 부적격(학회초록·업계지·비영어)을 걸러내고, 나머지만 회수한다.
출력: fulltext/oa_supp_retrieval.csv · oa_supp_pdf/*.pdf
"""
import sys, os, csv, re, json, time, urllib.parse, urllib.request, urllib.error

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
PDF = os.path.join(BASE, "oa_supp_pdf")
os.makedirs(PDF, exist_ok=True)
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
      "Chrome/125.0 Safari/537.36")
MAIL = "wh8502@naver.com"


def get(url, tries=3):
    req = urllib.request.Request(url, headers={"User-Agent": f"paper32-SR/0.1 (mailto:{MAIL})"})
    for i in range(tries):
        try:
            with urllib.request.urlopen(req, timeout=45) as r:
                return json.loads(r.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            if e.code in (429, 500, 502, 503) and i < tries - 1:
                time.sleep(4 * (i + 1)); continue
            return None
        except Exception:
            if i < tries - 1:
                time.sleep(4 * (i + 1)); continue
            return None


def safe(s, n=80):
    s = re.sub(r"[^\w\s\-.,()]", "", s, flags=re.UNICODE)
    return re.sub(r"\s+", " ", s).strip()[:n]


def download(url, path):
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "application/pdf,*/*"})
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


# 문헌유형 사전 배제 규칙 — 등록 기준 ⑤(영어 동료심사 저널 논문)
PRESCREEN = [
    ("The Hearing Journal", "업계지(trade magazine) — 동료심사 저널 아님"),
    ("ITF Coaching", "코칭 뉴스레터 — 동료심사 저널 아님"),
]


def main():
    rows = [r for r in csv.DictReader(open(os.path.join(FT, "oa_supp_screen_final.csv"),
                                           encoding="utf-8-sig"))
            if r["verdict"] in ("RETRIEVE", "UNCERTAIN")]
    print(f"[0] 대상 {len(rows)}건")

    out = []
    for r in rows:
        j = r.get("journal", "") or ""
        pre = next((why for pat, why in PRESCREEN if pat.lower() in j.lower()), None)

        meta = {}
        if r.get("doi"):
            d = get("https://api.openalex.org/works/doi:" + urllib.parse.quote(r["doi"]))
            if d:
                loc = (d.get("best_oa_location") or {})
                meta = {"type": d.get("type", ""), "language": d.get("language", ""),
                        "crossref_type": d.get("type_crossref", ""),
                        "pdf_url": loc.get("pdf_url") or "",
                        "oa_status": (d.get("open_access") or {}).get("oa_status", "")}
            time.sleep(0.8)

        # 유형 기반 사전 배제
        if not pre:
            ct = (meta.get("crossref_type") or "").lower()
            if ct in ("proceedings-article", "posted-content", "component"):
                pre = f"문헌유형 {ct} — 등록 기준(저널 논문) 미충족"
            elif meta.get("language") and meta["language"] != "en":
                pre = f"언어 {meta['language']} — 등록 기준(영어) 미충족"

        if pre:
            out.append({**r, **meta, "result": "prescreen-exclude", "reason": pre, "file": ""})
            print(f"    SID {r['sid']:>3} ✕ {pre}")
            continue

        fn = f"OAS{int(r['sid']):04d}_{r.get('year','')}_{safe(r['title'])}.pdf"
        path = os.path.join(PDF, fn)
        res = "no-oa-pdf"
        if os.path.exists(path) and os.path.getsize(path) > 20000:
            res = "already"
        elif meta.get("pdf_url"):
            res = download(meta["pdf_url"], path)
            if res != "ok" and os.path.exists(path):
                os.remove(path)
            time.sleep(1.2)
        out.append({**r, **meta, "result": res, "reason": "",
                    "file": fn if res in ("ok", "already") else ""})
        print(f"    SID {r['sid']:>3} {meta.get('oa_status','?'):<8} {res}")

    keys = ["sid", "verdict", "year", "journal", "doi", "title", "type", "crossref_type",
            "language", "oa_status", "result", "reason", "file"]
    with open(os.path.join(FT, "oa_supp_retrieval.csv"), "w", newline="",
              encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=keys, extrasaction="ignore")
        w.writeheader(); w.writerows(out)

    from collections import Counter
    got = sum(1 for r in out if r["result"] in ("ok", "already"))
    pre_n = sum(1 for r in out if r["result"] == "prescreen-exclude")
    print(f"\n[완료] 사전배제 {pre_n} · 확보 {got} · 미확보 {len(out)-pre_n-got}")
    print(f"  결과 {dict(Counter(r['result'] for r in out))}")


if __name__ == "__main__":
    main()
