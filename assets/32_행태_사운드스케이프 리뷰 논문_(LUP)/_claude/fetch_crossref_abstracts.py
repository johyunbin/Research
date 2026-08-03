# -*- coding: utf-8 -*-
"""
Paper32 — 인용추적 UNCERTAIN(초록 미확보) 건의 초록을 Crossref에서 보완 수집
출력: fulltext/ct_round3_chunks/ct3_XX.txt (초록 확보분) + ct_uncertain_no_abstract.csv
"""
import sys, os, csv, json, time, re, html, urllib.parse, urllib.request, urllib.error

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
UA = {"User-Agent": "paper32-SR/0.1 (mailto:wh8502@naver.com)"}


def get(url, tries=4):
    req = urllib.request.Request(url, headers=UA)
    for i in range(tries):
        try:
            with urllib.request.urlopen(req, timeout=45) as r:
                return json.loads(r.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            if e.code == 404:
                return None
            if e.code in (429, 500, 502, 503) and i < tries - 1:
                time.sleep(3 * (i + 1)); continue
            return None
        except Exception:
            if i < tries - 1:
                time.sleep(3 * (i + 1)); continue
            return None


def clean(x):
    if not x:
        return ""
    x = re.sub(r"<[^>]+>", " ", x)
    x = html.unescape(x)
    x = re.sub(r"\s+", " ", x).strip()
    return re.sub(r"^(Abstract|ABSTRACT)\s*", "", x)


def main():
    rows = [r for r in csv.DictReader(open(os.path.join(FT, "ct_screen_final.csv"),
                                           encoding="utf-8-sig"))
            if r["final"] == "UNCERTAIN"]
    print(f"[0] UNCERTAIN {len(rows)}건")

    got, miss = [], []
    for i, r in enumerate(rows, 1):
        doi = (r["doi"] or "").strip()
        ab = ""
        if doi:
            d = get("https://api.crossref.org/works/" + urllib.parse.quote(doi))
            if d:
                ab = clean((d.get("message") or {}).get("abstract", ""))
            time.sleep(0.6)
        (got if ab else miss).append({**r, "abstract": ab})
        if i % 10 == 0:
            print(f"    {i}/{len(rows)} · 확보 {len(got)}")
    print(f"[1] Crossref 초록 확보 {len(got)} · 미확보 {len(miss)}")

    OUT = os.path.join(FT, "ct_round3_chunks")
    os.makedirs(OUT, exist_ok=True)
    for p in os.listdir(OUT):
        os.remove(os.path.join(OUT, p))
    SZ = 15
    for c, i in enumerate(range(0, len(got), SZ), start=1):
        lines = []
        for v in got[i:i + SZ]:
            ab = v["abstract"]
            if len(ab) > 2600:
                ab = ab[:2600] + " …[절단]"
            lines.append(f"### REC {v['rec']} | {v.get('year','')} | {v.get('journal','')}\n"
                         f"TITLE: {v['title']}\nDOI: {v['doi']}\n"
                         f"1차메모: {v.get('note','')}\nABSTRACT: {ab}\n")
        open(os.path.join(OUT, f"ct3_{c:02d}.txt"), "w", encoding="utf-8").write("\n".join(lines))
    print(f"[2] 3차 청크 {(len(got)+SZ-1)//SZ}개 → {OUT}")

    with open(os.path.join(FT, "ct_uncertain_no_abstract.csv"), "w", newline="",
              encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["rec", "year", "journal", "doi", "title", "note"],
                           extrasaction="ignore")
        w.writeheader(); w.writerows(miss)
    print(f"[3] 초록 끝내 미확보 {len(miss)}건 → ct_uncertain_no_abstract.csv (전문 확보 대상)")


if __name__ == "__main__":
    main()
