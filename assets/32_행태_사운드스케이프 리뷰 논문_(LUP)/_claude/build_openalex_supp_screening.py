# -*- coding: utf-8 -*-
"""
Paper32 — OpenAlex 보조검색 미스크리닝분 처리 (등록 약속 이행)
등록 프로토콜에 보조 식별원으로 명시한 OpenAlex 고유분 352건 중, 3-DB 풀·인용추적 풀·최종
코퍼스 어디에도 들어가지 않은 건을 찾아 제목 스크리닝 배치를 만든다.
초록은 OpenAlex에서 함께 받아 붙인다(있는 것만).
출력: fulltext/oa_supp_unscreened.csv · fulltext/oa_supp_chunks/oas_XX.txt
"""
import sys, os, csv, re, glob, json, time, unicodedata
import urllib.parse, urllib.request, urllib.error

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
OUT = os.path.join(FT, "oa_supp_chunks")
os.makedirs(OUT, exist_ok=True)
UA = {"User-Agent": "paper32-SR/0.1 (mailto:wh8502@naver.com)"}


def nt(t):
    t = unicodedata.normalize("NFKD", (t or "").lower())
    return re.sub(r"[^a-z0-9]", "", t)


def nd(d):
    return re.sub(r"^https?://(dx\.)?doi\.org/", "", (d or "").strip().lower())


def get(url, tries=4):
    req = urllib.request.Request(url, headers=UA)
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


def inv_to_text(inv):
    if not inv:
        return ""
    pos = {}
    for w, idxs in inv.items():
        for i in idxs:
            pos[i] = w
    return " ".join(pos[i] for i in sorted(pos))


def main():
    sup = list(csv.DictReader(open(glob.glob(os.path.join(BASE, "openalex_supplement_*.csv"))[0],
                                   encoding="utf-8-sig")))
    seen_t, seen_d = set(), set()
    srcs = [(os.path.join(BASE, "merged_pool_20260802_211715.csv"), "title", "doi"),
            (os.path.join(BASE, "screening", "screening_results_20260802_221808.csv"), "title", "doi"),
            (os.path.join(FT, "citation_tracking.csv"), "title", "doi"),
            (os.path.join(FT, "corpus_v3_extraction.csv"), "title", None)]
    for p, tc, dc in srcs:
        if not os.path.exists(p):
            print(f"  ⚠️ 없음: {os.path.basename(p)}"); continue
        for r in csv.DictReader(open(p, encoding="utf-8-sig")):
            seen_t.add(nt(r.get(tc)))
            if dc and r.get(dc):
                seen_d.add(nd(r[dc]))
    new = [r for r in sup if nt(r["title"]) not in seen_t and nd(r.get("doi")) not in seen_d]
    print(f"[0] 보조검색 {len(sup)}건 → 미스크리닝 {len(new)}건")

    # 초록 확보(OpenAlex DOI 배치 조회)
    ab = {}
    dois = [nd(r["doi"]) for r in new if r.get("doi")]
    for i in range(0, len(dois), 45):
        chunk = dois[i:i + 45]
        d = get("https://api.openalex.org/works?filter=" +
                urllib.parse.quote("doi:" + "|".join(chunk)) +
                "&per-page=45&select=doi,abstract_inverted_index,type,language")
        if d:
            for w in d.get("results", []):
                ab[nd(w.get("doi"))] = {"abstract": inv_to_text(w.get("abstract_inverted_index")),
                                        "type": w.get("type", ""), "language": w.get("language", "")}
        time.sleep(1.0)
        if (i // 45) % 3 == 0:
            print(f"    초록 {i}/{len(dois)} · 확보 {sum(1 for v in ab.values() if v['abstract'])}")
    got = sum(1 for v in ab.values() if v["abstract"])
    print(f"[1] 초록 확보 {got}/{len(dois)}")

    rows = []
    for i, r in enumerate(new, 1):
        m = ab.get(nd(r.get("doi")), {})
        rows.append({"sid": i, "year": r.get("year", ""), "journal": r.get("journal", ""),
                     "doi": r.get("doi", ""), "title": r.get("title", ""),
                     "type": m.get("type", ""), "language": m.get("language", ""),
                     "has_abstract": "yes" if m.get("abstract") else "no",
                     "abstract": m.get("abstract", "")})
    with open(os.path.join(FT, "oa_supp_unscreened.csv"), "w", newline="",
              encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys())); w.writeheader(); w.writerows(rows)

    for p in os.listdir(OUT):
        os.remove(os.path.join(OUT, p))
    SZ = 110
    n = 0
    for i in range(0, len(rows), SZ):
        n += 1
        lines = []
        for r in rows[i:i + SZ]:
            a = r["abstract"]
            if len(a) > 1400:
                a = a[:1400] + " …[절단]"
            lines.append(f"### SID {r['sid']} | {r['year']} | {r['journal']}\n"
                         f"TITLE: {r['title']}\nDOI: {r['doi']}\n"
                         f"TYPE: {r['type']} | LANG: {r['language']}\n"
                         f"ABSTRACT: {a or '(초록 미확보 — 제목·저널로만 판단)'}\n")
        open(os.path.join(OUT, f"oas_{n:02d}.txt"), "w",
             encoding="utf-8").write("\n".join(lines))
    from collections import Counter
    print(f"[2] 배치 {n}개 (건당 최대 {SZ}) → {OUT}")
    print(f"    유형 {dict(Counter(r['type'] for r in rows).most_common(5))}")
    print(f"    언어 {dict(Counter(r['language'] for r in rows).most_common(5))}")


if __name__ == "__main__":
    main()
