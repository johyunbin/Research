# -*- coding: utf-8 -*-
"""
Paper32 — 인용추적 1차(제목) 스크리닝 통합 + CANDIDATE/MAYBE 초록 수집(2차 스크리닝용)
출력: fulltext/ct_screen_round1.csv · ct_round2_chunks/ct2_XX.txt
"""
import sys, os, csv, glob, json, time, urllib.parse, urllib.request, urllib.error
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
UA = {"User-Agent": "paper32-SR-citation-tracking/0.1 (academic systematic review)"}


def get(url, tries=5):
    req = urllib.request.Request(url, headers=UA)
    for i in range(tries):
        try:
            with urllib.request.urlopen(req, timeout=60) as r:
                return json.loads(r.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            if e.code in (429, 500, 502, 503) and i < tries - 1:
                time.sleep(4 * (i + 1)); continue
            raise
        except urllib.error.URLError:
            if i < tries - 1:
                time.sleep(4 * (i + 1)); continue
            raise


def inv_to_text(inv):
    if not inv:
        return ""
    pos = {}
    for w, idxs in inv.items():
        for i in idxs:
            pos[i] = w
    return " ".join(pos[i] for i in sorted(pos))


def main():
    # ── 1차 결과 통합 ──────────────────────────────────────────────
    res = {}
    for p in sorted(glob.glob(os.path.join(FT, "ct_results", "ct_0*_screen.csv"))):
        for r in csv.DictReader(open(p, encoding="utf-8-sig")):
            rec = int(str(r["rec"]).strip())
            if rec in res:
                print(f"⚠️ 중복 REC {rec}")
            res[rec] = {"rec": rec, "verdict": r["verdict"].strip().upper(),
                        "reason_code": (r.get("reason_code") or "").strip(),
                        "note": (r.get("note") or "").strip(),
                        "doi": (r.get("doi") or "").strip(), "title": (r.get("title") or "").strip()}
    print(f"[1] 1차 스크리닝 통합 {len(res)}건 · {dict(Counter(v['verdict'] for v in res.values()))}")

    # 원 triage 메타 결합 — tier는 triaged, openalex_id는 원본 citation_tracking에서
    tri, oa = {}, {}
    with open(os.path.join(FT, "citation_tracking_triaged.csv"), encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            tri[(r.get("title") or "").strip()] = r
    with open(os.path.join(FT, "citation_tracking.csv"), encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            oa[(r.get("title") or "").strip()] = r.get("openalex_id", "")
    hit, hit_oa = 0, 0
    for v in res.values():
        m = tri.get(v["title"])
        if m:
            hit += 1
            v["route"] = m.get("route", "")
            v["year"] = m.get("year", "")
            v["journal"] = m.get("journal", "")
            v["tier"] = m.get("tier", "")
        v["openalex_id"] = oa.get(v["title"], "")
        if v["openalex_id"]:
            hit_oa += 1
    print(f"[2] triage 메타 결합 {hit}/{len(res)} · OpenAlex ID {hit_oa}/{len(res)}")

    keys = ["rec", "verdict", "reason_code", "note", "openalex_id", "route", "tier",
            "year", "journal", "doi", "title"]
    with open(os.path.join(FT, "ct_screen_round1.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=keys, extrasaction="ignore"); w.writeheader()
        for k in sorted(res): w.writerow(res[k])

    # ── 2차 대상: CANDIDATE + MAYBE 초록 수집 ────────────────────────
    targets = [v for v in res.values() if v["verdict"] in ("CANDIDATE", "MAYBE")]
    ids = [v.get("openalex_id", "") for v in targets if v.get("openalex_id")]
    print(f"[3] 2차 대상 {len(targets)}건 (OpenAlex ID 보유 {len(ids)})")

    abst = {}
    for i in range(0, len(ids), 50):
        chunk = ids[i:i + 50]
        url = ("https://api.openalex.org/works?filter=" +
               urllib.parse.quote("openalex_id:" + "|".join(chunk)) +
               "&per-page=50&select=id,abstract_inverted_index,type,publication_year")
        try:
            data = get(url)
        except Exception as e:
            print(f"    초록 배치 {i//50+1} 실패: {e}"); continue
        for w_ in data.get("results", []):
            abst[w_["id"].replace("https://openalex.org/", "")] = inv_to_text(
                w_.get("abstract_inverted_index"))
        time.sleep(1.0)
    got = sum(1 for v in abst.values() if v)
    print(f"[4] 초록 확보 {got}/{len(ids)}")

    # ── 청크 작성 ─────────────────────────────────────────────────
    OUT = os.path.join(FT, "ct_round2_chunks")
    os.makedirs(OUT, exist_ok=True)
    for p in glob.glob(os.path.join(OUT, "*.txt")):
        os.remove(p)
    targets.sort(key=lambda v: v["rec"])
    SZ = 30
    n_chunk = 0
    for i in range(0, len(targets), SZ):
        n_chunk += 1
        lines = []
        for v in targets[i:i + SZ]:
            ab = abst.get(v.get("openalex_id", ""), "")
            ab = ab if ab else "(초록 미확보 — 제목·서지로만 판단)"
            if len(ab) > 2400:
                ab = ab[:2400] + " …[절단]"
            lines.append(f"### REC {v['rec']} | 1차={v['verdict']} | {v.get('year','')} | "
                         f"{v.get('journal','')} | {v.get('route','')}\n"
                         f"TITLE: {v['title']}\nDOI: {v['doi']}\n1차메모: {v['note']}\n"
                         f"ABSTRACT: {ab}\n")
        open(os.path.join(OUT, f"ct2_{n_chunk:02d}.txt"), "w", encoding="utf-8").write("\n".join(lines))
    print(f"[5] 2차 청크 {n_chunk}개 (건당 최대 {SZ}) → {OUT}")


if __name__ == "__main__":
    main()
