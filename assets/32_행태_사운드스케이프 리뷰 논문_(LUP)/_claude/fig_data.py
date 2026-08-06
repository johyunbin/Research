# -*- coding: utf-8 -*-
"""
Paper32 — 모든 Figure 가 쓰는 수치의 **단일 출처**

★ 왜 이 파일이 생겼는가
  Fig 1·5·6 이 각각 96편·102편·81편 기준의 숫자를 그리고 있었다(2026-08-06 육안 점검에서 발견).
  원인은 하나다 — **각 그림 스크립트가 수치를 코드에 하드코딩**했고, 정본 데이터가 갱신돼도
  그림은 따라오지 않았다. 여기서 한 번 산출해 `figures/fig_data.json` 으로 내보내고,
  모든 그림은 그 파일만 읽는다. **그림 스크립트에 숫자 리터럴을 두지 않는다.**

출력: figures/fig_data.json
"""
import sys, os, csv, re, json
from collections import Counter, defaultdict

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
FIG = os.path.join(BASE, "figures")
os.makedirs(FIG, exist_ok=True)
sys.path.insert(0, BASE)
from make_fig7_geo_time import norm_countries

DOMAINS = ["movement", "staying", "space-use", "social", "activity"]
DOM_EN = {"movement": "Movement", "staying": "Staying", "space-use": "Space use",
          "social": "Social", "activity": "Activity"}
SOURCES = [
    ("Traffic / road", r"traffic|road|도로|교통|차량|vehicle|street noise|LAeq.*road"),
    ("General noise", r"\bnoise\b|소음|dB\b|LAeq|LZeq|sound (?:pressure )?level|SPL"),
    ("Natural sounds", r"natural sound|nature sound|bird|water|stream|fountain|wind|자연음|새소리|물소리|조류|leaves"),
    ("Music / added", r"music|음악|speaker|스피커|broadcast|audio|APS|audible.*signal|masking"),
    ("Human / crowd", r"crowd|human sound|voice|conversation|대화|군중|footstep|발소리|people sound|speech"),
    ("Aircraft", r"aircraft|airport|항공|비행"),
]


def rd(p):
    return list(csv.DictReader(open(os.path.join(FT, p), encoding="utf-8-sig")))


def main():
    ext = rd("corpus_v4_extraction.csv")
    verd = {r["uid"]: r["final_verdict"] for r in rd("corpus_v4_verdicts.csv")}
    qual = {r["uid"]: r for r in rd("quality_v2.csv")}
    t1 = {r["uid"]: r for r in rd("table1_v2.csv")}
    inc = [r for r in ext if verd.get(r["uid"]) == "FINAL_INCLUDE"]
    N = len(inc)
    D = {"n_included": N,
         "n_sens": sum(1 for v in verd.values() if v == "SENS_ONLY"),
         "by_route": dict(Counter(r["source"] for r in inc))}
    D["n_analysed"] = N + D["n_sens"]

    # ── Fig 3: 도메인 × 음원 ─────────────────────────────────────────
    def srcs(r):
        blob = f"{r['exposure']} {r['title']}".lower()
        hit = [n for n, p in SOURCES if re.search(p, blob, re.I)]
        if len(hit) > 1 and "General noise" in hit:
            hit = [h for h in hit if h != "General noise"]
        return hit or ["Other"]

    grid = defaultdict(Counter)
    for r in inc:
        for d in (x.strip().lower() for x in (r["behaviour_domain"] or "").split(";")):
            if d:
                for s in srcs(r):
                    grid[d][s] += 1
    D["evidence_map"] = {
        "rows": [DOM_EN[d] for d in DOMAINS],
        "cols": [n for n, _ in SOURCES],
        "cells": [[grid[d][s] for s, _ in SOURCES] for d in DOMAINS],
    }
    D["n_aircraft_records"] = sum(grid[d]["Aircraft"] for d in DOMAINS)
    D["n_aircraft_studies"] = sum(1 for r in inc if "Aircraft" in srcs(r))

    # ── Fig 4: 방향 × 도메인 ─────────────────────────────────────────
    dom = defaultdict(Counter)
    for r in inc:
        for d in (x.strip().lower() for x in (r["behaviour_domain"] or "").split(";")):
            if d:
                dom[d][(r["direction"] or "").strip()] += 1
    D["direction"] = {"rows": [DOM_EN[d] for d in DOMAINS],
                      "forward": [dom[d]["forward"] for d in DOMAINS],
                      "reverse": [dom[d]["reverse"] for d in DOMAINS],
                      "both": [dom[d]["both"] for d in DOMAINS]}
    tf = sum(D["direction"]["forward"]); tr = sum(D["direction"]["reverse"])
    tb = sum(D["direction"]["both"])
    D["direction"].update(total_forward=tf, total_reverse=tr, total_both=tb,
                          pct_reverse_of_directional=round(tr / (tf + tr) * 100),
                          n_directional=tf + tr, n_records=tf + tr + tb)

    # ── Fig 5: 측정세대 × 시기 (Table 1 의 measure_gen 이 단일 출처) ──
    BANDS = ["≤2009", "2010–2019", "2020–2026"]
    band = defaultdict(Counter); multi = 0
    for r in inc:
        gs = {g.strip() for g in (t1.get(r["uid"], {}).get("measure_gen", "") or "").split(";")
              if g.strip() in ("G1", "G2", "G3")}
        if len(gs) >= 2:
            multi += 1
        try:
            y = int(r["year"])
        except (TypeError, ValueError):
            continue
        b = BANDS[0] if y < 2010 else (BANDS[1] if y < 2020 else BANDS[2])
        for g in gs:
            band[b][g] += 1
    D["generations"] = {"bands": BANDS,
                        "G1": [band[b]["G1"] for b in BANDS],
                        "G2": [band[b]["G2"] for b in BANDS],
                        "G3": [band[b]["G3"] for b in BANDS],
                        "multi_generation_studies": multi}
    for g in ("G1", "G2", "G3"):
        D["generations"][f"total_{g}"] = sum(D["generations"][g])

    # ── Fig 8: MMAT ─────────────────────────────────────────────────
    CAT = {"1": "Qualitative", "2": "Quantitative RCT", "3": "Quantitative non-randomised",
           "4": "Quantitative descriptive", "5": "Mixed methods"}
    det = [r for r in rd("quality_detail_v2.csv") if verd.get(r["uid"]) == "FINAL_INCLUDE"]
    cat_of = {}
    for r in det:
        if r["item_no"]:
            cat_of[r["uid"]] = r["item_no"][0]
    tier = Counter(qual[r["uid"]]["quality_tier"] for r in inc if r["uid"] in qual)
    by_cat = defaultdict(Counter)
    for u, c in cat_of.items():
        by_cat[c][qual.get(u, {}).get("quality_tier", "?")] += 1
    items = defaultdict(Counter)
    for r in det:
        if r["item_no"]:
            items[r["item_no"]][r["verdict"]] += 1
    D["quality"] = {
        "tier": {k: tier[k] for k in ("high", "moderate", "low")},
        "categories": [{"key": c, "name": CAT[c],
                        "n": sum(by_cat[c].values()),
                        "high": by_cat[c]["high"], "moderate": by_cat[c]["moderate"],
                        "low": by_cat[c]["low"]} for c in "12345"],
        "items": {k: {"Y": v["Y"], "N": v["N"], "CT": v["CT"], "n": sum(v.values())}
                  for k, v in sorted(items.items())},
    }

    # ── Fig 2·6: 메타분석 ────────────────────────────────────────────
    fp = os.path.join(FT, "ma", "ma_forest_data.json")
    if os.path.exists(fp):
        D["ma"] = json.load(open(fp, encoding="utf-8"))
    else:
        print("⚠️ ma_forest_data.json 없음 — ma_v2.py 를 먼저 실행")
    # 각 클러스터 기여 연구의 품질 구성(Fig 6 관여경사 밴드용)
    if "ma" in D:
        for key, blk in D["ma"].items():
            c = Counter(qual.get(str(e["uid"]), {}).get("quality_tier", "?")
                        for e in blk["effects"])
            blk["quality_mix"] = {k: c[k] for k in ("high", "moderate", "low")}

    # ── Fig 7: 지리·시기 ─────────────────────────────────────────────
    cc = Counter(); yr = defaultdict(Counter)
    for r in inc:
        cs = norm_countries(r["country"])
        for c in (cs or ["Not reported"]):
            cc[c] += 1
        try:
            yr[int(r["year"])][(r["direction"] or "forward").strip()] += 1
        except (TypeError, ValueError):
            pass
    D["geo"] = {"countries": cc.most_common(), "n_countries": len(cc)}
    D["time"] = {"years": sorted(yr),
                 "forward": [yr[y]["forward"] for y in sorted(yr)],
                 "reverse": [yr[y]["reverse"] for y in sorted(yr)],
                 "both": [yr[y]["both"] for y in sorted(yr)]}
    D["n_since_2020"] = sum(1 for r in inc
                            if (r["year"] or "").strip().isdigit() and int(r["year"]) >= 2020)

    # ── Fig 1: PRISMA (흐름 수치는 prisma_flow.md 가 정본) ───────────
    D["prisma"] = {
        "db": {"identified": 2073, "wos": 1010, "scopus": 850, "pubmed": 213,
               "duplicates": 757, "screened": 1316, "excluded": 1127,
               "excl": [("Animal / wildlife", 479), ("Perception / health", 325),
                        ("Setting not eligible", 133), ("Not empirical", 109),
                        ("No acoustic variable", 81)],
               "sought": 189, "not_retrieved": 89,
               "nr": [("No institution access", 84), ("Pay-per-view only", 3),
                      ("Abstract only", 2)],
               "assessed": 100, "ft_excluded": 16,
               "ftx": [("Setting not eligible", 7), ("No observed behaviour", 6),
                       ("Perceptual outcome only", 1), ("No acoustic exposure", 1),
                       ("Not empirical", 1)],
               "sens": 3, "included": 81},
        "ct": {"identified": 2073, "backward": 413, "forward": 1660, "seeds": 84,
               "deprioritised": 1645,
               "dep": [("Animal / acoustics", 26), ("≤1 block matched", 1619)],
               "screened_title": 428, "excl_title": 282,
               "et": [("No behavioural outcome", 211), ("No acoustic variable", 29),
                      ("Not empirical", 26), ("Setting not eligible", 10), ("Animal", 6)],
               "screened_abs": 146, "excl_abs": 67,
               "ea": [("No acoustic variable", 37), ("No behavioural outcome", 26),
                      ("Other", 4)],
               "sought": 79, "not_retrieved": 25,
               "nr": [("No institution access", 22), ("Pay-per-view only", 3)],
               "assessed": 54, "ft_excluded": 38,
               "ftx": [("No acoustic variable", 20), ("No observed behaviour", 17),
                       ("Not in English", 1)],
               "sens": 1, "included": 15},
        "supp": {"identified": 352, "duplicates": 23, "screened": 329, "excluded": 314,
                 "excl": [("Not empirical (reviews etc.)", 147),
                          ("No behavioural outcome", 46), ("Animal / bioacoustics", 39),
                          ("Off topic", 31), ("Setting not eligible", 27),
                          ("No acoustic variable", 12), ("Language / document type", 12)],
                 "sought": 15, "prescreen": 2, "not_retrieved": 8,
                 "assessed": 5, "ft_excluded": 3,
                 "ftx": [("Not in English", 2),
                         ("No behavioural outcome", 1)],
                 "included": 2},
    }
    p = D["prisma"]
    tot = p["db"]["included"] + p["ct"]["included"] + p["supp"]["included"]
    if tot != N:
        raise SystemExit(f"⚠️ PRISMA 합({tot}) ≠ 코퍼스({N}) — prisma_flow.md 와 대조 필요")

    with open(os.path.join(FIG, "fig_data.json"), "w", encoding="utf-8") as f:
        json.dump(D, f, ensure_ascii=False, indent=1)

    print(f"[검증] PRISMA 3갈래 합 {tot} = 코퍼스 {N} ✅")
    print(f"  포함 {N} · 민감도 {D['n_sens']} · 분석 {D['n_analysed']} · 갈래 {D['by_route']}")
    print(f"  방향 fwd {tf} · rev {tr} · both {tb} → 역방향 "
          f"{D['direction']['pct_reverse_of_directional']}%")
    print(f"  세대 G1 {D['generations']['total_G1']} · G2 {D['generations']['total_G2']} · "
          f"G3 {D['generations']['total_G3']} · 병용 {multi}")
    print(f"  품질 {D['quality']['tier']}")
    print(f"  항공기 레코드 {D['n_aircraft_records']} (연구 {D['n_aircraft_studies']}편)")
    if "ma" in D:
        for k, v in D["ma"].items():
            pl = v["pooled"]
            print(f"  MA {k:12s} k={pl['k']} est={pl['est']:+.3f} p={pl['p']:.3f} "
                  f"quality={v['quality_mix']}")
    print("[저장] figures/fig_data.json")


if __name__ == "__main__":
    main()
