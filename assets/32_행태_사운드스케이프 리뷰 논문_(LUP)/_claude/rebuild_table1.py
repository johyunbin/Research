# -*- coding: utf-8 -*-
"""
Paper32 — Table 1 재작성 v2 (두 갈래 통합 100편)
기존 table1_study_characteristics.csv(84행, DB 갈래)를 보존하고, 인용추적 16편을 같은
정규화 규칙으로 코딩해 합친다. sid는 연도→저자 순으로 다시 매긴다.
출력: fulltext/table1_v2.csv · table1_v2.md
"""
import sys, os, csv, re
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")

SETTING = [("lab(outdoor scene)", r"\blab\b|실험실|무향|anechoic|VR|CAVE|IVP|앰비소닉|ambisonic|simulator"),
           ("waterfront", r"waterfront|수변|해안|coastal|lake|호수|river"),
           ("campus", r"campus|캠퍼스|university|대학"),
           ("residential", r"residential|주거|courtyard|중정|neighbourhood|neighborhood"),
           ("recreation", r"recreation|휴양|forest park|trail|tourist|관광"),
           ("square", r"square|광장|plaza"),
           ("park", r"park|공원|garden|정원|green space"),
           ("street", r"street|가로|road|sidewalk|보도|crossing|횡단|urban open"),
           ("mixed", r"mixed|복합|multiple|여러")]
DESIGN = [("field experiment", r"field experiment|현장실험|현장 실험|in-situ.*experiment"),
          ("lab experiment", r"lab experiment|실험실 실험|laboratory experiment|listening test|청취실험"),
          ("quasi-experiment", r"quasi|준실험|natural experiment|자연실험"),
          ("qualitative", r"qualitative|질적|ethnograph|민족지|interview only"),
          ("mixed", r"mixed|혼합"),
          ("observational", r"observational|관찰|behaviour map|행태매핑|behavioural map|big data|빅데이터"),
          ("survey", r"survey|설문|questionnaire|cross-sectional|횡단")]


def pick(text, table, default="mixed"):
    t = (text or "").lower()
    for label, pat in table:
        if re.search(pat, t, re.I):
            return label
    return default


def gens(method):
    m = (method or "").lower()
    g = []
    if re.search(r"자기보고|self|survey|questionnaire|설문|interview|인터뷰|recall", m):
        g.append("G1")
    if re.search(r"관찰|observation|mapping|counts|계수|추적|behaviour map", m):
        g.append("G2")
    if re.search(r"센서|sensor|gps|영상|video|빅데이터|big data|wearable|ml|camera|tracking|기기", m):
        g.append("G3")
    return ";".join(g) or "NR"


def short(s, n=52):
    s = re.sub(r"\s+", " ", (s or "").strip())
    return (s[:n] + "…") if len(s) > n else s


def main():
    old = {r["sid"]: r for r in csv.DictReader(open(
        os.path.join(FT, "table1_study_characteristics.csv"), encoding="utf-8-sig"))}
    ext = {r["uid"]: r for r in csv.DictReader(open(
        os.path.join(FT, "corpus_v4_extraction.csv"), encoding="utf-8-sig"))}
    qual = {r["uid"]: r["quality_tier"] for r in csv.DictReader(open(
        os.path.join(FT, "quality_v2.csv"), encoding="utf-8-sig"))}
    verd = {r["uid"]: r["final_verdict"] for r in csv.DictReader(open(
        os.path.join(FT, "corpus_v4_verdicts.csv"), encoding="utf-8-sig"))}

    # ── 기존 84행 ↔ uid 매핑 ──────────────────────────────────────
    # 구 Table 1에는 uid 컬럼이 없다. 저자명은 제목에 없으므로 제목 매칭은 불가.
    # (year, direction, behaviour_domain) 조합이 거의 고유하므로 이를 키로 쓰고,
    # 충돌하면 n·country로 좁힌다. 매칭률을 반드시 보고한다.
    def norm_dom(s):
        return ";".join(sorted(d.strip().lower()
                               for d in (s or "").replace(",", ";").split(";") if d.strip()))

    old_idx = {}
    for r in old.values():
        old_idx.setdefault((r["year"], r["direction"], norm_dom(r["behaviour_domain"])),
                           []).append(r)

    def find_old(r):
        cands = old_idx.get((r["year"], r["direction"], norm_dom(r["behaviour_domain"])), [])
        if len(cands) == 1:
            return cands[0]
        for c in cands:                       # 동률이면 n → country 순으로 좁힌다
            if c["n"] and c["n"] == short(r["sample_n"], 18):
                return c
        for c in cands:
            if c["country"] and c["country"][:8].lower() in (r["country"] or "").lower():
                return c
        return cands[0] if cands else None

    MA = {"461": "MA1", "532": "MA1", "617": "MA1", "481": "MA1(sens)",
          "323": "MA2", "665": "MA2", "1280": "MA2",
          "931": "MA3", "1069": "MA3", "14": "MA3", "CT0025": "MA3",
          "1076": "MA4", "1221": "MA4", "1177": "MA4", "980": "MA4", "1018": "MA4",
          "CT0126": "MA4", "CT0184": "MA4",
          "CT0414": "MA3(sens)", "CT0175": "MA2(sens)", "CT0090": "MA1(sens)",
          "CT0137": "MA4(var)"}

    rows, matched, used = [], 0, set()
    for uid, r in ext.items():
        src = r["source"]
        base = {}
        if src == "db-search":
            hit = find_old(r)
            if hit and hit["sid"] not in used:
                base = hit
                used.add(hit["sid"])
                matched += 1
        rows.append({
            "uid": uid, "source": src,
            "study": base.get("study") or f"[{uid}] {short(r['title'], 34)}",
            "year": r["year"],
            "country": base.get("country") or short(r["country"], 26) or "NR",
            "setting": base.get("setting") or pick(f"{r['setting']} {r['title']}", SETTING),
            "design": base.get("design") or pick(f"{r['design']} {r['measurement_method']}", DESIGN),
            "n": base.get("n") or short(r["sample_n"], 18) or "NR",
            "exposure_short": base.get("exposure_short") or short(r["exposure"], 46),
            "behaviour_domain": r["behaviour_domain"],
            "measure_gen": base.get("measure_gen") or gens(r["measurement_method"]),
            "direction": r["direction"],
            "quality": qual.get(uid, "NR"),
            "in_ma": MA.get(uid, ""),
            "verdict": verd.get(uid, ""),
        })

    def sortkey(r):
        try:
            y = int(r["year"])
        except (TypeError, ValueError):
            y = 9999
        return (y, r["study"].lower())

    rows.sort(key=sortkey)
    for i, r in enumerate(rows, 1):
        r["sid"] = i

    cols = ["sid", "uid", "source", "study", "year", "country", "setting", "design", "n",
            "exposure_short", "behaviour_domain", "measure_gen", "direction", "quality",
            "in_ma", "verdict"]
    with open(os.path.join(FT, "table1_v2.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
        w.writeheader(); w.writerows(rows)

    # ── 마크다운 ───────────────────────────────────────────────────
    L = [f"# Table 1 — Study characteristics ({len(rows)}편)\n",
         "\n두 식별 경로 통합. `source` = db-search(데이터베이스 검색) / citation-tracking(인용 추적).\n",
         "\n| # | Study | Year | Country | Setting | Design | n | Exposure | Domain | Gen | Dir | Quality | MA |\n",
         "|---|---|---|---|---|---|---|---|---|---|---|---|---|\n"]
    for r in rows:
        mark = " ▲" if r["source"] == "citation-tracking" else ""
        L.append(f"| {r['sid']} | {short(r['study'], 30)}{mark} | {r['year']} | "
                 f"{short(r['country'], 16)} | {r['setting']} | {r['design']} | {r['n']} | "
                 f"{short(r['exposure_short'], 40)} | {short(r['behaviour_domain'], 22)} | "
                 f"{r['measure_gen']} | {r['direction']} | {r['quality']} | {r['in_ma']} |\n")
    L.append("\n**약어** — Gen: G1 자기보고 · G2 체계적 관찰 · G3 센싱·궤적(복수 가능) · "
             "Dir: forward(음→행태) / reverse(행태→음) / both · "
             "MA: 메타분석 기여 클러스터, `(sens)`=민감도 전용, `(var)`=변형분석 · "
             "NR: not reported · ▲ = 인용추적으로 추가된 연구\n")
    L.append(f"\n**민감도 전용 {sum(1 for r in rows if r['verdict']=='SENS_ONLY')}편**은 "
             "주분석에서 제외되고 민감도에만 쓰인다(행동의향 아웃컴 3편 + 자택 앰비소닉 재생 1편).\n")
    open(os.path.join(FT, "table1_v2.md"), "w", encoding="utf-8").write("".join(L))

    print(f"  구 Table 1 재사용 매칭 {matched}/84행")
    print(f"[완료] {len(rows)}행 · 갈래 {dict(Counter(r['source'] for r in rows))}")
    print(f"  setting {dict(Counter(r['setting'] for r in rows).most_common(5))}")
    print(f"  design  {dict(Counter(r['design'] for r in rows).most_common(5))}")
    print(f"  quality {dict(Counter(r['quality'] for r in rows))}")
    print(f"  measure_gen {dict(Counter(r['measure_gen'] for r in rows).most_common(6))}")
    print(f"  MA 기여 {sum(1 for r in rows if r['in_ma'])}편")
    print("[저장] table1_v2.csv · table1_v2.md")


if __name__ == "__main__":
    main()
