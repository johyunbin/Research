# -*- coding: utf-8 -*-
"""
Paper32 — 원고 본문용 Table 1·Table 3 영문 생성
본문이 두 표를 인용하는데 실물이 없었다(독립 점검에서 발견).
Table 1 = 연구 특성 요약(전건 목록은 S11) · Table 3 = 계획 레버 10개
출력: fulltext/table1_en.md · table3_en.md
"""
import sys, os, csv, re
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
sys.path.insert(0, BASE)
from make_fig7_geo_time import norm_countries


def rd(p):
    return list(csv.DictReader(open(os.path.join(FT, p), encoding="utf-8-sig")))


def wrap(s, n=118):
    return re.sub(r"\s+", " ", (s or "")).strip()[:n]


def table1():
    ext = rd("corpus_v4_extraction.csv")
    verd = {r["uid"]: r["final_verdict"] for r in rd("corpus_v4_verdicts.csv")}
    qual = {r["uid"]: r["quality_tier"] for r in rd("quality_v2.csv")}
    t1 = {r["uid"]: r for r in rd("table1_v2.csv")}
    inc = [r for r in ext if verd.get(r["uid"]) == "FINAL_INCLUDE"]
    N = len(inc)

    DES = [("Field or natural experiment", r"^field[\s-]?experiment|^natural[\s-]?experiment"),
           ("Laboratory / VR experiment", r"^lab|^vr\b|virtual"),
           ("Mixed observation + survey", r"^mixed"),
           ("Survey", r"^survey|^questionnaire"),
           ("Field observation", r"^field[\s-]?observation|^observation"),
           ("Sensor / big data", r"^sensor|^big[\s-]?data")]
    SET = [("Street", r"street|road|sidewalk|pavement"), ("Park", r"park|garden|green space"),
           ("Square / plaza", r"square|plaza|piazza"), ("Campus", r"campus|university|school"),
           ("Residential open space", r"residential|neighbourhood|neighborhood|housing"),
           ("Waterfront", r"waterfront|river|lake|coast|harbour|harbor")]

    def bucket(v, pats, head=False):
        s = (v or "").lower()
        if head:
            s = s.split("(")[0].strip()
        return [nm for nm, p in pats if re.search(p, s)]

    des, setc, ctry, per, gen, dirn = Counter(), Counter(), Counter(), Counter(), Counter(), Counter()
    q = Counter()
    for r in inc:
        hit = bucket(r["design"], DES, head=True)
        des[hit[0] if hit else "Other / not reported"] += 1
        for b in bucket(r["setting"], SET) or ["Other"]:
            setc[b] += 1
        cs = norm_countries(r["country"])
        for c in (cs or ["Not reported"]):
            ctry[c] += 1
        try:
            y = int(r["year"])
            per["≤2009" if y < 2010 else ("2010–2019" if y < 2020 else "2020–2026")] += 1
        except (TypeError, ValueError):
            pass
        for g in {x.strip() for x in (t1.get(r["uid"], {}).get("measure_gen", "") or "").split(";")}:
            if g in ("G1", "G2", "G3"):
                gen[g] += 1
        dirn[(r["direction"] or "").strip()] += 1
        q[qual.get(r["uid"], "?")] += 1

    def blk(title, counter, order=None, note=""):
        keys = order or [k for k, _ in counter.most_common()]
        out = [f"| **{title}**{note} | | |\n"]
        for k in keys:
            if counter[k]:
                out.append(f"| {k} | {counter[k]} | {counter[k]/N*100:.0f}% |\n")
        return out

    L = [f"**Table 1.** Characteristics of the {N} included studies. Percentages are of {N} studies; "
         "categories marked † allow a study to appear more than once. The full study-level listing "
         "(identifier, route, year, journal, title, country, setting, design, sample, exposure, "
         "behavioural domain, measurement generation, direction, quality) is Supplementary S11.\n\n",
         "| Characteristic | Studies | % |\n|---|---|---|\n"]
    L += blk("Publication period", per, ["≤2009", "2010–2019", "2020–2026"])
    L += blk("Study design", des)
    L += blk("Setting", setc, note=" †")
    top = [k for k, v in ctry.most_common(8)]
    L += blk("Country", ctry, top, note=" †")
    L.append(f"| Other countries (n = {len(ctry) - len(top)}) | "
             f"{sum(v for k, v in ctry.items() if k not in top)} | — |\n")
    L += blk("Direction of relationship", dirn, ["forward", "reverse", "both"])
    L += blk("Behavioural measurement generation", gen, ["G1", "G2", "G3"], note=" †")
    L.append("| **Methodological quality (MMAT 2018)** | | |\n")
    for k in ("high", "moderate", "low"):
        L.append(f"| {k.capitalize()} | {q[k]} | {q[k]/N*100:.0f}% |\n")
    L.append("\nG1 = self-report; G2 = systematic observation; G3 = sensing, GPS, video or "
             "big-data measurement of behaviour.\n")
    open(os.path.join(FT, "table1_en.md"), "w", encoding="utf-8").write("".join(L))
    print(f"[Table 1] {N}편 · 설계 {len(des)}종 · 국가 {len(ctry)}종 · 품질 {dict(q)}")


# design_matrix_v2.csv 의 mechanism·behaviour_outcome 은 한국어다. 영문 원고용 대응문.
# 원문을 축약·의역하지 않고 같은 내용을 옮긴 것이며, 원본은 S14 에 그대로 남는다.
LEVER_EN = {
    "L1": ("Sound draws attention, pulls people towards the source and slows them, extending stay",
           "Staying (dwell time); space use (approach to source); movement (slower wandering)"),
    "L2": ("Natural sound raises perceived restoration and safety, encouraging talk and lingering",
           "Social interaction; staying (dwell time); vitality of activity"),
    "L3": ("Noise avoidance shifts route and mode choice, moving trips onto quieter alignments",
           "Movement (route choice, mode choice, cycling volume)"),
    "L4": ("Unpleasant machinery noise creates avoidance routes and acceleration, cutting stay and talk",
           "Social interaction; staying; space use (removal of avoidance routes)"),
    "L5": ("A quiet façade or designated zone lowers the psychological barrier to outdoor stay and walking",
           "Activity (walking, exercise, rest); space use; staying"),
    "L6": ("Separating and enclosing noise-generating functions changes the density and interaction of adjacent activity",
           "Space use (crowd density); social interaction (frequency, duration); staying"),
    "L7": ("Human sound and activity attract watching and joining, converting passage into stay",
           "Staying (watching, lingering); social interaction; space use"),
    "L8": ("Directional signal sound directly adjusts crossing trajectory and response timing",
           "Movement (crossing accuracy, detection timing, smoothness of deceleration)"),
    "L9": ("Noise provokes avoidance and speeds passage; natural sound is assumed to reverse it",
           "Movement (walking speed)"),
    "L10": ("Noise interferes with speech, forcing conversation to stop or voices to rise, and "
            "eventually deterring verbal interaction altogether",
            "Social interaction (conversation duration, vocal effort); acceptance of verbal interaction"),
}

CONF_NOTE = {
    "moderate": "Include in a design proposal, with post-occupancy monitoring",
    "low": "Test as a hypothesis; do not write into a standard or guideline",
    "very low": "No prescriptive basis at present",
}


def table3():
    rows = rd("design_matrix_v2.csv")
    L = ["**Table 3.** Ten planning levers derived from the corpus, with confidence grading. "
         "Confidence combines the number of contributing studies, their MMAT composition, whether "
         "a pooled interval excludes zero, diversity of settings, and behaviour under sensitivity "
         "analysis. **No lever reaches high confidence.** `Mixed` under Direction means that studies "
         "disagree in sign, not that the lever has several effects. Full evidence, caveats and "
         "study lists are Supplementary S14.\n\n",
         "| Lever | What is changed | Behavioural outcome | Direction | Studies | MMAT mix | "
         "Confidence |\n|---|---|---|---|---|---|---|\n"]
    conf = Counter()
    missing = [r["lever_id"] for r in rows if r["lever_id"] not in LEVER_EN]
    if missing:
        raise SystemExit(f"⚠️ 영문 대응문 없는 레버: {missing} — LEVER_EN 보완 필요")
    for r in rows:
        conf[r["confidence"]] += 1
        mech, outc = LEVER_EN[r["lever_id"]]
        L.append(f"| **{r['lever_id']}** {wrap(r['lever_en'], 52)} | {mech} | {outc} | "
                 f"{r['direction']} | {r['n_studies']} | "
                 f"{wrap(r['quality_mix'], 26)} | **{r['confidence']}** |\n")
    L.append("\n**How to read the grades.** "
             + " · ".join(f"*{k}* ({conf[k]}) — {v}" for k, v in CONF_NOTE.items() if conf[k])
             + ".\n")
    open(os.path.join(FT, "table3_en.md"), "w", encoding="utf-8").write("".join(L))
    print(f"[Table 3] 레버 {len(rows)}개 · confidence {dict(conf)}")


if __name__ == "__main__":
    table1()
    table3()
    print("[저장] fulltext/table1_en.md · table3_en.md")
