# -*- coding: utf-8 -*-
"""
Paper32 — Table 3: 메타분석 기여 연구 특성표 (Zhang et al. 2025 LUP Table 2 양식 준용)

사용자 요청(2026-08-16): "선행연구들을 첨부한 table 처럼 정리" — 연구별 국가·설계·
통계량·대비·표본·계수를 한 표에. 우리 코퍼스는 98편이라 전건은 S11 몫이고,
본문 표는 **네 클러스터에 실제로 합성된 효과의 기여 연구**(k=3+3+4+6)로 한다.

수치는 전부 정본에서 읽는다: 효과·분산 = figures/fig_data.json (ma_forest_data 경유),
국가·세팅·설계·품질 = fulltext/table1_v2.csv, 분석 n = ma/*_input.csv.
출력: fulltext/ma_char_table.md (원고 §3.4 에 붙여넣는 markdown)
"""
import sys, os, csv, json, math

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
MA = os.path.join(FT, "ma")
D = json.load(open(os.path.join(BASE, "figures", "fig_data.json"), encoding="utf-8"))

T1 = {r["uid"]: r for r in csv.DictReader(open(os.path.join(FT, "table1_v2.csv"),
                                               encoding="utf-8-sig"))}
# 국가 = 부록 B·Fig. 7 과 같은 원천(추출표 country)·같은 정규화. 종전에는 table1 의 첫 국가만 적어
# 다국가 연구(Cao & Kang 2021 = United Kingdom; China)가 부록 B 와 달랐다(독립 검증 적발, 2026-09-17).
sys.path.insert(0, BASE)
from make_fig7_geo_time import norm_countries   # noqa: E402
EXT_COUNTRY = {r["uid"]: r["country"] for r in csv.DictReader(
    open(os.path.join(FT, "corpus_v4_extraction.csv"), encoding="utf-8-sig"))}

# 분석 표본 n — ma 입력표·추출표에서 확정한 값 (효과 단위·클러스터 규약 반영 · 추정치 금지)
#   walking: n_pos+n_neg (532 는 Exp1+Exp2 논문 내 합성 = 58+44)
#   14 Moser: 열쇠 도움행동 Ro vs R+ 대비 = 6조건 450시행 중 2조건 × 75 = 150
#   CT0126: 원문 §2.2 명시 분석단위 = sampling point N=29 / CT0184: 회귀분석 N=301
#   1177: n 미보고 — 보고된 CI 에서 분산 역산(원고 §2.7 D7-3)
N_ANALYTIC = {
    "461": "57", "532": "102", "617": "54",
    "323": "596", "665": "97", "1280": "241",
    "931": "73", "1069": "146", "14": "150", "CT0025": "80",
    "1076": "419", "1221": "315", "1177": "NR", "980": "180",
    "CT0126": "29", "CT0184": "301",
}

# 대비·측정 짧은 영문 표기 (그림 라벨과 동일 어휘)
CONTRAST = {
    "461": "birdsong vs traffic noise", "532": "birdsong vs city noise (Exp 1 and 2)",
    "617": "natural vs traffic sound", "323": "music vs no music",
    "665": "music vs no sound", "1280": "natural sound index (high vs low)",
    "931": "natural vs noise (group interaction)", "1069": "natural vs noise (paired interaction)",
    "14": "quiet vs roadworks noise (helping)", "CT0025": "quiet vs lawnmower noise (helping)",
    "1076": "pleasantness with static behaviour", "1221": "natural sound events with queuing",
    "1177": "sound comfort with walking comfort", "980": "dwell time with perceived restoration",
    "CT0126": "LAeq with vocal effort", "CT0184": "companion presence with sound noticing",
}

CLUSTERS = [
    ("walking", "Walking speed (Hedges' g; negative = faster under noise)", "g"),
    ("staying", "Staying / dwell time (Hedges' g; positive = longer stay)", "g"),
    ("social", "Social interaction (Hedges' g; positive = more interaction)", "g"),
    ("correlation", "Sound–behaviour correlation (r; acoustic or perceptual measure, either direction)", "r"),
]


def short_study(uid, label):
    """그림 라벨(FOREST_LABEL) 형식과 동일한 '저자 연도' 표기."""
    from make_figures import FOREST_LABEL
    s = FOREST_LABEL.get(str(uid), label)
    return s.split("·")[0].strip()


def fmt_effect(e, kind):
    se = math.sqrt(e["var"])
    lo, hi = e["est"] - 1.96 * se, e["est"] + 1.96 * se
    if kind == "r":
        lo, hi, est = math.tanh(lo), math.tanh(hi), math.tanh(e["est"])
        return f"{est:+.2f} [{lo:+.2f}, {hi:+.2f}]"
    return f"{e['est']:+.2f} [{lo:+.2f}, {hi:+.2f}]"


def main():
    rows, missing = [], []
    no = 0
    for key, head, kind in CLUSTERS:
        rows.append(f"| **{head}** | | | | | | | |")
        for e in D["ma"][key]["effects"]:
            uid = str(e["uid"])
            t = T1.get(uid)
            if not t:
                missing.append(uid); continue
            no += 1
            study = short_study(uid, e.get("label", uid))
            if e.get("route") == "citation-tracking":
                study += " ▲"
            country = "; ".join(norm_countries(EXT_COUNTRY[uid])) or "NR"
            setting = t["setting"].split(";")[0].strip()
            setting = {"lab(outdoor scene)": "laboratory (outdoor scene)"}.get(setting, setting)
            design = t["design"].split(";")[0].strip()
            # 부록 B(Table B1)와 같은 표기 — 첫 글자 대문자(2026-09-17)
            design = {"observational": "observation"}.get(design, design)   # Table 1 "Field observation" 과 같은 말
            setting, design = setting[:1].upper() + setting[1:], design[:1].upper() + design[1:]
            eff = fmt_effect(e, kind).replace("-", "−")   # 표 전반과 동일한 typographic minus
            rows.append(f"| {study} | {country} | {setting} | {design} "
                        f"| {CONTRAST.get(uid, '')} | {N_ANALYTIC.get(uid, 'NR')} "
                        f"| {t['quality'].capitalize()} | {eff} |")

    md = ["| Study | Country | Setting | Design | Contrast or measure | *n* | MMAT | Effect [95% CI] |",
          "|---|---|---|---|---|---|---|---|"]
    # 클러스터 머리행은 9열에 맞춘다
    out = []
    for r in rows:
        if r.startswith("| **"):
            out.append("| " + r.strip("| ").split("|")[0].strip() + " | | | | | | | |")
        else:
            out.append(r)
    md += out
    text = "\n".join(md) + "\n"
    open(os.path.join(FT, "ma_char_table.md"), "w", encoding="utf-8").write(text)
    print(text)
    if missing:
        print(f"⚠️ table1 에 없는 uid: {missing}")
    print(f"[저장] fulltext/ma_char_table.md · 효과 {no}건")


if __name__ == "__main__":
    main()
