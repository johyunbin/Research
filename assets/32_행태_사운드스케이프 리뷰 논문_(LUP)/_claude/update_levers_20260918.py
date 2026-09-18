# -*- coding: utf-8 -*-
"""
Paper32 — 계획 레버 근거 연구 배정 갱신(9월 편입 16편 + 1226 배제) 및 수치 열 재계산 (2026-09-18)

배정 원칙(2.8절): 레버 = 설계자·운영자가 바꿀 수 있는 조치. 포함 연구(FINAL_INCLUDE)만 근거로 센다(민감도 전용 제외).
이용 행태가 사운드스케이프 평가를 바꾸는 역방향 평가 연구(80·104·499·713·CT0223·1305)는 조작 가능한 조치가 아니므로
레버에 배정하지 않는다. 137(야간 하이킹)은 활동 선택이 지각을 바꾸는 연구라 같은 이유로 배정하지 않는다.
n_studies·quality_mix 는 quality_v2.csv 에서 다시 계산한다(수기 금지). effect_summary 의 메타분석 값은 ma_forest_data.json 현행값.
"""
import csv, json, math, os, sys
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
DM = os.path.join(FT, "design_matrix_v2.csv")

ADD = {"L2": ["1191"], "L3": ["1190"], "L5": ["709"], "L6": ["CT0171"],
       "L7": ["OAS0011", "OAS0041", "1218", "888"], "L8": ["CT0391"]}
REMOVE = {"L5": ["1226"]}          # 2026-09-18 R1 재판정으로 코퍼스에서 배제

rows = list(csv.DictReader(open(DM, encoding="utf-8-sig")))
keys = list(rows[0].keys())
verd = {r["uid"]: r["final_verdict"] for r in csv.DictReader(open(os.path.join(FT, "corpus_v4_verdicts.csv"), encoding="utf-8-sig"))}
tier = {r["uid"]: r["quality_tier"] for r in csv.DictReader(open(os.path.join(FT, "quality_v2.csv"), encoding="utf-8-sig"))}
fd = json.load(open(os.path.join(FT, "ma", "ma_forest_data.json"), encoding="utf-8"))


def ma(name, r=False):
    p = fd[name]["pooled"]
    f = (lambda z: math.tanh(z)) if r else (lambda z: z)
    sym = "r" if r else "g"
    return (f"k={p['k']}, {sym}={f(p['est']):+.3f} (95% CI {f(p['lo']):+.3f}~{f(p['hi']):+.3f}, p={p['p']:.3f}"
            f"{', CI 0 포함' if p['lo'] <= 0 <= p['hi'] else ''})")


for r in rows:
    lid = r["lever_id"]
    ids = [x for x in r["evidence_studies"].split(";") if x]
    ids = [x for x in ids if x not in REMOVE.get(lid, [])] + [x for x in ADD.get(lid, []) if x not in ids]
    bad = [x for x in ids if verd.get(x) != "FINAL_INCLUDE"]
    if bad:
        sys.exit(f"{lid}: 포함 연구가 아닌 근거 {bad}")
    r["evidence_studies"] = ";".join(ids)
    r["n_studies"] = str(len(ids))
    c = Counter(tier[x] for x in ids)
    r["quality_mix"] = " · ".join(f"{k[:4] if k == 'moderate' else k} {c[k]}" for k in ("high", "moderate", "low") if c[k]).replace("mode", "mod")
    # 옛 메타분석 값(8월 REML 정정 전) → 현행
    es = r["effect_summary"]
    es = es.replace("MA2 k=3, g=+0.313 (95% CI -0.076~+0.702, p=0.074, CI 0 포함)", "MA2 " + ma("staying"))
    es = es.replace("MA3 k=4, g=+0.646 (95% CI +0.188~+1.104, p=0.021)", "MA3 " + ma("social"))
    es = es.replace("MA1 k=4 (3편·532는 2개 실험), g=-0.500 (95% CI -1.411~+0.412, p=0.179", "MA1 " + ma("walking").rstrip(")") + "")
    r["effect_summary"] = es
    if lid == "L5":
        r["caveat"] = r["caveat"].replace("1226·951은 부호가 반대(접근성 교락)", "951은 부호가 반대(접근성 교락)")

with open(DM, "w", newline="", encoding="utf-8-sig") as f:
    w = csv.DictWriter(f, fieldnames=keys)
    w.writeheader()
    w.writerows(rows)
for r in rows:
    print(r["lever_id"], r["n_studies"], r["quality_mix"], r["confidence"], "|", r["effect_summary"][:90])
