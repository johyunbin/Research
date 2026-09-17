# -*- coding: utf-8 -*-
"""
Paper32 — 추가 전문 평가(nft_XX) 결과 검사·통합 (2026-09-18)

검사(지시문 newft_instructions_20260918.md §3·§4):
  - 배치별 verdict 행 = 배정 레코드, id 누락·중복 없음, 어휘(verdict·reason_code·boundary_rule·confidence)
  - INCLUDE/SENS 는 extract·mmat 에 있고 detail 은 id 당 5행, item_no 첫 자리 = mmat_category, n_yes·tier 재계산 일치
  - 통제어휘(setting·design 선두어, behaviour_domain, measurement_method, direction)
  - 경계 규칙 P1(branch 별 처리)·R2(SENS_ONLY) 정합
출력: fulltext/newft_combined/{verdict,extract,mmat,detail,es}.csv + 요약 출력. 오류가 있으면 목록을 출력하고 exit 1.
사용: python validate_newft_results.py [--partial]   (--partial: 아직 안 끝난 배치는 건너뜀)
"""
import csv, json, os, re, sys
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
RES = os.path.join(FT, "newft_results")
OUT = os.path.join(FT, "newft_combined")
PARTIAL = "--partial" in sys.argv

VERDICTS = {"FINAL_INCLUDE", "SENS_ONLY", "FINAL_EXCLUDE"}
REASONS = {"", "X1", "X2", "X3", "X4", "X5", "X6", "X7"}
RULES = {"", "R1", "R2", "R3", "P1", "P3"}
CONF = {"high", "medium", "low"}
SETTING = ("street", "park", "square", "campus", "residential", "recreation", "waterfront", "VR-lab", "mixed")
DESIGN = ("survey", "mixed", "field-experiment", "natural-experiment", "quasi-experiment", "field-observation",
          "observational", "sensor-bigdata", "lab-VR-experiment", "qualitative", "NR")
DOMAINS = {"movement", "staying", "space-use", "social", "activity"}
METHODS = ("self-report", "observation", "sensor-GPS-video")
CATS = {"Qualitative": "1", "Quantitative RCT": "2", "Quantitative non-randomised": "3",
        "Quantitative descriptive": "4", "Mixed methods": "5"}
CLUSTERS = {"walking_speed", "staying", "social", "correlation", "other"}


def read(path):
    return list(csv.DictReader(open(path, encoding="utf-8-sig"))) if os.path.exists(path) else None


def lead(value, vocab):
    v = (value or "").strip()
    return any(v == w or v.startswith(w + " ") or v.startswith(w + "(") for w in vocab)


def main():
    batches = json.load(open(os.path.join(FT, "newft_batches_20260918.json"), encoding="utf-8"))
    manifest = {r["id"]: r for r in csv.DictReader(open(os.path.join(FT, "newft_manifest_20260918.csv"), encoding="utf-8-sig"))}
    errs, combined = [], {k: [] for k in ("verdict", "extract", "mmat", "detail", "es")}
    done = 0
    for b in batches:
        name, ids = b["batch"], b["ids"]
        files = {k: read(os.path.join(RES, f"{name}_{k}.csv")) for k in combined}
        if files["verdict"] is None:
            if not PARTIAL:
                errs.append(f"{name}: verdict 파일 없음")
            continue
        done += 1
        missing_files = [k for k, v in files.items() if v is None]
        if missing_files:
            errs.append(f"{name}: 파일 없음 {missing_files}")
            continue
        V = files["verdict"]
        vid = [r["id"] for r in V]
        if sorted(vid) != sorted(ids) or len(set(vid)) != len(vid):
            errs.append(f"{name}: verdict id 불일치 {sorted(vid)} vs {sorted(ids)}")
        inc = set()
        for r in V:
            i = r["id"]
            if r["verdict"] not in VERDICTS: errs.append(f"{name} {i}: verdict {r['verdict']!r}")
            if r["reason_code"] not in REASONS: errs.append(f"{name} {i}: reason_code {r['reason_code']!r}")
            if r["boundary_rule"] not in RULES: errs.append(f"{name} {i}: boundary_rule {r['boundary_rule']!r}")
            if r["confidence"] not in CONF: errs.append(f"{name} {i}: confidence {r['confidence']!r}")
            if r["verdict"] == "FINAL_EXCLUDE" and not r["reason_code"]: errs.append(f"{name} {i}: 배제인데 reason_code 없음")
            if r["verdict"] != "FINAL_EXCLUDE" and r["reason_code"]: errs.append(f"{name} {i}: 포함인데 reason_code 있음")
            if r["branch"] != manifest[i]["branch"]: errs.append(f"{name} {i}: branch {r['branch']} ≠ {manifest[i]['branch']}")
            if r["boundary_rule"] == "P1":
                want = "FINAL_INCLUDE" if manifest[i]["branch"] == "DB" else "SENS_ONLY"
                if r["verdict"] != want: errs.append(f"{name} {i}: P1 인데 {r['verdict']} (branch {manifest[i]['branch']} → {want})")
            if r["boundary_rule"] == "R2" and r["verdict"] != "SENS_ONLY": errs.append(f"{name} {i}: R2 인데 {r['verdict']}")
            if not r["rationale"].strip(): errs.append(f"{name} {i}: rationale 없음")
            if r["verdict"] != "FINAL_EXCLUDE":
                inc.add(i)
        E = {r["id"]: r for r in files["extract"]}
        M = {r["id"]: r for r in files["mmat"]}
        D = Counter(r["id"] for r in files["detail"])
        if set(E) != inc: errs.append(f"{name}: extract id {sorted(E)} ≠ 포함 {sorted(inc)}")
        if set(M) != inc: errs.append(f"{name}: mmat id {sorted(M)} ≠ 포함 {sorted(inc)}")
        for i in inc:
            e = E.get(i)
            if e:
                if e["final_verdict"] != next(r["verdict"] for r in V if r["id"] == i): errs.append(f"{name} {i}: extract final_verdict 불일치")
                if not lead(e["setting"], SETTING): errs.append(f"{name} {i}: setting {e['setting'][:30]!r}")
                if not lead(e["design"], DESIGN): errs.append(f"{name} {i}: design {e['design'][:30]!r}")
                doms = {d.strip() for d in e["behaviour_domain"].split(";") if d.strip()}
                if not doms or doms - DOMAINS: errs.append(f"{name} {i}: behaviour_domain {e['behaviour_domain']!r}")
                meth = [m.strip() for m in re.sub(r"\([^)]*\)", "", e["measurement_method"]).split(";") if m.strip()]
                if not meth or any(m not in METHODS for m in meth): errs.append(f"{name} {i}: measurement_method {e['measurement_method'][:40]!r}")
                if e["direction"] not in ("forward", "reverse", "both"): errs.append(f"{name} {i}: direction {e['direction']!r}")
            m = M.get(i)
            if m:
                cat = CATS.get(m["mmat_category"])
                if not cat: errs.append(f"{name} {i}: mmat_category {m['mmat_category']!r}"); continue
                qs = [m[f"Q{k}"] for k in range(1, 6)]
                if any(q not in ("Y", "N", "CT") for q in qs + [m["S1"], m["S2"]]): errs.append(f"{name} {i}: MMAT 값 형식")
                ny = sum(q == "Y" for q in qs)
                tier = "high" if ny >= 4 else "moderate" if ny == 3 else "low"
                if str(ny) != m["n_yes"].strip() or tier != m["quality_tier"]: errs.append(f"{name} {i}: n_yes/tier 재계산 불일치 ({ny},{tier})")
                if D[i] != 5: errs.append(f"{name} {i}: detail {D[i]}행")
                dets = [r for r in files["detail"] if r["id"] == i]
                if any(not r["item_no"].startswith(cat + ".") for r in dets): errs.append(f"{name} {i}: detail item_no 범주 불일치")
                if [r["verdict"] for r in sorted(dets, key=lambda r: r["item_no"])] != qs: errs.append(f"{name} {i}: detail 판정 ≠ Q1–Q5")
        for r in files["es"]:
            if r["id"] not in inc: errs.append(f"{name} {r['id']}: es 인데 포함 아님")
            if r["cluster"] not in CLUSTERS: errs.append(f"{name} {r['id']}: cluster {r['cluster']!r}")
        for k in combined:
            for r in files[k]:
                combined[k].append({"batch": name, **r})

    os.makedirs(OUT, exist_ok=True)
    for k, rows in combined.items():
        if rows:
            with open(os.path.join(OUT, f"{k}.csv"), "w", newline="", encoding="utf-8-sig") as f:
                keys = list(dict.fromkeys(key for r in rows for key in r))
                w = csv.DictWriter(f, fieldnames=keys)
                w.writeheader()
                w.writerows(rows)
    V = combined["verdict"]
    print(f"완료 배치 {done}/{len(batches)} · 레코드 {len(V)}/{len(manifest)}")
    print("판정:", Counter((r["branch"], r["verdict"]) for r in V))
    print("배제 사유:", Counter(r["reason_code"] for r in V if r["reason_code"]))
    print("경계 규칙:", Counter(r["boundary_rule"] for r in V if r["boundary_rule"]), "· low confidence:",
          [r["id"] for r in V if r["confidence"] == "low"])
    print("MMAT:", Counter(r["quality_tier"] for r in combined["mmat"]), "· 효과 후보:", Counter(r["cluster"] for r in combined["es"]))
    if errs:
        print(f"\n오류 {len(errs)}건:")
        print("\n".join("  " + e for e in errs))
        sys.exit(1)
    print("검사 통과")


if __name__ == "__main__":
    main()
