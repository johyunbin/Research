# -*- coding: utf-8 -*-
"""
Paper32 — 최종 통합: UNCERTAIN 재판정 반영 + 효과크기 테이블 집계
① ua_01~04 verdict를 ft_verdicts_final에 반영 → ft_verdicts_v2.csv (100편 확정)
② ua extract를 ft_extraction_all에 반영(갱신) → ft_extraction_v2.csv
③ es_01~10 통합 → effect_sizes_all.csv + 클러스터별 MA 가능성 집계
"""
import sys, csv, os, glob
from collections import Counter, defaultdict

sys.stdout.reconfigure(encoding="utf-8")
FT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fulltext")
RES = os.path.join(FT, "ft_results")


def main():
    # ① verdict 반영
    rows = list(csv.DictReader(open(os.path.join(FT, "ft_verdicts_final.csv"), encoding="utf-8-sig")))
    ua = {}
    for p in sorted(glob.glob(os.path.join(RES, "ua_*_verdict.csv"))):
        for r in csv.DictReader(open(p, encoding="utf-8-sig")):
            ua[int(str(r["no"]).strip())] = r
    changed = 0
    for r in rows:
        no = int(r["no"])
        if no in ua:
            r["final_verdict"] = ua[no]["final_verdict"].strip().upper()
            r["reason"] = "전문재판정: " + ua[no].get("reason", "").strip()
            r["confidence"] = ua[no].get("confidence", "")
            changed += 1
    remaining_unc = [int(r["no"]) for r in rows if r["final_verdict"] == "UNCERTAIN"]
    with open(os.path.join(FT, "ft_verdicts_v2.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader(); w.writerows(rows)
    c = Counter(r["final_verdict"] for r in rows)
    print(f"① 재판정 반영 {changed}건 · 잔존 UNCERTAIN {remaining_unc}")
    print(f"   확정 분포: {dict(c)}")

    # ② extraction 갱신
    ex = {int(r["no"]): r for r in csv.DictReader(open(os.path.join(FT, "ft_extraction_all.csv"), encoding="utf-8-sig"))}
    for p in sorted(glob.glob(os.path.join(RES, "ua_*_extract.csv"))):
        for r in csv.DictReader(open(p, encoding="utf-8-sig")):
            no = int(str(r["no"]).strip())
            base = ex.get(no, {})
            base.update({k: v for k, v in r.items() if v})
            base["no"] = str(no)
            ex[no] = base
    vmap = {int(r["no"]): r["final_verdict"] for r in rows}
    keep = {n for n, v in vmap.items() if v in ("FINAL_INCLUDE", "SENS_ONLY")}
    cols = ["no", "final_verdict", "year", "journal", "title", "country", "setting", "design",
            "sample_n", "exposure", "behaviour_domain", "behaviour_measure",
            "measurement_method", "direction", "key_finding", "effect_stats"]
    meta = {int(r["no"]): r for r in csv.DictReader(open(os.path.join(FT, "packet_index.csv"), encoding="utf-8-sig"))}
    with open(os.path.join(FT, "ft_extraction_v2.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f); w.writerow(cols)
        for no in sorted(keep):
            e, m = ex.get(no, {}), meta.get(no, {})
            w.writerow([no, vmap[no], m.get("year", ""), m.get("journal", ""), m.get("title", "")]
                       + [e.get(k, "") for k in cols[5:]])
    print(f"② 추출표 v2: {len(keep)}행 (INCLUDE {sum(1 for n in keep if vmap[n]=='FINAL_INCLUDE')} + SENS {sum(1 for n in keep if vmap[n]=='SENS_ONLY')})")

    # ③ 효과크기 통합
    es_rows = []
    for p in sorted(glob.glob(os.path.join(RES, "es_*.csv"))):
        for r in csv.DictReader(open(p, encoding="utf-8-sig")):
            es_rows.append(r)
    with open(os.path.join(FT, "effect_sizes_all.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["no", "cluster", "outcome_measure", "comparison",
                                          "statistic_type", "values_verbatim", "n_info",
                                          "location", "quote", "computable"])
        w.writeheader()
        for r in es_rows:
            w.writerow({k: r.get(k, "") for k in w.fieldnames})
    yes = [r for r in es_rows if (r.get("computable") or "").strip().upper().startswith("Y")]
    covered = {int(str(r["no"]).strip()) for r in es_rows}
    print(f"③ 효과크기 행 {len(es_rows)} (논문 {len(covered)}편) · computable YES {len(yes)}")

    # 클러스터별: computable YES 기준 '서로 다른 논문 수'
    clus = defaultdict(set)
    for r in yes:
        c0 = (r.get("cluster") or "").strip().lower()
        no = int(str(r["no"]).strip())
        if vmap.get(no) != "FINAL_INCLUDE":
            continue
        for key in ("movement", "staying", "social", "space-use", "reverse-acoustic"):
            if key in c0:
                clus[key].add(no)
    print("   [MA 가능성 — computable YES·FINAL_INCLUDE 기준 논문 수]")
    for k in ("movement", "staying", "social", "space-use", "reverse-acoustic"):
        ids = sorted(clus[k])
        flag = "✅MA 가능(≥3)" if len(ids) >= 3 else "⚠️부족"
        print(f"   {k:16} {len(ids):>2}편 {flag}  {ids}")


if __name__ == "__main__":
    main()
