# -*- coding: utf-8 -*-
"""
Paper32 — MMAT 품질평가 통합·검증
※범주(mmat_category)는 전문을 읽은 에이전트 판정을 정본으로 채택
  (사전 mmat_assignment.csv는 design 문자열 휴리스틱 → 재배정 다수 발생, 이탈 로그 기록).
출력: fulltext/quality_all.csv · quality_detail_all.csv · quality_summary.md
"""
import sys, csv, os, glob
from collections import Counter, defaultdict

sys.stdout.reconfigure(encoding="utf-8")
FT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fulltext")
QA = os.path.join(FT, "qa_results")


def main():
    v = {int(r["no"]): r for r in csv.DictReader(open(os.path.join(FT, "ft_verdicts_v2.csv"), encoding="utf-8-sig"))}
    meta = {int(r["no"]): r for r in csv.DictReader(open(os.path.join(FT, "packet_index.csv"), encoding="utf-8-sig"))}
    target = {n for n, r in v.items() if r["final_verdict"] in ("FINAL_INCLUDE", "SENS_ONLY")}

    rows, problems = {}, []
    for p in sorted(glob.glob(os.path.join(QA, "qa_*_mmat.csv"))):
        for r in csv.DictReader(open(p, encoding="utf-8-sig")):
            n = int(str(r["no"]).strip())
            if n in rows:
                problems.append(f"중복 {n}")
            ny = sum(1 for k in ("Q1", "Q2", "Q3", "Q4", "Q5") if (r.get(k) or "").strip().upper() == "Y")
            declared = int(r["n_yes"]) if str(r.get("n_yes", "")).strip().isdigit() else -1
            if declared != ny:
                problems.append(f"n_yes 불일치 ID {n}: 선언 {declared} vs 재계산 {ny}")
            tier = "high" if ny >= 4 else ("moderate" if ny == 3 else "low")
            if (r.get("quality_tier") or "").strip().lower() != tier:
                problems.append(f"tier 불일치 ID {n}")
            rows[n] = {"no": n, "mmat_category": (r.get("mmat_category") or "").strip(),
                       "S1": r.get("S1", ""), "S2": r.get("S2", ""),
                       **{k: (r.get(k) or "").strip().upper() for k in ("Q1", "Q2", "Q3", "Q4", "Q5")},
                       "n_yes": ny, "quality_tier": tier, "note": (r.get("note") or "").strip(),
                       "year": meta.get(n, {}).get("year", ""), "journal": meta.get(n, {}).get("journal", ""),
                       "title": meta.get(n, {}).get("title", "")}
    missing = sorted(target - set(rows))
    extra = sorted(set(rows) - target)
    if missing: problems.append(f"평가 누락 {len(missing)}건: {missing}")
    if extra: problems.append(f"대상 밖 {extra}")

    det = []
    for p in sorted(glob.glob(os.path.join(QA, "qa_*_detail.csv"))):
        for r in csv.DictReader(open(p, encoding="utf-8-sig")):
            det.append({k: r.get(k, "") for k in ("no", "item", "verdict", "rationale")})

    print("=== 검증 ===")
    print("문제 없음" if not problems else "\n".join("⚠️ " + x for x in problems))

    with open(os.path.join(FT, "quality_all.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["no", "mmat_category", "S1", "S2", "Q1", "Q2", "Q3", "Q4", "Q5",
                                          "n_yes", "quality_tier", "note", "year", "journal", "title"])
        w.writeheader()
        for n in sorted(rows): w.writerow(rows[n])
    with open(os.path.join(FT, "quality_detail_all.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["no", "item", "verdict", "rationale"]); w.writeheader(); w.writerows(det)

    tier = Counter(r["quality_tier"] for r in rows.values())
    cat = Counter(r["mmat_category"] for r in rows.values())
    # 문항별 실패 집계
    itemfail = Counter()
    for d in det:
        if (d["verdict"] or "").strip().upper() in ("N", "CT"):
            itemfail[d["item"][:44]] += 1
    print(f"\n평가 완료 {len(rows)}편 · 상세 {len(det)}행")
    print("tier:", dict(tier))
    print("범주:", dict(cat))
    print("\n최다 감점 문항 top8:")
    for k, c in itemfail.most_common(8): print(f"   {c:>3}  {k}")

    # MA 기여 15편의 품질
    ma_ids = [53,461,481,494,532,617,661,665,746,804,901,941,951,1005,1076,1177,1178,1272,
              323,1280,14,492,539,732,930,931,1069,1122,1221,980,1018]
    ma_t = Counter(rows[n]["quality_tier"] for n in ma_ids if n in rows)
    print(f"\nMA 기여 연구군 품질: {dict(ma_t)}")

    with open(os.path.join(FT, "quality_summary.md"), "w", encoding="utf-8") as f:
        f.write("# Paper32 — MMAT 2018 품질평가 결과\n\n")
        f.write(f"- 평가 {len(rows)}편(FINAL_INCLUDE 81 + SENS_ONLY 3) · 문항 판정 {len(det)}건\n")
        f.write(f"- 등급: high(4~5) {tier['high']} · moderate(3) {tier['moderate']} · low(0~2) {tier['low']}\n")
        f.write(f"- MMAT 범주: {dict(cat)}\n\n## 최다 감점 문항\n\n")
        for k, c in itemfail.most_common(10): f.write(f"- {k} — {c}건\n")
        f.write(f"\n## 메타분석 기여 연구군 품질\n\n{dict(ma_t)}\n")
        if problems: f.write("\n## 검증 문제\n" + "\n".join("- " + x for x in problems) + "\n")
    print("\nquality_summary.md 저장")


if __name__ == "__main__":
    main()
