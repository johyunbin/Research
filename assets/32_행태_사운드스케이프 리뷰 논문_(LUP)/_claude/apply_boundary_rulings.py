# -*- coding: utf-8 -*-
"""
Paper32 — 사용자 경계판정 3건 적용 (2026-08-03 승인)
 R1 주거소음×비특정공간 PA        → FINAL_EXCLUDE (세팅 기준 위반)
 R2 행동 의향(LoW·PEB 등 의도)    → SENS_ONLY (본분석 제외·민감도 분석 편입)
 R3 역방향 & 아웃컴=지각/평가      → FINAL_INCLUDE (등록 역방향 정의 부합)
그 외 UNCERTAIN 은 유지(원문 재판정 대상). 출력: ft_verdicts_final.csv + ruling_audit.csv
"""
import sys, csv, os
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8")
FT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fulltext")


def main():
    ex_dir = {int(r["no"]): (r.get("direction") or "").strip().lower()
              for r in csv.DictReader(open(os.path.join(FT, "ft_extraction_all.csv"), encoding="utf-8-sig"))}
    rows = list(csv.DictReader(open(os.path.join(FT, "ft_verdicts_all.csv"), encoding="utf-8-sig")))

    audit = []
    for r in rows:
        if r["final_verdict"] != "UNCERTAIN":
            continue
        no = int(r["no"])
        reason = r.get("reason", "")
        d = ex_dir.get(no, "")
        rule, new = "", None
        if any(k in reason for k in ("주거", "근린", "주소")) and any(k in reason for k in ("PA", "신체활동", "활동 아웃컴", "physical")):
            rule, new = "R1", ("FINAL_EXCLUDE", "경계판정① 주거노출·공간비특정 PA — 세팅 기준 위반")
        elif any(k in reason for k in ("의향", "의도", "LoW", "PEB", "willingness", "ERB")):
            rule, new = "R2", ("SENS_ONLY", "경계판정② 행동의향 아웃컴 — 민감도 분석 전용")
        elif d.startswith("reverse") and any(k in reason for k in ("지각", "평가", "PAQ", "comfort", "dominance")):
            rule, new = "R3", ("FINAL_INCLUDE", "경계판정③ 역방향·지각/평가 아웃컴 — 등록 정의 부합")
        if new:
            audit.append({"no": no, "rule": rule, "old": "UNCERTAIN", "new": new[0],
                          "orig_reason": reason[:80], "title": r["title"][:70]})
            r["final_verdict"], r["reason"] = new[0], new[1]

    out = os.path.join(FT, "ft_verdicts_final.csv")
    with open(out, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader(); w.writerows(rows)
    with open(os.path.join(FT, "ruling_audit.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["no", "rule", "old", "new", "orig_reason", "title"])
        w.writeheader(); w.writerows(audit)

    c = Counter(r["final_verdict"] for r in rows)
    print("판정 변경:", Counter(a["rule"] for a in audit))
    for a in audit:
        print(f"  [{a['rule']}] ID {a['no']:>4} → {a['new']:14} | {a['title']}")
    print("\n최종 분포:", dict(c))
    print(f"저장: {out} · ruling_audit.csv")


if __name__ == "__main__":
    main()
