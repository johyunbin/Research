# -*- coding: utf-8 -*-
"""
Paper32 — 전문심사 결과 통합·검증
입력: fulltext/ft_results/batch_NN_verdict.csv + batch_NN_extract.csv
검증: 51 ID 전수·중복 없음·verdict 유효·extract 집합 정합
출력: ft_verdicts_all.csv · ft_extraction_all.csv · ft_summary.md · uncertain_list.csv
"""
import sys, csv, os, glob
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
RES = os.path.join(FT, "ft_results")
VALID = {"FINAL_INCLUDE", "FINAL_EXCLUDE", "UNCERTAIN"}
EX_COLS = ["no", "country", "setting", "design", "sample_n", "exposure", "behaviour_domain",
           "behaviour_measure", "measurement_method", "direction", "key_finding", "effect_stats"]


def main():
    meta = {int(r["no"]): r for r in csv.DictReader(
        open(os.path.join(FT, "packet_index.csv"), encoding="utf-8-sig"))}

    verdicts, extracts, problems = {}, {}, []
    for pat in ["batch_%02d" % i for i in range(1,7)] + ["b2_%02d" % i for i in range(1,7)]:
        vp = os.path.join(RES, f"{pat}_verdict.csv")
        ep = os.path.join(RES, f"{pat}_extract.csv")
        if not os.path.exists(vp):
            problems.append(f"{pat}_verdict.csv 없음"); continue
        for r in csv.DictReader(open(vp, encoding="utf-8-sig")):
            rid = int(str(r["no"]).strip())
            v = (r.get("final_verdict") or "").strip().upper()
            if rid in verdicts:
                problems.append(f"중복 verdict ID {rid} ({pat})")
            if v not in VALID:
                problems.append(f"무효 verdict '{v}' ID {rid}")
            verdicts[rid] = {"verdict": v, "reason": (r.get("exclude_reason") or "").strip(),
                             "conf": (r.get("confidence") or "").strip(), "batch": pat}
        if os.path.exists(ep):
            for r in csv.DictReader(open(ep, encoding="utf-8-sig")):
                try:
                    extracts[int(str(r["no"]).strip())] = r
                except ValueError:
                    problems.append(f"{pat}_extract 잘못된 no")

    missing = [k for k in meta if k not in verdicts]
    if missing:
        problems.append(f"판정 누락 {len(missing)}건: {missing}")
    keep = {k for k, v in verdicts.items() if v["verdict"] in ("FINAL_INCLUDE", "UNCERTAIN")}
    ex_missing = keep - set(extracts)
    if ex_missing:
        problems.append(f"추출 누락 {len(ex_missing)}건: {sorted(ex_missing)}")

    print("=== 검증 ===")
    print("문제 없음" if not problems else "\n".join("⚠️ " + p for p in problems))

    # 통합 verdict
    with open(os.path.join(FT, "ft_verdicts_all.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["no", "final_verdict", "confidence", "reason", "screening_verdict", "year", "journal", "title"])
        for rid in sorted(verdicts):
            v, m = verdicts[rid], meta.get(rid, {})
            w.writerow([rid, v["verdict"], v["conf"], v["reason"],
                        m.get("verdict", ""), m.get("year", ""), m.get("journal", ""), m.get("title", "")])

    # 통합 extraction (+ 서지 붙임)
    with open(os.path.join(FT, "ft_extraction_all.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["no", "final_verdict", "year", "journal", "title"] + EX_COLS[1:])
        for rid in sorted(extracts):
            e, m = extracts[rid], meta.get(rid, {})
            w.writerow([rid, verdicts.get(rid, {}).get("verdict", ""), m.get("year", ""),
                        m.get("journal", ""), m.get("title", "")] + [e.get(c, "") for c in EX_COLS[1:]])

    # UNCERTAIN 목록(사용자 판정용)
    with open(os.path.join(FT, "uncertain_list.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["no", "reason", "year", "journal", "title"])
        for rid in sorted(verdicts):
            if verdicts[rid]["verdict"] == "UNCERTAIN":
                m = meta.get(rid, {})
                w.writerow([rid, verdicts[rid]["reason"], m.get("year", ""), m.get("journal", ""), m.get("title", "")])

    c = Counter(v["verdict"] for v in verdicts.values())
    dirs = Counter((extracts[k].get("direction") or "").strip().lower() for k in extracts)
    doms = Counter()
    for k in extracts:
        for d in (extracts[k].get("behaviour_domain") or "").replace(";", ",").split(","):
            d = d.strip().lower()
            if d and d != "nr":
                doms[d] += 1
    meth = Counter((extracts[k].get("measurement_method") or "").strip().lower() for k in extracts)
    eff = sum(1 for k in extracts if (extracts[k].get("effect_stats") or "").strip().upper() not in ("", "NR"))

    print("\n=== 전문심사 결과 ===")
    for k in ["FINAL_INCLUDE", "UNCERTAIN", "FINAL_EXCLUDE"]:
        print(f"  {k}: {c.get(k,0)}")
    print(f"  추출 행: {len(extracts)} · 효과크기 원문 보유: {eff}")
    print(f"  방향: {dict(dirs)}")
    print(f"  행태 도메인: {dict(doms.most_common())}")
    print(f"  측정방법: {dict(meth.most_common(6))}")

    with open(os.path.join(FT, "ft_summary.md"), "w", encoding="utf-8") as f:
        f.write("# Paper32 전문심사 결과 (1·2차 통합 — 확보 PDF 100편 기준)\n\n")
        f.write(f"- FINAL_INCLUDE {c.get('FINAL_INCLUDE',0)} · UNCERTAIN {c.get('UNCERTAIN',0)} · FINAL_EXCLUDE {c.get('FINAL_EXCLUDE',0)}\n")
        f.write(f"- 추출 {len(extracts)}건 · 효과크기 원문 {eff}건\n- 방향: {dict(dirs)}\n- 도메인: {dict(doms.most_common())}\n")
        f.write(f"- 측정방법: {dict(meth.most_common())}\n")
        if problems:
            f.write("\n## 검증 문제\n" + "\n".join("- " + p for p in problems) + "\n")


if __name__ == "__main__":
    main()
