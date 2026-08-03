# -*- coding: utf-8 -*-
"""
Paper32 — 스크리닝 결과 통합·검증
입력: screening/results/chunk_NN_result.csv ×10 + screening_dataset
검증: 1,316 ID 전수·중복 없음·verdict 유효성
출력: screening_results_*.csv(전체) · screening_summary_*.md · borderline_*.csv · include_*.csv
"""
import sys, csv, os, re
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8")
TIMECODE = "20260802_221808"
BASE = r"C:\Users\wh850\Research\assets\32_행태_사운드스케이프 리뷰 논문_(LUP)\_claude"
SCR = os.path.join(BASE, "screening")
VALID = {"INCLUDE", "BORDERLINE", "EX1", "EX2", "EX3", "EX4", "EX5", "EX6", "EX7"}

def main():
    # 데이터셋 로드
    with open(os.path.join(SCR, f"screening_dataset_{TIMECODE}.csv"), encoding="utf-8-sig") as f:
        meta = {r["no"]: r for r in csv.DictReader(f)}
    print(f"데이터셋 {len(meta)}건 로드")

    # 결과 취합
    verdicts, problems = {}, []
    for n in range(1, 11):
        path = os.path.join(SCR, "results", f"chunk_{n:02d}_result.csv")
        if not os.path.exists(path):
            problems.append(f"chunk_{n:02d}_result.csv 없음")
            continue
        with open(path, encoding="utf-8-sig") as f:
            rows = list(csv.DictReader(f))
        for r in rows:
            rid = str(r.get("no", "")).strip()
            v = (r.get("verdict", "") or "").strip().upper()
            if rid in verdicts:
                problems.append(f"중복 ID {rid} (chunk_{n:02d})")
            if v not in VALID:
                problems.append(f"무효 verdict '{v}' ID {rid} (chunk_{n:02d})")
            verdicts[rid] = {
                "verdict": v,
                "direction": (r.get("direction", "") or "").strip().lower(),
                "reason": (r.get("reason", "") or "").strip(),
            }
        print(f"chunk_{n:02d}: {len(rows)}행")

    missing = [k for k in meta if k not in verdicts]
    extra = [k for k in verdicts if k not in meta]
    if missing:
        problems.append(f"누락 ID {len(missing)}건: {missing[:20]}")
    if extra:
        problems.append(f"풀 밖 ID {len(extra)}건: {extra[:20]}")

    print("\n=== 검증 ===")
    print("문제 없음" if not problems else "\n".join("⚠️ " + p for p in problems))

    # 통합 저장
    out = os.path.join(SCR, f"screening_results_{TIMECODE}.csv")
    with open(out, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["no", "verdict", "direction", "reason", "sources", "doi", "year", "journal", "title"])
        for rid in sorted(meta, key=lambda x: int(x)):
            m, v = meta[rid], verdicts.get(rid, {"verdict": "MISSING", "direction": "", "reason": ""})
            w.writerow([rid, v["verdict"], v["direction"], v["reason"],
                        m["sources"], m["doi"], m["year"], m["journal"], m["title"]])

    # 카운트
    counts = Counter(v["verdict"] for v in verdicts.values())
    dirs = Counter(v["direction"] for v in verdicts.values() if v["verdict"] in ("INCLUDE", "BORDERLINE"))
    total_ex = sum(c for k, c in counts.items() if k.startswith("EX"))
    print("\n=== 결과 분포 ===")
    for k in ["INCLUDE", "BORDERLINE", "EX1", "EX2", "EX3", "EX4", "EX5", "EX6", "EX7"]:
        print(f"  {k}: {counts.get(k, 0)}")
    print(f"  배제 합계: {total_ex} · 전문심사 진출(INCLUDE+BORDERLINE): {counts.get('INCLUDE',0)+counts.get('BORDERLINE',0)}")
    print(f"  방향 분포(포함·경계): {dict(dirs)}")

    # 서브셋 파일
    for label, keys in [("include", ("INCLUDE",)), ("borderline", ("BORDERLINE",))]:
        p = os.path.join(SCR, f"{label}_{TIMECODE}.csv")
        with open(p, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["no", "direction", "reason", "year", "journal", "title", "doi"])
            for rid in sorted(meta, key=lambda x: int(x)):
                v = verdicts.get(rid)
                if v and v["verdict"] in keys:
                    m = meta[rid]
                    w.writerow([rid, v["direction"], v["reason"], m["year"], m["journal"], m["title"], m["doi"]])
        print(f"{label} 파일: {p}")

    # 요약 md
    with open(os.path.join(SCR, f"screening_summary_{TIMECODE}.md"), "w", encoding="utf-8") as f:
        f.write(f"# Paper32 제목·초록 스크리닝 AI 사전분류 결과 ({TIMECODE})\n\n")
        f.write(f"- 풀: {len(meta)}건 · 판정: {len(verdicts)}건 · 검증문제: {len(problems)}건\n")
        f.write(f"- INCLUDE {counts.get('INCLUDE',0)} · BORDERLINE {counts.get('BORDERLINE',0)} · EXCLUDE {total_ex}\n")
        for k in ["EX1","EX2","EX3","EX4","EX5","EX6","EX7"]:
            f.write(f"  - {k}: {counts.get(k,0)}\n")
        f.write(f"- 방향(포함·경계): {dict(dirs)}\n")
        if problems:
            f.write("\n## 검증 문제\n" + "\n".join("- " + p for p in problems) + "\n")
        f.write("\n※ 프로토콜상 모든 EX 판정은 인간 리뷰어 검증 대상. BORDERLINE은 전문심사로 진출.\n")
    print("요약 md 저장 완료")


if __name__ == "__main__":
    main()
