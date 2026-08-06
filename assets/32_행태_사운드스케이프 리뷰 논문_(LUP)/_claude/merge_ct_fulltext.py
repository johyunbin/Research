# -*- coding: utf-8 -*-
"""
Paper32 — 인용추적 전문심사 결과 통합·검증
9배치의 판정·추출을 합치고, 심사 대상 54편과 1:1 대조한다.
출력: fulltext/ct_verdicts_final.csv · ct_extraction_final.csv · ct_fulltext_summary.md
"""
import sys, os, csv, glob
from collections import Counter, defaultdict

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
RES = os.path.join(FT, "ct_results_ft")

VERD = {"FINAL_INCLUDE", "SENS_ONLY", "FINAL_EXCLUDE", "UNCERTAIN"}
EXT_COLS = ["no", "final_verdict", "year", "journal", "title", "country", "setting", "design",
            "sample_n", "exposure", "behaviour_domain", "behaviour_measure",
            "measurement_method", "direction", "key_finding", "effect_stats"]


def main():
    idx = {r["rec"]: r for r in csv.DictReader(open(os.path.join(FT, "ct_packet_index.csv"),
                                                    encoding="utf-8-sig"))}
    problems = []

    # ── 판정 통합 ──────────────────────────────────────────────────
    verd, seen_batch = {}, {}
    for p in sorted(glob.glob(os.path.join(RES, "ctft_*_verdict.csv"))):
        b = os.path.basename(p).split("_verdict")[0]
        n = 0
        for r in csv.DictReader(open(p, encoding="utf-8-sig")):
            rec = str(int(str(r["rec"]).strip()))
            if rec in verd:
                problems.append(f"중복 판정 REC {rec} ({seen_batch[rec]} vs {b})")
            v = (r["verdict"] or "").strip().upper()
            if v not in VERD:
                problems.append(f"REC {rec}: 판정값 이상 '{v}'")
            verd[rec] = {"rec": rec, "verdict": v,
                         "reason_code": (r.get("reason_code") or "").strip(),
                         "rationale": (r.get("rationale") or "").strip(),
                         "boundary_rule": (r.get("boundary_rule") or "").strip(),
                         "batch": b, "screen1": idx.get(rec, {}).get("screen", ""),
                         "year": idx.get(rec, {}).get("year", ""),
                         "journal": idx.get(rec, {}).get("journal", ""),
                         "doi": idx.get(rec, {}).get("doi", ""),
                         "title": idx.get(rec, {}).get("title", "")}
            seen_batch[rec] = b
            n += 1
        print(f"  {b}: 판정 {n}건")

    missing = sorted(set(idx) - set(verd), key=int)
    extra = sorted(set(verd) - set(idx), key=int)
    if missing:
        problems.append(f"판정 누락 {len(missing)}건: {missing}")
    if extra:
        problems.append(f"대상 밖 판정: {extra}")

    # ── 추출 통합 ──────────────────────────────────────────────────
    ext, ext_seen = [], set()
    for p in sorted(glob.glob(os.path.join(RES, "ctft_*_extract.csv"))):
        for r in csv.DictReader(open(p, encoding="utf-8-sig")):
            if not (r.get("no") or "").strip():
                continue
            rec = str(int(str(r["no"]).strip()))
            if rec in ext_seen:
                problems.append(f"중복 추출 REC {rec}")
            ext_seen.add(rec)
            ext.append({c: (r.get(c) or "").strip() for c in EXT_COLS} | {"no": rec})

    keep = {k for k, v in verd.items() if v["verdict"] in ("FINAL_INCLUDE", "SENS_ONLY")}
    if ext_seen != keep:
        if keep - ext_seen:
            problems.append(f"추출 누락(포함인데 행 없음): {sorted(keep - ext_seen, key=int)}")
        if ext_seen - keep:
            problems.append(f"추출 과잉(배제인데 행 있음): {sorted(ext_seen - keep, key=int)}")

    # 추출 행의 verdict가 판정과 일치하는가
    for r in ext:
        v = verd.get(r["no"], {}).get("verdict", "")
        if r["final_verdict"] != v:
            problems.append(f"REC {r['no']}: 추출 verdict '{r['final_verdict']}' ≠ 판정 '{v}'")

    # ── 저장 ───────────────────────────────────────────────────────
    with open(os.path.join(FT, "ct_verdicts_final.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["rec", "verdict", "reason_code", "rationale",
                                          "boundary_rule", "batch", "screen1", "year",
                                          "journal", "doi", "title"])
        w.writeheader()
        for k in sorted(verd, key=int):
            w.writerow(verd[k])
    with open(os.path.join(FT, "ct_extraction_final.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=EXT_COLS); w.writeheader()
        w.writerows(sorted(ext, key=lambda r: int(r["no"])))

    # ── 요약 ───────────────────────────────────────────────────────
    vc = Counter(v["verdict"] for v in verd.values())
    by_screen = defaultdict(Counter)
    for v in verd.values():
        by_screen[v["screen1"]][v["verdict"]] += 1
    rc = Counter(v["reason_code"] for v in verd.values() if v["reason_code"])
    br = Counter(v["boundary_rule"] for v in verd.values() if v["boundary_rule"])
    with_es = sum(1 for r in ext if r["effect_stats"] and r["effect_stats"].upper() != "NR")
    dom = Counter()
    for r in ext:
        for d in (r["behaviour_domain"] or "").split(";"):
            if d.strip():
                dom[d.strip()] += 1

    print("\n=== 검증 ===")
    print("문제 없음" if not problems else "\n".join("⚠️ " + x for x in problems))
    print(f"\n심사 {len(verd)}편 · 판정 {dict(vc)}")
    print(f"1차판정별: " + " / ".join(f"{k} → {dict(c)}" for k, c in by_screen.items()))
    print(f"배제 사유 {dict(rc)} · 경계원칙 {dict(br)}")
    print(f"추출 {len(ext)}행 · 효과크기 보유 {with_es}행 · 도메인 {dict(dom)}")

    L = ["# Paper32 — 인용추적 갈래 전문심사 결과\n",
         f"\n등록 프로토콜(osf.io/7ew8q)의 citation tracking 이행분. **본검색과 동일한 적격 5기준·"
         f"경계 3원칙**을 적용해 {len(verd)}편을 전문 심사했다.\n",
         f"\n## 판정\n\n| 판정 | 건수 |\n|---|---|\n"]
    for k in ("FINAL_INCLUDE", "SENS_ONLY", "FINAL_EXCLUDE", "UNCERTAIN"):
        if vc.get(k):
            L.append(f"| {k} | {vc[k]} |\n")
    L.append(f"\n## 1차 스크리닝 판정별 적중률\n\n"
             "초록에서 행태 아웃컴을 확인한 A군과 제목만으로 통과시킨 B군의 차이를 보여준다 — "
             "**초록 확보 여부가 곧 적중률**이라는 방법론적 관찰.\n\n"
             "| 1차판정 | 심사 | 포함(+민감도) | 적중률 |\n|---|---|---|---|\n")
    for k in ("RETRIEVE", "UNCERTAIN"):
        c = by_screen.get(k)
        if not c:
            continue
        tot = sum(c.values())
        hit = c.get("FINAL_INCLUDE", 0) + c.get("SENS_ONLY", 0)
        lab = "A군(초록 확인)" if k == "RETRIEVE" else "B군(제목만)"
        L.append(f"| {lab} | {tot} | {hit} | {hit/tot*100:.0f}% |\n")
    L.append(f"\n## 배제 사유\n\n")
    NAMES = {"X1": "동물·생태음향", "X2": "세팅 부적합", "X3": "음환경 변수 없음(또는 통제변수만)",
             "X4": "관찰가능 행태 아웃컴 없음", "X5": "비실증", "X6": "중복·철회"}
    for k, v in rc.most_common():
        L.append(f"- **{k}** {NAMES.get(k, '')} — {v}건\n")
    if br:
        L.append(f"\n경계 판정 원칙 적용: {dict(br)}\n")
    L.append(f"\n## 추출\n\n- 추출표 {len(ext)}행 · **효과크기(verbatim) 보유 {with_es}행**\n"
             f"- 행태 도메인: {dict(dom)}\n")
    if problems:
        L.append("\n## ⚠️ 검증 문제\n\n" + "\n".join(f"- {x}" for x in problems) + "\n")
    else:
        L.append("\n검증: 판정 누락·중복 0 · 추출/판정 정합 일치 · verdict 어휘 유효.\n")
    open(os.path.join(FT, "ct_fulltext_summary.md"), "w", encoding="utf-8").write("".join(L))
    print("\n[저장] ct_verdicts_final.csv · ct_extraction_final.csv · ct_fulltext_summary.md")


if __name__ == "__main__":
    main()
