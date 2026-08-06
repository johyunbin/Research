# -*- coding: utf-8 -*-
"""
Paper32 — MMAT 품질평가 통합 v2 (본검색 84편 + 인용추적 16편)
인용추적분은 1차(절단본)와 재평가(전문)가 겹칠 수 있다 — **전문 기준 재평가를 정본**으로 채택하고
어느 판정이 바뀌었는지 기록한다(추출 파이프라인이 품질평가를 오염시킨 사례).
출력: fulltext/quality_v2.csv · quality_detail_v2.csv · quality_v2_summary.md
     + fulltext/quality_truncation_effect.md (1차 vs 재평가 대조)
"""
import sys, os, csv, glob, re
from collections import Counter, defaultdict

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
QA = os.path.join(FT, "qa_results")

CAT_NORM = {"qualitative": "Qualitative", "quantitative rct": "Quantitative RCT",
            "quantitative non-randomised": "Quantitative non-randomised",
            "quantitative non_randomised": "Quantitative non-randomised",
            "quantitative descriptive": "Quantitative descriptive",
            "mixed methods": "Mixed methods", "mixed_methods": "Mixed methods"}


def norm_cat(s):
    s = re.sub(r"\s+", " ", (s or "").strip().lower())
    return CAT_NORM.get(s, s.title() if s else "Undetermined")


def bare_no(v):
    """QA 배치마다 `no`를 bare number(126)로도, uid 형식(CT0126)으로도 썼다. 둘 다 받는다."""
    s = str(v or "").strip()
    m = re.search(r"(\d+)", s)
    if not m:
        raise ValueError(f"no 파싱 불가: {v!r}")
    return str(int(m.group(1)))


def read_mmat(path, uid_prefix="", dup_log=None):
    out = {}
    for r in csv.DictReader(open(path, encoding="utf-8-sig")):
        n = bare_no(r["no"])
        uid = f"{uid_prefix}{int(n):04d}" if uid_prefix else n
        if uid in out and dup_log is not None:
            dup_log.append(f"{os.path.basename(path)} 내 중복행 {uid}")
        ny = sum(1 for k in ("Q1", "Q2", "Q3", "Q4", "Q5")
                 if (r.get(k) or "").strip().upper() == "Y")
        tier = "high" if ny >= 4 else ("moderate" if ny == 3 else "low")
        out[uid] = {"uid": uid, "orig_no": n, "mmat_category": norm_cat(r.get("mmat_category")),
                    "S1": r.get("S1", ""), "S2": r.get("S2", ""),
                    **{k: (r.get(k) or "").strip().upper() for k in ("Q1", "Q2", "Q3", "Q4", "Q5")},
                    "n_yes": ny, "quality_tier": tier, "note": (r.get("note") or "").strip()}
    return out


def read_detail(path, uid_prefix=""):
    out = []
    for r in csv.DictReader(open(path, encoding="utf-8-sig")):
        n = bare_no(r["no"])
        uid = f"{uid_prefix}{int(n):04d}" if uid_prefix else n
        out.append({"uid": uid, "item": r.get("item", ""), "item_no": r.get("item_no", ""),
                    "verdict": (r.get("verdict") or "").strip().upper(),
                    "rationale": r.get("rationale", "")})
    return out


def main():
    problems = []

    # ── 본검색 84편(기존 정본) ────────────────────────────────────
    main_q = {r["no"]: r for r in csv.DictReader(open(os.path.join(FT, "quality_all.csv"),
                                                      encoding="utf-8-sig"))}
    rows = {}
    for n, r in main_q.items():
        rows[n] = {"uid": n, "source": "db-search", "orig_no": n,
                   "mmat_category": norm_cat(r["mmat_category"]),
                   **{k: r.get(k, "") for k in ("S1", "S2", "Q1", "Q2", "Q3", "Q4", "Q5")},
                   "n_yes": int(r["n_yes"]), "quality_tier": r["quality_tier"],
                   "note": r.get("note", ""), "text_basis": "full"}
    detail = [{"uid": r["no"], "item": r["item"], "item_no": r["item_no"],
               "verdict": r["verdict"], "rationale": r["rationale"]}
              for r in csv.DictReader(open(os.path.join(FT, "quality_detail_all.csv"),
                                           encoding="utf-8-sig"))]

    # ── 인용추적: 1차(절단본) ─────────────────────────────────────
    first, first_detail = {}, []
    for p in sorted(glob.glob(os.path.join(QA, "qa_ct_[ABC]_mmat.csv"))):
        first.update(read_mmat(p, "CT", problems))
    for p in sorted(glob.glob(os.path.join(QA, "qa_ct_[ABC]_detail.csv"))):
        first_detail += read_detail(p, "CT")

    # ── 인용추적: 재평가(전문) — 정본 ─────────────────────────────
    redo, redo_detail = {}, []
    for p in sorted(glob.glob(os.path.join(QA, "qa_ctfull_*_mmat.csv"))):
        redo.update(read_mmat(p, "CT", problems))
    for p in sorted(glob.glob(os.path.join(QA, "qa_ctfull_*_detail.csv"))):
        redo_detail += read_detail(p, "CT")

    # 전문 재추출 대상(절단됐던 논문) 확인
    trunc = set()
    fp = os.path.join(FT, "ct_full_text_index.csv")
    if os.path.exists(fp):
        for r in csv.DictReader(open(fp, encoding="utf-8-sig")):
            if r["was_truncated"] == "yes":
                trunc.add(f"CT{int(r['rec']):04d}")
    missing_redo = sorted(trunc - set(redo))
    if missing_redo:
        problems.append(f"절단 논문인데 재평가 없음: {missing_redo}")

    changes = []
    for uid in sorted(set(first) | set(redo)):
        if uid in redo:
            rec, basis = redo[uid], "full"
            if uid in first:
                a, b = first[uid], redo[uid]
                diff = [k for k in ("mmat_category", "Q1", "Q2", "Q3", "Q4", "Q5")
                        if a.get(k) != b.get(k)]
                if diff or a["quality_tier"] != b["quality_tier"]:
                    changes.append({"uid": uid,
                                    "tier_before": a["quality_tier"], "tier_after": b["quality_tier"],
                                    "nyes_before": a["n_yes"], "nyes_after": b["n_yes"],
                                    "cat_before": a["mmat_category"], "cat_after": b["mmat_category"],
                                    "changed_items": ";".join(diff)})
        else:
            rec, basis = first[uid], ("truncated" if uid in trunc else "full")
        rows[uid] = {**rec, "source": "citation-tracking", "text_basis": basis}

    # detail: 재평가분 우선
    redo_uids = {d["uid"] for d in redo_detail}
    detail += [d for d in first_detail if d["uid"] not in redo_uids] + redo_detail

    # ── 보조검색 갈래(OAS) 품질 ────────────────────────────────────
    oas_p = os.path.join(QA, "qa_oas_mmat.csv")
    if os.path.exists(oas_p):
        for r in csv.DictReader(open(oas_p, encoding="utf-8-sig")):
            uid = str(r["no"]).strip()
            ny = sum(1 for k in ("Q1", "Q2", "Q3", "Q4", "Q5")
                     if (r.get(k) or "").strip().upper() == "Y")
            rows[uid] = {"uid": uid, "source": "openalex-supplementary", "text_basis": "full",
                         "orig_no": uid, "mmat_category": norm_cat(r.get("mmat_category")),
                         "S1": r.get("S1", ""), "S2": r.get("S2", ""),
                         **{k: (r.get(k) or "").strip().upper() for k in ("Q1","Q2","Q3","Q4","Q5")},
                         "n_yes": ny,
                         "quality_tier": "high" if ny >= 4 else ("moderate" if ny == 3 else "low"),
                         "note": (r.get("note") or "").strip()}
        dp = os.path.join(QA, "qa_oas_detail.csv")
        if os.path.exists(dp):
            for r in csv.DictReader(open(dp, encoding="utf-8-sig")):
                detail.append({"uid": str(r["no"]).strip(), "item": r.get("item", ""),
                               "item_no": r.get("item_no", ""),
                               "verdict": (r.get("verdict") or "").strip().upper(),
                               "rationale": r.get("rationale", "")})

    # ── 검증 ──────────────────────────────────────────────────────
    corpus = {r["uid"] for r in csv.DictReader(open(os.path.join(FT, "corpus_v4_verdicts.csv"),
                                                    encoding="utf-8-sig"))
              if r["final_verdict"] in ("FINAL_INCLUDE", "SENS_ONLY")}
    miss = sorted(corpus - set(rows))
    extra = sorted(set(rows) - corpus)
    if miss:
        problems.append(f"품질평가 누락 {len(miss)}건: {miss[:10]}")
    if extra:
        problems.append(f"대상 밖 {len(extra)}건: {extra[:10]}")
    for uid, r in rows.items():
        ny = sum(1 for k in ("Q1", "Q2", "Q3", "Q4", "Q5") if r.get(k) == "Y")
        if ny != r["n_yes"]:
            problems.append(f"{uid}: n_yes 불일치 {r['n_yes']} vs 재계산 {ny}")
    dcount = Counter(d["uid"] for d in detail)
    bad_d = [u for u, c in dcount.items() if c != 5]
    if bad_d:
        problems.append(f"문항 5행이 아닌 논문: {bad_d[:8]}")

    # ── 저장 ──────────────────────────────────────────────────────
    cols = ["uid", "source", "text_basis", "orig_no", "mmat_category", "S1", "S2",
            "Q1", "Q2", "Q3", "Q4", "Q5", "n_yes", "quality_tier", "note"]
    with open(os.path.join(FT, "quality_v2.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore"); w.writeheader()
        for uid in sorted(rows, key=lambda x: (x.startswith("CT"), x)):
            w.writerow(rows[uid])
    with open(os.path.join(FT, "quality_detail_v2.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["uid", "item", "item_no", "verdict", "rationale"])
        w.writeheader(); w.writerows(detail)

    tier = Counter(r["quality_tier"] for r in rows.values())
    by_src = defaultdict(Counter)
    for r in rows.values():
        by_src[r["source"]][r["quality_tier"]] += 1
    cat = Counter(r["mmat_category"] for r in rows.values())

    print("=== 검증 ===")
    print("문제 없음" if not problems else "\n".join("⚠️ " + x for x in problems))
    print(f"\n품질평가 {len(rows)}편 · 등급 {dict(tier)}")
    for s, c in by_src.items():
        print(f"  {s}: {dict(c)}")
    print(f"범주 {dict(cat)}")
    print(f"절단본→전문 재평가로 판정이 바뀐 논문 {len(changes)}편")

    # ── 절단 효과 기록 ────────────────────────────────────────────
    L = ["# Paper32 — 텍스트 절단이 품질평가에 미친 영향 (자체 감사)\n",
         "\n인용추적 갈래 1차 품질평가는 논문당 46,000자로 잘린 텍스트로 수행됐다. "
         "MMAT는 *'보고가 없으면 CT(can't tell)'* 로 판정하는 도구이므로, **추출 한도가 그대로 "
         "품질 점수를 깎는 인공물**이 된다. 포함 16편 중 **13편이 한도를 초과**했다(최장 96,472자).\n",
         "\n절단 없는 전문으로 재평가한 뒤 두 결과를 대조했다. **전문 기준 재평가가 정본**이다.\n"]
    if changes:
        L.append(f"\n## 판정이 바뀐 논문 — {len(changes)}편\n\n")
        L.append("| 논문 | 등급 | n_yes | 범주 | 바뀐 문항 |\n|---|---|---|---|---|\n")
        for c in changes:
            tl = (f"{c['tier_before']} → **{c['tier_after']}**"
                  if c["tier_before"] != c["tier_after"] else c["tier_after"])
            cl = (f"{c['cat_before']} → **{c['cat_after']}**"
                  if c["cat_before"] != c["cat_after"] else c["cat_after"])
            L.append(f"| {c['uid']} | {tl} | {c['nyes_before']} → {c['nyes_after']} | {cl} | "
                     f"{c['changed_items'] or '—'} |\n")
    else:
        L.append("\n## 판정 변화 없음\n\n재평가 결과가 1차와 동일했다.\n")
    L.append("\n## 함의\n\n"
             "이 감사는 **자동 전문 추출을 쓰는 리뷰가 품질평가에서 체계적 하향 편의를 가질 수 있음**을 "
             "보여준다. 추출 한도·OCR 실패·부록 누락은 모두 'CT'로 흘러들어가고, CT는 `n_yes`를 낮춰 "
             "등급을 떨어뜨린다. 본 리뷰는 절단이 확인된 전건을 전문으로 재평가해 해소했고, "
             "그 과정을 프로토콜 이탈 로그에 기록한다.\n")
    open(os.path.join(FT, "quality_truncation_effect.md"), "w", encoding="utf-8").write("".join(L))
    print("[저장] quality_v2.csv · quality_detail_v2.csv · quality_truncation_effect.md")


if __name__ == "__main__":
    main()
