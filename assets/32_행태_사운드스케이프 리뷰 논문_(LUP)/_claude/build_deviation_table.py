# -*- coding: utf-8 -*-
"""
Paper32 — 원고 §2.7 이탈 보고표를 `deviation_log.md`에서 기계 생성
손으로 옮기면 건수가 어긋난다(실제로 어긋났다 — 로그 21건 vs 원고 표 11행 vs 본문 "Ten").
출력: fulltext/deviation_table_en.md
"""
import sys, os, re

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")

# 로그 항목 → 원고 표에 실을 영문 요약 + 결과 영향 여부.
# `RESULT` = 보고된 추정치·판정을 실제로 바꾼 것(원고 본문에서 별도 서술).
EN = {
    "D1-1": ("All records still uncertain after title/abstract screening were carried to full text "
             "rather than excluded", "No — widened retrieval, criteria unchanged"),
    "D1-2": ("Three recurring boundary cases formalised as rules after registration, then applied "
             "uniformly", "Moved 3 studies to sensitivity-only"),
    "D1-3": ("Title-level priority filter applied to citation-tracking output (registered "
             "three-block logic)", "Coverage limit; stated in Limitations"),
    "D1-4": ("Citation-tracking branch initially stopped at 'reports sought'; completed by manual "
             "retrieval of 71 reports", "Resolved — branch now complete (15 studies)"),
    "D2-1": ("MMAT categories reassigned from full-text reading rather than from the extracted "
             "design string", "**Yes** — category determines items, so tiers changed"),
    "D2-2": ("RoB 2 not applied separately: only two randomised studies, neither reporting "
             "procedure or blinding", "No"),
    "D2-3": ("One 1978 scanned report left category-undetermined (no text layer)",
             "No — contributes no effect"),
    "D3-1": ("Five registered clusters reduced to four; space-use effects were unpoolable "
             "(no variance information)", "Cluster set; studies remain in narrative synthesis"),
    "D3-2": ("Analysis rules (multiple-effect priority, pseudoreplication, non-convertible inputs) "
             "fixed in a separate document before any computation", "No — pre-specified"),
    "D3-3": ("Three unreported values reconstructed (equal-split n; SD back-calculated from p; "
             "SE→SD)", "Bounded by sensitivity ⑤"),
    "D3-4": ("A sign-coding error of our own found and corrected (I² 93.6% → 43.5%)",
             "**Yes** — corrected before any reporting"),
    "D3-5": ("Two sensitivity axes only partially executable (observation-n available for one "
             "study; two studies report group-level n only)", "Stated per axis"),
    "D3-6": ("Walking-speed pool defined without the music-stimulus study",
             "**Yes** — exclusion is the conservative direction (g −0.474 vs −0.773)"),
    "D4-1": ("Conceptual framework placed in Discussion rather than Results",
             "No — presentation only"),
    "D4-2": ("English-language, journal-only restriction retained as registered",
             "No — registered, not a deviation; consequential for interpretation"),
    "D5-1": ("Exclusion code for language added (Chinese-language journal indexed with English "
             "metadata)", "One verdict changed"),
    "D5-2": ("Intention-outcome rule stated inconsistently across branch instructions",
             "No — no verdict depended on it"),
    "D5-3": ("Full-text extraction cap truncated 13 of 16 citation-tracked studies; all "
             "re-appraised on complete text", "**Yes** — 7 items and 4 tiers changed"),
    "D5-4": ("Four cluster-membership judgements on newly retrieved studies, including one "
             "exclusion that would have strengthened the result",
             "**Yes** — see §2.7(4)"),
    "D5-5": ("Two publication pairs share samples; one member of each pair admitted to any pool",
             "Prevents double-counting"),
    "D5-6": ("Result narrative revised after the social-interaction cluster crossed p = .05",
             "**Yes** — claim changed because the result changed"),
    "D6-1": ("Laboratory-reproduction rule applied to the citation-tracking branch only",
             "No — no such study is poolable"),
    "D6-2": ("Registered supplementary index (OpenAlex) screened late, after self-audit found it "
             "unscreened", "No — the 2 studies found contribute no effect"),
    "D6-3": ("Index language metadata found unreliable (2 of 5 retrieved texts not in English "
             "despite `language=en`)", "Two verdicts changed"),
    "D7-1": ("τ² estimator was maximum likelihood although reported as REML; corrected and "
             "consolidated into a single module", "**Yes** — all four clusters recomputed"),
    "D7-2": ("Walking-speed pool violated our own one-effect-per-study rule; the two experiments "
             "of one study combined as pre-specified", "**Yes** — k 4 → 3, p .179 → .326"),
    "D7-3": ("Two correlational inputs did not match the source: one effect had been constructed "
             "at pooling although extraction recorded none; one used a sample size that was not "
             "the analytic unit", "**Yes** — k 7 → 6, r .409 → .425"),
    "D7-4": ("The rank-correlation sensitivity excluded nothing because the source field was "
             "overwritten before filtering", "**Yes** — real result k = 4, r = .484"),
}

SECTION = {"D1": "Search and selection", "D2": "Quality appraisal", "D3": "Meta-analysis",
           "D4": "Reporting", "D5": "Citation-tracking branch",
           "D6": "Supplementary-index branch",
           "D7": "Post-hoc statistical verification"}


def main():
    log = open(os.path.join(FT, "deviation_log.md"), encoding="utf-8").read()
    ids = re.findall(r"^### (D\d+-\d+)\.", log, re.M)
    order = sorted(set(ids), key=lambda x: (x.split("-")[0], int(x.split("-")[1])))
    missing = [i for i in order if i not in EN]
    extra = [i for i in EN if i not in order]
    if missing or extra:
        print(f"⚠️ 로그에만 있음 {missing} · 표에만 있음 {extra}")

    L = [f"<!-- 기계 생성: build_deviation_table.py · 로그 {len(order)}건 -->\n\n",
         f"| # | Section | Departure | Effect on results |\n|---|---|---|---|\n"]
    n_result = 0
    for i, d in enumerate(order, 1):
        txt, eff = EN[d]
        if eff.startswith("**Yes"):
            n_result += 1
        L.append(f"| {d} | {SECTION[d.split('-')[0]]} | {txt} | {eff} |\n")
    open(os.path.join(FT, "deviation_table_en.md"), "w", encoding="utf-8").write("".join(L))
    print(f"[완료] 이탈 {len(order)}건 · 결과에 영향 준 것 {n_result}건")
    print(f"       섹션별: " + " · ".join(
        f"{SECTION[s]} {sum(1 for d in order if d.startswith(s + '-'))}" for s in SECTION))
    print("[저장] fulltext/deviation_table_en.md")
    return len(order), n_result


if __name__ == "__main__":
    main()
