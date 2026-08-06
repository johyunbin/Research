# -*- coding: utf-8 -*-
# LUP 감축 3차(~47단어) — 8,015 → ≤7,970 목표. 수치·인용·주장 불변.
import os, sys
from docx import Document
sys.path.insert(0, os.path.dirname(__file__))
import docxtools_v9 as T

def para_size(p, default=11.0):
    for r in p.runs:
        if r.font.size is not None:
            return r.font.size.pt
    st = p.style
    while st is not None:
        if st.font.size is not None:
            return st.font.size.pt
        st = st.base_style
    return default

EDITS = [
 ("Our estimate stands in contrast", [
   ("is estimated across roughly 1,100 locations, from within-sensor changes, along the whole range of mobility, over the full period;",
    "is estimated across roughly 1,100 locations, from within-sensor changes, over the whole range of mobility and the full period;"),
 ]),
 ("Rather than comparing the two kinds of estimate", [
   ("The drop in the daytime coefficient from +1.13 to +0.65 when date fixed effects are added reflects",
    "The drop in the daytime coefficient from +1.13 to +0.65 under date fixed effects reflects"),
 ]),
 ("This study combines Korea's graded social distancing", [
   ("Specifically, we adopt a two-way fixed-effects strategy [48,49] within a natural-experiment framework [50] that identifies the slope only from within-sensor changes and from variation across dong within the same date, never from absolute levels, and we check the drift and resolution limits of the low-cost network against official monitoring stations.",
    "Specifically, we adopt a two-way fixed-effects strategy [48,49] within a natural-experiment framework [50], identifying the slope only from within-sensor changes and from variation across dong within the same date, never from absolute levels, and check the drift and resolution limits of the low-cost network against official stations."),
   ("We ask three questions, and answering them doubles",
    "We ask three questions; answering them doubles"),
 ]),
 ("Mobility exposure is measured with Seoul's", [
   ("estimated from mobile-network signals rather than from residence records (Table 2)",
    "estimated from mobile-network signals rather than residence records (Table 2)"),
 ]),
 ("By contrast, the long-run spatial cross-section", [
   ("Across all analysable dong (those with at least two sensors), the rank correlation",
    "Across all analysable dong (at least two sensors), the rank correlation"),
   ("but that value is driven by a few extreme neighbourhoods; robust statistics on the same dong leave almost nothing (Spearman ρ=+0.08, Theil-Sen slope near zero; Fig. 7c, detail in Supplementary Fig. S1)",
    "but that value is driven by a few extreme neighbourhoods; robust statistics leave almost nothing (Spearman ρ=+0.08, Theil-Sen slope near zero; Fig. 7c; Supplementary Fig. S1)"),
 ]),
 ("Five sensitivity checks establish", [
   ("First, reshuffling the dong-level dose across dong within each date (300 shuffles, applied identically to all sensors of a dong) yields placebo coefficients",
    "First, reshuffling the dong-level dose across dong within each date (300 shuffles) yields placebo coefficients"),
 ]),
 ("Several limitations qualify these conclusions", [
   ("(8) The analysis window opens with S-DoT provision (2020-04) and misses the strongest mobility shock, the first wave and the first intensive distancing of March 2020; with the extremes",
    "(8) The analysis window opens with S-DoT provision (2020-04) and misses the strongest mobility shock of March 2020; with the extremes"),
 ]),
 ("The outcome variable is noise from the Smart Seoul", [
   ("computing the full-day Leq,24h, the daytime Lday (06-21 h) and the night-time Lnight (22-05 h) as 10·log10 of the mean of 10^(L/10), and retain only sensor-days with at least 12 observed hours and values in a physically plausible range (20-95 dB)",
    "computing the full-day Leq,24h, daytime Lday (06-21 h) and night-time Lnight (22-05 h) as 10·log10 of the mean of 10^(L/10), and retain sensor-days with at least 12 observed hours and values in a plausible range (20-95 dB)"),
   ("Day and night windows follow the calendar date; Lnight combines the 00-05 h and 22-23 h records of the same calendar day.",
    "Day and night windows follow the calendar date; Lnight combines the same day's 00-05 h and 22-23 h records."),
 ]),
 ("This rapidly accumulated literature", [
   ("no representative study satisfies more than two of them at once (Table 1), and none brings all four together in a dense East Asian city. This study addresses that gap.",
    "no representative study satisfies more than two at once, and none brings all four together in a dense East Asian city (Table 1). This study addresses that gap."),
 ]),
]

MANUSCRIPTS = [
    r"C:\Users\wh850\Research\assets\31_환경소음_이동성\01_논문작업\Manuscript_EN_20260727_005544.docx",
    r"C:\Users\wh850\Research\assets\31_환경소음_이동성\260806_LUP_투고\03_Manuscript.docx",
]
for path in MANUSCRIPTS:
    doc = Document(path)
    n = 0
    for marker, subs in EDITS:
        hits = [p for p in doc.paragraphs if p.text.strip().startswith(marker)]
        assert len(hits) == 1, f"{os.path.basename(path)}: 마커 {len(hits)}건 — {marker}"
        p = hits[0]
        text = p.text
        for old, new in subs:
            assert old in text, f"{os.path.basename(path)}: 원문 미발견 — {old[:70]}"
            text = text.replace(old, new)
        T.fill_para(p, text, size=para_size(p))
        n += 1
    assert n == len(EDITS)
    doc.save(path)
    print(f"3차 감축 {n}문단 적용: {os.path.basename(path)}")
