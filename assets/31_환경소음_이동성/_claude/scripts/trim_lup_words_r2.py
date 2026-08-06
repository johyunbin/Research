# -*- coding: utf-8 -*-
# LUP 감축 2차(~90단어) — 1차 후 8,066(한도 8,000) 잔여 해소. 수치·인용·주장 불변.
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
 ("Separating day and night", [
   ("(β=+0.50, 95% CI -0.73 to +1.74, p=0.42), which indicates that much of the apparent night-time response reflected time-varying confounding shared across the city (drift, season, secular trends). The daytime association survives the same test.",
    "(β=+0.50, 95% CI -0.73 to +1.74, p=0.42), indicating that much of the apparent night-time response reflected city-wide time-varying confounding (drift, season, secular trends); the daytime association survives the same test."),
   ("Consistent with this, the day-night gap, regressed on the same daytime dose, has a coefficient of essentially zero (β=+0.02, p=0.92), which also fits daytime and night-time noise moving together in similar measure with daytime activity.",
    "Consistent with this, the day-night gap regressed on the same daytime dose has a coefficient of essentially zero (β=+0.02, p=0.92), fitting daytime and night-time noise moving together with daytime activity."),
 ]),
 ("How the exposure moved through time", [
   ("but where that population sits is reallocated sharply with each phase of distancing",
    "but its location is reallocated sharply with each phase of distancing"),
 ]),
 ("Table 4 reports descriptive statistics", [
   ("Adding date fixed effects, which strip out everything the city shared on a given day (weather, trends, drift), shrinks the response:",
    "Adding date fixed effects, which strip out everything the city shared on a given day, shrinks the response:"),
   ("In concrete terms, when a dong's daytime mobility falls",
    "Concretely, when a dong's daytime mobility falls"),
 ]),
 ("The outcome variable is noise from the Smart Seoul", [
   ("We therefore use only within-sensor changes over time, and we check their reliability against the official calibrated LAeq networks",
    "We therefore use only within-sensor changes over time, checked against the official calibrated LAeq networks"),
 ]),
 ("The methodological lessons matter", [
   ("The methodological lessons matter as much as the estimates. Attempts to measure urban environmental effects with dense low-cost noise networks run into three pitfalls:",
    "The methodological lessons matter as much as the estimates. Dense low-cost noise networks run into three pitfalls:"),
   ("These lessons apply to the spreading families of low-cost noise monitoring, SONYC [40], wireless acoustic sensor networks [42], NoiseCapture [44] and low-cost devices generally [41], and to policy uses of smart-city sensor data at large [45,46].",
    "These lessons apply to the spreading families of low-cost noise monitoring (SONYC [40], wireless acoustic sensor networks [42], NoiseCapture [44], low-cost devices generally [41]) and to policy uses of smart-city sensor data [45,46]."),
 ]),
 ("Several auxiliary analyses sit on top", [
   ("an equal-weight re-estimation on the dong-day panel, to confirm that sensor-rich dong do not dominate; a quadratic term, to probe nonlinearity; and Benjamini-Hochberg FDR adjustment",
    "an equal-weight re-estimation on the dong-day panel; a quadratic term for nonlinearity; and Benjamini-Hochberg FDR adjustment"),
   ("Finally, all dong-level spatial analyses use outlier-robust statistics (Theil-Sen regression, Spearman rank correlation) and are restricted to dong with at least two sensors, so that no small set of extreme neighbourhoods drives the results.",
    "All dong-level spatial analyses use outlier-robust statistics (Theil-Sen regression, Spearman rank correlation) and are restricted to dong with at least two sensors, so no small set of extreme neighbourhoods drives the results."),
 ]),
 ("Rather than comparing the two kinds of estimate", [
   ("In perceptual terms a 0.65 dB shift lies below",
    "A 0.65 dB shift lies below"),
 ]),
 ("The pattern of the response is consistent", [
   ("The absence of a clean long-run spatial gradient deserves emphasis. It does not indicate an absent effect; it indicates that",
    "The absence of a clean long-run spatial gradient does not indicate an absent effect; it indicates that"),
   ("Indeed the briefly positive correlation among commercial dong turned out to be an artefact of a few extreme neighbourhoods under Pearson correlation, and it disappears under robust statistics:",
    "The briefly positive correlation among commercial dong was an artefact of a few extreme neighbourhoods under Pearson correlation and disappears under robust statistics:"),
 ]),
 ("Reading the night-time null", [
   ("and at that precision an estimated coefficient would need to reach about 1.2 dB to attain two-sided 5% significance, so a night-time response equal in size to the daytime one would go undetected in these data.",
    "and at that precision an estimated coefficient would need to reach about 1.2 dB to attain two-sided 5% significance; a night-time response equal in size to the daytime one would go undetected."),
   ("A second is exposure variation. Because the practical instruments of distancing were closing hours and gathering caps, the relative variation",
    "A second is exposure variation: with closing hours and gathering caps as the practical instruments of distancing, the relative variation"),
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
    print(f"2차 감축 {n}문단 적용: {os.path.basename(path)}")
