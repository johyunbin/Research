# -*- coding: utf-8 -*-
# LUP 분량 감축(사용자 승인: 최소 감축 ~500단어) — 본문(캡션 제외)+참고문헌 8,414 → ≤8,000.
# 원칙: 수치·인용번호·핵심 주장 불변 · 연결어/중복/메타 문장만 압축 · §2.4 상세 → Supplementary Methods S1.
# 대상: repo EN 정본 + 260806_LUP_투고/03_Manuscript (동일 본문) + Supplementary 2본에 S1 삽입.
# 최근 교신저자 수정 구역(§1.3·§3.2·§4.3·§4.5·§5)은 불변.
import os, sys, copy
from docx import Document
from docx.text.paragraph import Paragraph
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

# (문단 시작 마커, [(old, new), ...]) — old는 현재 본문 exact substring
EDITS = [
 ("How much quieter does a city become", [
   ("The World Health Organization identifies it as a leading cause of disease burden in European cities and estimates that",
    "The World Health Organization estimates that"),
   ("The health effects go beyond annoyance. Night-time noise fragments sleep [3], chronic exposure raises the risk of hypertension and ischaemic heart disease through autonomic and endocrine stress pathways [4,5], and daytime exposure causes annoyance and impairs cognitive performance [6,7]. These effects translate into",
    "Night-time noise fragments sleep [3], chronic exposure raises the risk of hypertension and ischaemic heart disease through autonomic and endocrine stress pathways [4,5], daytime exposure causes annoyance and impairs cognitive performance [6,7], and these effects translate into"),
 ]),
 ("Noise is therefore a central physical stressor", [
   ("Managing noise, however, requires an answer to a basic question. The lever that sustainability policy actually holds is not the vehicle fleet itself but the volume and placement of travel and activity. How sensitively, then, does urban noise respond to that lever, to human activity and mobility? This responsiveness is the starting point",
    "The lever that policy actually holds, however, is not the vehicle fleet itself but the volume and placement of travel and activity. How sensitively urban noise responds to that lever is the starting point"),
 ]),
 ("There are clear physical grounds", [
   ("There are clear physical grounds for expecting such responsiveness. Road traffic is the dominant source of urban noise, and sound levels rise with traffic volume [17]. But policy acts on more than traffic counts; it acts on the whole of human activity and movement, and",
    "The physical grounds for expecting such responsiveness are clear: road traffic is the dominant source of urban noise, and sound levels rise with traffic volume [17]. But policy acts on the whole of human activity and movement, not on traffic counts alone, and"),
   ("The obstacle is identification. In a normally functioning city, mobility does not vary exogenously. It is bound up with land use, time of day, day of week and weather, the same factors that drive noise. Simple correlations between observed mobility and noise are therefore confounded, and they say little",
    "The obstacle is identification. In a normally functioning city mobility is bound up with land use, time of day, day of week and weather, the same factors that drive noise, so simple correlations between observed mobility and noise are confounded and say little"),
 ]),
 ("Two recent developments in urban data", [
   ("Two recent developments in urban data make the first three requirements attainable. The first is de-facto population estimated from mobile-network signals, which records where people actually are at a given moment, at fine spatial and temporal resolution, rather than where they reside [35,36,37,38,39]. The second is the low-cost acoustic sensor network. SONYC [40], low-cost monitoring devices [41], wireless acoustic sensor networks [42,43], the participatory NoiseCapture platform [44] and smart-city sensing programmes [45,46] now sustain continuous observation of urban sound at hundreds to thousands of points. Korea's graded distancing, the citywide S-DoT noise network in Seoul (about 1,100 sensors) and dong-level de-facto population provide these ingredients together in a single city, a rare combination.",
    "Two recent developments in urban data make the first three requirements attainable. De-facto population estimated from mobile-network signals records where people actually are, at fine spatial and temporal resolution, rather than where they reside [35,36,37,38,39]. Low-cost acoustic sensing, from SONYC [40] and low-cost devices [41] to wireless acoustic sensor networks [42,43], the participatory NoiseCapture platform [44] and smart-city programmes [45,46], now sustains continuous observation of urban sound at hundreds to thousands of points. Korea's graded distancing, Seoul's citywide S-DoT network (about 1,100 sensors) and dong-level de-facto population bring these ingredients together in a single city, a rare combination."),
 ]),
 ("Low-cost noise networks, however,", [
   ("each sensor has its own calibration offset, and readings drift slowly with age and environmental exposure. The same problem is well documented for low-cost air-quality networks [47]. These devices are simply not calibrated to measure absolute sound pressure levels. Ignoring this and comparing absolute levels over time or across space produces spurious patterns or spurious conclusions, and this study demonstrates a concrete case. What remains is the fourth requirement: removing this contamination at the design stage of identification rather than at the stage of interpretation.",
    "each sensor has its own calibration offset, readings drift with age and exposure, and the same problem is well documented for low-cost air-quality networks [47]. These devices are not calibrated to measure absolute sound pressure levels, and comparing absolute levels over time or space produces spurious patterns or spurious conclusions; this study demonstrates a concrete case. The fourth requirement is therefore to remove this contamination at the design stage of identification rather than at interpretation."),
 ]),
 ("This study combines Korea's graded social distancing", [
   ("; the position of this study relative to prior work is summarised in the last row of Table 1",
    " (Table 1, last row)"),
 ]),
 ("The study area is the whole of Seoul", [
   ("The study area is the whole of Seoul, covering all 25 autonomous districts. With about 9.6 million residents on 605 km², Seoul is among the densest megacities in the world, and its fine grid of 421 administrative neighbourhoods (dong) over intensely mixed land use makes it well suited to capturing within-city contrasts (Fig. 1).",
    "The study area is the whole of Seoul: about 9.6 million residents on 605 km² across 25 autonomous districts, among the densest megacities in the world, whose fine grid of 421 administrative neighbourhoods (dong) over intensely mixed land use is well suited to capturing within-city contrasts (Fig. 1)."),
 ]),
 ("The outcome variable is noise from the Smart Seoul", [
   ("One caveat matters throughout. S-DoT is a general-purpose urban sensing node rather than an environmental-noise monitor: it reports uncalibrated broadband decibels (the values carry no A-weighted LAeq label, and the microphone specification, the frequency and time weighting, and the rule aggregating 2-minute raw records into hourly values are not documented in the public specification), so absolute readings cannot be compared across sensors.",
    "One caveat matters throughout: S-DoT is a general-purpose urban sensing node rather than an environmental-noise monitor. It reports uncalibrated broadband decibels, with no A-weighted LAeq label and no public documentation of the microphone specification, the frequency and time weighting, or the rule aggregating 2-minute raw records into hourly values, so absolute readings cannot be compared across sensors."),
 ]),
 ("Mobility exposure is measured with Seoul's", [
   ("One feature of these data shapes the dose definition. The citywide total",
    "The citywide total"),
   ("so that it captures how far a dong departs from its own normal and is unaffected by",
    "which captures how far a dong departs from its own normal, unaffected by"),
   ("Two clarifications are in order. First, de-facto population is a stock of presence, not a flow of movement, so 'mobility' in this paper is an operational term for changes in activity presence. Second, the baseline period is 2022-07 to 2023-12, chosen to exclude the transitional weeks immediately after full lifting (2022-04 to 06).",
    "De-facto population is a stock of presence rather than a flow of movement, so 'mobility' in this paper is an operational term for changes in activity presence. The baseline period, 2022-07 to 2023-12, excludes the transitional weeks immediately after full lifting (2022-04 to 06)."),
 ]),
 ("Sensor metadata contain street addresses", [
   ("Sensor metadata contain street addresses but no administrative-neighbourhood name, so each sensor was assigned to a dong by point-in-polygon matching of its coordinates against public boundary data (the vuski/admdongkor GeoJSON, 2022-01 vintage). The de-facto population data and the boundary file use slightly different dong-code systems in a few reorganised areas of Gangbuk and Gangdong; these were reconciled by name matching. In the result, 1,165 of 1,170 sensors (99.6%) were assigned to a dong, and the assigned dong agreed with the district recorded in the street address in 99.6% of cases. After the noise quality filters (at least 12 h per day, 20-95 dB) and the merge with daily mobility,",
    "Each sensor was assigned to a dong by point-in-polygon matching of its coordinates against public administrative boundaries; 1,165 of 1,170 sensors (99.6%) were assigned, and the assigned dong agreed with the district in the sensor's street address in 99.6% of cases (Supplementary Methods S1). After the quality filters and the merge with daily mobility,"),
 ]),
 ("Distancing 'tiers'", [
   ("Distancing 'tiers' (Levels 1, 2, 2.5 and 4, among others) are unsuitable as an exposure variable: the tier system was redefined twice (2020-11 and 2021-07), the spacing between tiers is not uniform, and the same tier corresponded to different mobility at different times. We instead re-coded the daily regime into a continuous stringency index from 0 (no restriction) to 7 (strongest), built from two concrete enforceable components, the business closing hour and the private-gathering cap (Table 3). This index remains auxiliary; the primary exposure throughout is measured mobility (de-facto population). Weather comes from the Open-Meteo ERA5 daily reanalysis (temperature, precipitation, wind). The calendar controls entering M1 are weekend and Korean public-holiday indicators; day-of-week and season are available in the data, but under M2 the date fixed effects absorb all calendar effects, so separate terms are unnecessary.",
    "Distancing 'tiers' are unsuitable as an exposure variable: the tier system was redefined twice (2020-11 and 2021-07) and the same tier corresponded to different mobility at different times. We instead re-coded the daily regime into a continuous stringency index from 0 (no restriction) to 7 (strongest), built from the two enforceable components, the business closing hour and the private-gathering cap (Table 3); the index remains auxiliary, and the primary exposure throughout is measured mobility. Weather comes from the Open-Meteo ERA5 daily reanalysis (temperature, precipitation, wind). M1 enters weekend and Korean public-holiday indicators; under M2 the date fixed effects absorb all calendar effects."),
 ]),
 ("A low-cost S-DoT sensor is not precision-calibrated", [
   ("A low-cost S-DoT sensor is not precision-calibrated: presented with the same sound, each unit records a level that is systematically high or low by its own constant amount. Because this offset differs across sensors and is unknown, comparing absolute decibels between two sensors is meaningless. If we look only at how a sensor changes over time, however, for instance how today compares with that sensor's own average, the constant offset cancels in the subtraction. Every analysis in this paper therefore rests on within-sensor relative change rather than on absolute levels. Statistically this takes the form of a fixed-effects model: subtracting each sensor's mean removes its idiosyncratic offset (we use the mathematically equivalent within-transformation rather than 1.25 million dummy variables).",
    "A low-cost S-DoT sensor is not precision-calibrated: presented with the same sound, each unit reads high or low by its own constant, unknown amount, so absolute decibels cannot be compared across sensors. The offset cancels, however, in any within-sensor change over time. Every analysis in this paper therefore rests on within-sensor relative change, implemented as a fixed-effects model in which subtracting each sensor's mean removes its idiosyncratic offset (via the mathematically equivalent within-transformation rather than 1.25 million dummy variables)."),
 ]),
 ("Several auxiliary analyses sit on top", [
   ("(commercial, mixed and residential terciles of each dong's daytime-to-night-time population ratio; Supplementary Fig. S3; the classification uses only post-lifting level information rather than the daily exposure variation, which limits circularity, and although it is in substance an activity-profile grouping we call it 'land use' for brevity)",
    "(commercial, mixed and residential terciles of the dong's daytime-to-night-time population ratio, classified from post-lifting levels only to limit circularity; in substance an activity-profile grouping that we call 'land use' for brevity; Supplementary Fig. S3)"),
   ("All computations use Python 3.13 (pandas, NumPy, SciPy, statsmodels); the two-way fixed effects are absorbed by iterated demeaning, and the robust estimators come from SciPy.",
    "All computations use Python 3.13 (pandas, NumPy, SciPy, statsmodels), with the two-way fixed effects absorbed by iterated demeaning."),
 ]),
 ("Table 4 reports descriptive statistics", [
   ("(β=+1.130 dB per log-unit, SE 0.292, p<0.001; standard errors are clustered at the dong level, where the dose is assigned)",
    "(β=+1.130 dB per log-unit, SE 0.292, p<0.001)"),
 ]),
 ("The same functional contrast is visible over time", [
   ("(a descriptive comparison, Fig. 6; the groups are defined ex post from realised mobility and both series are normalised to the post-lifting window, so this is not a formal event study)",
    "(a descriptive comparison rather than a formal event study, for the reasons noted in Section 2.6; Fig. 6)"),
 ]),
 ("By contrast, the long-run spatial cross-section", [
   ("The mobility-noise signal, in other words, lives in same-day differences",
    "The mobility-noise signal lives in same-day differences"),
 ]),
 ("We checked this decline against the calibrated official", [
   ("The calibrated rise deserves note for another reason: calibrated noise was lowest in 2020, when mobility was most depressed, and climbed through 2023 as activity recovered, matching the direction of our dose-response. The networks that can be trusted in multi-year absolute terms thus support the sign of our estimate; the series that moves the other way is the drift-contaminated S-DoT trend.",
    "The calibrated rise deserves note for another reason: calibrated noise was lowest in 2020, when mobility was most depressed, and climbed through 2023 as activity recovered, so the networks that can be trusted in multi-year absolute terms independently support the direction of our estimate."),
   ("with date fixed effects absorbing all multi-year and seasonal variation, and makes no use of the S-DoT multi-year trend. This is precisely why the analysis relies on within-sensor change and same-day comparison (M2) instead of absolute or time-series contrasts.",
    "with date fixed effects absorbing all multi-year and seasonal variation; it makes no use of the S-DoT multi-year trend."),
 ]),
 ("This study provides, to our knowledge, the first quantitative", [
   ("Fourth, the multi-year absolute levels of the low-cost network are contaminated by sensor drift and uneven recovery, so neither time-series nor absolute comparisons can be trusted; only designs built on within-sensor change and cross-neighbourhood comparison remain valid. The upward multi-year trend of the calibrated networks, which tracks recovering activity, independently supports the direction of the estimated response.",
    "Fourth, sensor drift and uneven recovery contaminate the network's multi-year absolute levels, leaving only designs built on within-sensor change and cross-neighbourhood comparison valid. The upward multi-year trend of the calibrated networks, tracking recovering activity, independently supports the direction of the response."),
 ]),
 ("Our estimate stands in contrast to the several-decibel", [
   ("Those studies mostly compare absolute noise at particular busy locations before and during confinement, an approach that readily generalises",
    "Most compare absolute noise at busy locations before and during confinement, which readily generalises"),
   ("One qualification is essential. The core model identifies",
    "One qualification is essential: the core model identifies"),
 ]),
 ("The two kinds of estimate answer different questions", [
   ("The two kinds of estimate answer different questions, so rather than comparing them head-on it is more useful to push our slope",
    "Rather than comparing the two kinds of estimate head-on, it is more useful to push our slope"),
 ]),
 ("Several limitations qualify these conclusions", [
   ("(2) De-facto population records presence, not the activities that generate noise (traffic above all), and dong-by-date shocks such as construction or events, which move both presence and noise, survive date fixed effects; this is why we read the estimate as a conditional association. Linking stop-level transit boardings to dong [35,37] would sharpen interpretation.",
    "(2) De-facto population records presence, not the activities that generate noise (traffic above all), and dong-by-date shocks such as construction or events, which move both presence and noise, survive date fixed effects, which is why we read the estimate as a conditional association; linking stop-level transit boardings to dong [35,37] would sharpen interpretation."),
   ("(4) Weather comes from reanalysis (ERA5) and could be replaced with official station observations.",
    "(4) Weather comes from reanalysis (ERA5) rather than official station observations."),
 ]),
]

S1_HEAD = "Supplementary Methods S1. Sensor-to-neighbourhood matching"
S1_BODY = ("Sensor metadata contain street addresses but no administrative-neighbourhood name. Each sensor was "
           "therefore assigned to a dong by point-in-polygon matching of its coordinates against public "
           "boundary data (the vuski/admdongkor GeoJSON, 2022-01 vintage). The de-facto population data and "
           "the boundary file use slightly different dong-code systems in a few reorganised areas of Gangbuk "
           "and Gangdong; these were reconciled by name matching. In the result, 1,165 of 1,170 sensors "
           "(99.6%) were assigned to a dong, and the assigned dong agreed with the district recorded in the "
           "street address in 99.6% of cases. Sensor-days then passed the noise quality filters (at least 12 "
           "observed hours per day, values within 20-95 dB) before the merge with daily mobility.")

MANUSCRIPTS = [
    r"C:\Users\wh850\Research\assets\31_환경소음_이동성\01_논문작업\Manuscript_EN_20260727_005544.docx",
    r"C:\Users\wh850\Research\assets\31_환경소음_이동성\260806_LUP_투고\03_Manuscript.docx",
]
SUPPS = [
    r"C:\Users\wh850\Research\assets\31_환경소음_이동성\01_논문작업\Supplementary_20260726_231342.docx",
    r"C:\Users\wh850\Research\assets\31_환경소음_이동성\260806_LUP_투고\04_Supplementary.docx",
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
    print(f"감축 {n}문단 적용: {os.path.basename(path)}")

for path in SUPPS:
    doc = Document(path)
    if any(p.text.strip().startswith("Supplementary Methods S1") for p in doc.paragraphs):
        print("S1 이미 존재:", os.path.basename(path)); continue
    anchor = [p for p in doc.paragraphs if p.text.strip().startswith("Supplementary Fig. S1")]
    assert len(anchor) == 1, f"{os.path.basename(path)}: S1 캡션 앵커 {len(anchor)}건"
    a = anchor[0]
    for txt, bold in ((S1_BODY, False), (S1_HEAD, True)):
        new_p = copy.deepcopy(a._p)
        a._p.addprevious(new_p)
        np = Paragraph(new_p, a._parent)
        for i, r in enumerate(np.runs):
            r.text = txt if i == 0 else ""
        for r in np.runs:
            r.font.bold = bold
    doc.save(path)
    print("S1 삽입:", os.path.basename(path))
