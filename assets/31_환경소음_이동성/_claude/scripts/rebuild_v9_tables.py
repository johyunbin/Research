# -*- coding: utf-8 -*-
# v9 2단계(표): booktabs 재구축. 신규 T1(선행연구 비교) 삽입, 구 T1+T2 → 신규 T2(데이터) 통합,
# T5 = M1+M2 병렬(핵심), T6 = FDR 열 추가, T3/T4/T7 정리. 각 표 밑 Note 문단(9pt italic).
import os, sys
from docx import Document
sys.path.insert(0, os.path.dirname(__file__))
import docxtools_v9 as T

SRC = sys.argv[1] if len(sys.argv) > 1 else os.path.join(T.ROOT, "_claude", "step1_text.docx")
OUT = sys.argv[2] if len(sys.argv) > 2 else os.path.join(T.ROOT, "_claude", "step2_tables.docx")

doc = Document(SRC)

def cap(prefix):
    return T.find(doc, prefix)

def note_after(anchor_el, text):
    p = T.para_after(doc, anchor_el)
    T.fill_para(p, text, size=9, italic=True)
    return p

def replace_table_after(cap_p, old_tbl, rows, widths, aligns, font=9, header_rows=1, note=None):
    tbl = T.build_table(doc, rows, widths, aligns, font=font, header_rows=header_rows,
                        anchor_el=cap_p._element)
    if old_tbl is not None:
        T.delete_el(old_tbl)
    if note:
        note_after(tbl._element, note)
    return tbl

tables = [b for b in T.iter_blocks(doc) if not hasattr(b, "runs") and hasattr(b, "rows")]
assert len(tables) == 7, f"기존 표 {len(tables)}개(7 기대)"
oldT = dict(zip(range(1, 8), tables))

# ================= 신규 Table 1: 선행연구 비교 (§1.3 뒤) =================
p13 = T.find(doc, "그러나 이 빠르게 축적된 문헌은")
cap1 = T.para_after(doc, p13._element)
T.fill_para(cap1, "Table 1. Design features of representative COVID-19 urban-noise studies and the present study.")
rows1 = [
 ["Study", "City (measurement points)", "Exposure contrast", "Mobility measured?", "Sensor-bias handling"],
 ["Asensio et al. [19]", "Madrid (municipal noise network)", "Binary: lockdown vs. pre-lockdown", "No", "Absolute levels"],
 ["Basu et al. [20]", "Dublin (12 fixed stations)", "Binary: lockdown phases vs. baseline", "No", "Absolute levels"],
 ["Aletta et al. [21]", "London (short-term site measurements)", "Binary: lockdown vs. pre-lockdown", "No", "Absolute levels"],
 ["Rumpler et al. [22]", "Stockholm (central station)", "Binary: recommendation period vs. before", "No", "Absolute levels"],
 ["Manzano et al. [24]", "Granada (urban measurement points)", "Binary: lockdown vs. reference", "No", "Absolute levels"],
 ["Mishra et al. [28]", "Kanpur (campaign sites)", "Binary: lockdown vs. before", "No", "Absolute levels"],
 ["This study", "Seoul (1,123 IoT sensors, 421 neighbourhoods)", "Graded: repeated tightening and relaxation, 2020-2023",
  "Yes: daily neighbourhood de-facto population", "Within-sensor change; within-date two-way FE; calibrated-network cross-validation"],
]
t1 = T.build_table(doc, rows1, [1.30, 1.45, 1.45, 1.05, 1.25], "lllll", font=8.5,
                   anchor_el=cap1._element)
note_after(t1._element, "Note: measurement-point descriptions are indicative; see the cited papers for details.")

# ================= 신규 Table 2: 데이터 소스 통합 (구 T1 위치) =================
capA = cap("Table 1. Provenance and processing")
T.fill_para(capA, "Table 2. Data sources and processing.")
rows2 = [
 ["Component", "Dataset (provider)", "Resolution / processing", "Period", "Use in analysis"],
 ["Outcome:\nurban noise", "S-DoT city-wide IoT sensing network (Seoul Open Data Plaza, OA-15969)",
  "~1,100 fixed sensors; hourly broadband SPL (dB) aggregated to daily Leq,24h, Lday (06-21 h), Lnight (22-05 h); QC: ≥12 valid h/day, 20-95 dB; 2023 schema change reconciled ('mean' field)",
  "2020-04 to 2023-12", "Outcome; within-sensor change only (absolute levels never compared across sensors)"],
 ["Exposure:\nmobility", "De-facto ('living') population per administrative dong (Seoul Open Data Plaza, OA-14991)",
  "Hourly counts per 424 dong, averaged to daily daytime (06-21 h) and night-time (22-05 h) values; 619,464 dong-days, no missing",
  "2020-01 to 2023-12", "Dose lp = log(daily / dong's post-lifting baseline mean); within-dong relative change"],
 ["Policy\ncovariate", "Social-distancing implementation history (KDCA)",
  "Daily regime re-coded to continuous stringency 0-7 from business curfew hour and gathering cap",
  "2020-01 to 2022-04", "Auxiliary descriptor (Table 3); absorbed by date FE in M2"],
 ["Weather\ncovariate", "Open-Meteo ERA5 reanalysis at Seoul city centre (37.57° N, 126.97° E)",
  "Daily mean/max/min temperature, precipitation, maximum wind speed",
  "2020-2023", "M1 covariates; absorbed by date FE in M2"],
 ["Calendar", "Day of week, weekend, Korean public holidays (incl. substitutes), season",
  "Daily indicators", "2020-2023", "Weekend/holiday indicators as M1 covariates; absorbed by date FE in M2"],
 ["Comparison\nnetworks", "Calibrated official stations: national environmental-noise network (noiseinfo.or.kr; automatic daily road stations, manual quarterly general/road stations) and four roadside LAeq stations (OA-15473)",
  "Station-level standard LAeq", "2020-2024", "Drift diagnosis and absolute-offset comparison (Fig. 8, Supplementary Fig. S2)"],
]
replace_table_after(capA, oldT[1], rows2, [0.85, 1.75, 1.85, 0.75, 1.30], "lllll", font=8.5)
# 구 Table 2(모빌리티 provenance) 캡션+표 삭제
capB = cap("Table 2. Provenance of the mobility")
T.delete_el(capB); T.delete_el(oldT[2])

# ================= Table 3: 거리두기 타임라인 =================
cap3 = cap("Table 3. Seoul capital-area")
T.fill_para(cap3, "Table 3. Seoul capital-area social-distancing timeline and its quasi-continuous re-coding "
                  "(selected regime changes).")
rows3 = [
 ["Effective from", "Regime", "Business curfew (h)", "Gathering cap", "Stringency (0-7)"],
 ["2020-01", "Pre-COVID (normal)", "24", "none", "0"],
 ["2020-03-22", "1st intensive distancing", "21", "≤10", "4"],
 ["2020-05-06", "Daily-life distancing", "24", "none", "0"],
 ["2020-08-30", "Capital Level 2.5 (first 21 h cap)", "21", "≤50", "5"],
 ["2020-12-23", "Level 2.5 + 5-person ban", "21", "≤4", "6"],
 ["2021-02-15", "Level 2 + 5-person ban", "22", "≤4", "5"],
 ["2021-07-12", "Capital Level 4", "22", "≤4", "5"],
 ["2021-11-01", "With-COVID recovery", "24", "≤10", "1"],
 ["2021-12-18", "Special measures", "21", "≤4", "6"],
 ["2022-03", "Gradual easing", "23", "≤8", "3"],
 ["2022-04-18", "Full lifting", "24", "none", "0"],
]
replace_table_after(cap3, oldT[3], rows3, [1.10, 2.20, 1.10, 1.00, 1.10], "llccc", font=9,
    note=("Note: the stringency index re-codes the two enforceable components (business curfew hour and "
          "private-gathering cap) onto a 0-7 scale. The official tier system itself was redefined in "
          "2020-11 and 2021-07, so tier labels are not comparable over time and are never used as the dose; "
          "the measured mobility of each dong is the exposure throughout."))

# ================= Table 4: 기술통계 =================
cap4 = cap("Table 4. Descriptive statistics")
T.fill_para(cap4, "Table 4. Descriptive statistics of the analysis panel (1,248,794 sensor-days; 1,123 sensors; "
                  "421 dongs; 2020-2023).")
rows4 = [
 ["Variable", "Mean", "SD", "P5", "P95", "N"],
 ["Lday (dB)", "50.52", "7.20", "40.87", "65.54", "1,248,793"],
 ["Lnight (dB)", "47.58", "6.81", "38.46", "61.80", "1,248,623"],
 ["Leq,24h (dB)", "49.93", "7.07", "40.51", "64.76", "1,248,794"],
 ["Daytime mobility (log relative)", "0.01", "0.13", "-0.20", "0.17", "1,247,547"],
 ["Mean temperature (°C)", "12.53", "10.45", "-5.80", "26.50", "1,248,794"],
 ["Precipitation (mm)", "4.40", "12.18", "0.00", "27.80", "1,248,794"],
 ["Max wind (m/s)", "4.61", "1.68", "2.50", "7.81", "1,248,794"],
]
replace_table_after(cap4, oldT[4], rows4, [2.20, 0.86, 0.86, 0.86, 0.86, 0.86], "lrrrrr", font=9,
    note=("Note: mobility is the within-dong log change relative to the dong's post-lifting baseline; "
          "N = sensor-days."))

# ================= Table 5: M1 + M2 병렬 (핵심) =================
cap5 = cap("Table 5. Sensor fixed-effects regression")
T.fill_para(cap5, "Table 5. Mobility dose-response of urban noise: sensor fixed-effects (M1) and two-way "
                  "fixed-effects (M2) estimates. Coefficient (SE), dong-clustered. *p<0.05, **p<0.01, ***p<0.001.")
rows5 = [
 ["", "M1: sensor FE", "", "", "M2: sensor + date FE", "", ""],
 ["Term", "Lday", "Lnight", "Leq,24h", "Lday", "Lnight", "Leq,24h"],
 ["Mobility (log rel., own period)", "+1.130***\n(0.292)", "+2.283**\n(0.748)", "+1.258***\n(0.352)",
  "+0.648**\n(0.245)", "+0.504\n(0.631)", "+0.628*\n(0.303)"],
 ["Mean temperature", "-0.113***\n(0.004)", "-0.098***\n(0.003)", "-0.113***\n(0.004)", "-", "-", "-"],
 ["Temperature²", "+0.007***\n(0.000)", "+0.004***\n(0.000)", "+0.006***\n(0.000)", "-", "-", "-"],
 ["Precipitation (mm)", "+0.045***\n(0.001)", "+0.048***\n(0.001)", "+0.047***\n(0.001)", "-", "-", "-"],
 ["Max wind (m/s)", "+0.052***\n(0.005)", "+0.055***\n(0.005)", "+0.052***\n(0.005)", "-", "-", "-"],
 ["Rain day (0/1)", "+0.098***\n(0.010)", "+0.155***\n(0.009)", "+0.116***\n(0.010)", "-", "-", "-"],
 ["Weekend (0/1)", "-0.833***\n(0.029)", "-0.022\n(0.019)", "-0.686***\n(0.027)", "-", "-", "-"],
 ["Holiday (0/1)", "-1.176***\n(0.032)", "-0.186***\n(0.040)", "-0.999***\n(0.033)", "-", "-", "-"],
 ["Date fixed effects", "No", "No", "No", "Yes", "Yes", "Yes"],
 ["Weather + calendar", "Yes", "Yes", "Yes", "absorbed", "absorbed", "absorbed"],
 ["Within-R²", "0.117", "0.060", "0.112", "-", "-", "-"],
 ["Clusters (dong)", "420", "420", "420", "420", "420", "420"],
 ["N (sensor-days)", "1,247,546", "1,247,376", "1,247,547", "1,247,546", "1,247,376", "1,247,547"],
]
t5 = replace_table_after(cap5, oldT[5], rows5, [1.34, 0.86, 0.86, 0.86, 0.86, 0.86, 0.86],
                         "lcccccc", font=8, header_rows=2,
    note=("Note: all columns use own-period doses (daytime dose for Lday, night-time dose for Lnight, "
          "whole-day dose for Leq,24h). M1 controls for weather and weekend/holiday with sensor fixed "
          "effects; M2 adds date fixed effects, which absorb all date-common terms. SEs are clustered by "
          "dong, the level at which the dose is assigned."))
# 상단 헤더 병합: (M1: 1-3열) (M2: 4-6열)
r0 = t5.rows[0]
r0.cells[1].merge(r0.cells[2]).merge(r0.cells[3])
r0.cells[4].merge(r0.cells[5]).merge(r0.cells[6])

# ================= Table 6: 분할추정 + FDR =================
cap6 = cap("Table 6. Function-segmented")
T.fill_para(cap6, "Table 6. Function-segmented dose-response (M2, two-way fixed effects): β (dB per "
                  "log-unit mobility), 95% CI, nominal and BH-FDR-adjusted p-values. *p<0.05, **p<0.01 (nominal).")
rows6 = [
 ["Segment", "Group", "β", "95% CI", "p", "FDR p", "N"],
 ["By outcome", "Daytime Lday", "+0.648**", "[+0.17, +1.13]", "0.008", "0.033", "1,247,546"],
 ["", "Nighttime Lnight", "+0.504", "[-0.73, +1.74]", "0.42", "0.46", "1,247,376"],
 ["", "Day-night gap", "+0.019", "[-0.35, +0.39]", "0.92", "0.92", "1,247,375"],
 ["By land use (daytime)", "Commercial", "+0.335", "[-0.41, +1.08]", "0.38", "0.45", "430,160"],
 ["", "Mixed", "+0.858", "[-0.09, +1.81]", "0.076", "0.15", "425,584"],
 ["", "Residential", "+0.847", "[-0.42, +2.12]", "0.19", "0.25", "391,802"],
 ["By day type (daytime)", "Weekday", "+0.486", "[-0.16, +1.13]", "0.14", "0.21", "899,272"],
 ["", "Weekend/holiday", "+0.641", "[-0.17, +1.45]", "0.12", "0.21", "348,274"],
 ["By season (daytime)", "Winter (DJF)", "+0.658*", "[+0.09, +1.22]", "0.022", "0.066", "276,254"],
 ["", "Spring (MAM)", "+0.695**", "[+0.26, +1.12]", "0.002", "0.018", "277,158"],
 ["", "Summer (JJA)", "+0.527", "[-0.02, +1.08]", "0.060", "0.14", "344,417"],
 ["", "Autumn (SON)", "+0.758**", "[+0.20, +1.31]", "0.007", "0.033", "349,717"],
]
replace_table_after(cap6, oldT[6], rows6, [1.30, 1.35, 0.72, 1.18, 0.55, 0.60, 0.80],
                    "llrcrrr", font=8.5,
    note=("Note: all segments use the two-way (sensor + date) fixed-effects specification with "
          "dong-clustered SEs; FDR p = Benjamini-Hochberg adjusted over the pre-specified exploratory family of 12 tests."))

# ================= Table 7: 강건성 =================
cap7 = cap("Table 7. Robustness and falsification")
T.fill_para(cap7, "Table 7. Robustness and sensitivity checks for the headline two-way fixed-effects "
                  "dose-response.")
rows7 = [
 ["Check", "Specification", "Result", "Conclusion"],
 ["Permutation sensitivity", "Dong-level dose reshuffled across dongs within each date, broadcast to all "
  "sensors in the dong (300 shuffles)",
  "Null β = −0.001 ± 0.023; actual β = +0.648; two-sided p = (B+1)/(N+1) = 0.003", "Signal is not mechanical"],
 ["Weighting sensitivity", "Dong-day panel re-estimated with equal dong weights (main estimate is "
  "sensor-location weighted)",
  "β = +0.826 (SE 0.416, p = 0.047)", "Sign and magnitude preserved"],
 ["Nonlinearity", "Quadratic mobility term added to M2",
  "β₂ = −0.59 (p = 0.44)", "Linear approximation adequate"],
 ["Multiple comparisons", "BH-FDR over the 12 segment tests of Table 6",
  "Daytime, spring (MAM) and autumn (SON) remain significant (FDR < 0.05)", "Daytime effect robust"],
 ["Hour-count filter", "Sensor-days with more than 24 recorded hours excluded (5.9% of rows)",
  "β = +0.651 (SE 0.248)", "Aggregation artefacts negligible"],
]
replace_table_after(cap7, oldT[7], rows7, [1.15, 1.90, 2.05, 1.40], "llll", font=8.5)

doc.save(OUT)
tbls = [b for b in T.iter_blocks(doc) if hasattr(b, "rows") and not hasattr(b, "runs")]
print(f"표 재구축 완료: 총 {len(tbls)}개 표 → {OUT}")
