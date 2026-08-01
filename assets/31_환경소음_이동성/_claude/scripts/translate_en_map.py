# -*- coding: utf-8 -*-
# 영역 맵: (한국어 문단 마커, 영문 번역). 규칙: em/en dash 금지·AI 상투어 금지·영국식 철자·수치/인용 불변.
# 첨자 토큰(Lday·Lnight·Leq,24h·LAeq)은 docxtools_v9.SRE가 처리.

PARA = [
("도시가 사람들의 이동을",
 "How much quieter does a city become when its people move less? Much of sustainable urban policy, from "
 "remote work to travel-demand management and land-use reform, acts on mobility, yet the slope of this "
 "simple relationship has rarely been quantified. Environmental noise is regarded as the second-largest "
 "environmental burden on human health in cities after air pollution. The World Health Organization "
 "identifies it as a leading cause of disease burden in European cities and estimates that more than one "
 "million healthy life-years are lost to traffic noise every year in western Europe alone [1,2]. The health "
 "effects go beyond annoyance. Night-time noise fragments sleep [3], chronic exposure raises the risk of "
 "hypertension and ischaemic heart disease through autonomic and endocrine stress pathways [4,5], and "
 "daytime exposure causes annoyance and impairs cognitive performance [6,7]. These effects translate into "
 "substantial social and economic costs [8]."),

("따라서 소음은 도시의 지속가능성과",
 "Noise is therefore a central physical stressor for urban sustainability and liveability, and because "
 "its burden falls disproportionately on socioeconomically disadvantaged groups it is also a matter of "
 "environmental justice, although the scope of this study is limited "
 "to changes in physical sound levels. The soundscape approach, which treats sound as human experience in "
 "context rather than as decibels alone [9,10,11,12,13,14], has broadened this perspective to cover the "
 "acoustic comfort of public spaces [15] and the spatiotemporal variation of urban sound environments "
 "across land use and urban form [16]. Managing noise, however, requires an answer to a basic question. "
 "The lever that sustainability policy actually holds is not the vehicle fleet itself but the volume and "
 "placement of travel and activity. How sensitively, then, does urban noise respond to that lever, to "
 "human activity and mobility? This responsiveness is the starting point for judging any mobility- or "
 "land-use-based noise policy, and it has rarely been quantified."),

("이 반응성을 기대할",
 "There are clear physical grounds for expecting such responsiveness. Road traffic is the dominant source "
 "of urban noise, and sound levels rise with traffic volume [17]. But policy acts on more than traffic "
 "counts; it acts on the whole of human activity and movement, and the relationship between citywide "
 "activity and noise has almost never been estimated as a causal slope. The obstacle is identification. In "
 "a normally functioning city, mobility does not vary exogenously. It is bound up with land use, time of "
 "day, day of week and weather, the same factors that drive noise. Simple correlations between observed "
 "mobility and noise are therefore confounded, and they say little about how much noise would change if "
 "mobility itself changed."),

("기존의 교통소음 예측 모델",
 "Established traffic-noise models, such as the CNOSSOS-EU standard and land-use regression [17,18,29], "
 "predict noise from traffic volume, road geometry and land use, but they are essentially cross-sectional "
 "and static. They reproduce the spatial pattern that busy places are loud, yet they cannot answer the "
 "dynamic dose-response question of how much quieter a place becomes when mobility falls. Filling this gap "
 "requires exogenous, graded variation in mobility together with dense noise observation, a combination "
 "that ordinary urban conditions almost never supply."),

("코로나-19 대유행은",
 "The COVID-19 pandemic supplied exactly this combination as a natural experiment. Lockdowns and "
 "distancing cut human mobility exogenously, and a series of studies documented reductions of several "
 "decibels in measured urban noise (4-6 dB in Madrid [19], with further reports from Dublin [20], London "
 "[21], Stockholm [22], Montreal [23], Granada [24] and Kanpur [28]). Perception surveys in Argentina "
 "[25], Lima [26] and Italy [27] likewise reported changes in the sound environment during "
 "lockdowns. Air quality shifted at the same time, with sharp drops in nitrogen dioxide and carbon dioxide "
 "[30,31,32], evidence of how broadly confinement reshaped the urban environment."),

("그러나 이 빠르게 축적된 문헌은",
 "This rapidly accumulated literature, however, shares four structural limitations that motivate the "
 "present study. First, most studies compare a binary lockdown period against a reference period, which "
 "cannot recover a graded dose-response curve across levels of restriction; whether noise responds "
 "proportionally to mobility, and with what slope, remains unknown. Second, most studies attribute the "
 "observed reductions to reduced traffic in narrative terms, without measuring mobility and linking it to "
 "noise quantitatively; the exposure is assumed rather than observed. A few studies did pair noise with "
 "measured vehicle traffic (four noise sensors with nearby inductive-loop counts in Porto [52], a "
 "floating-car traffic and emission simulation in Rome [53], and 23 recording sites with a single "
 "traffic-count station in Bochum [54]), but their coverage stays at a handful of locations or at vehicle "
 "flows alone, and to our knowledge none links citywide human activity to noise at the neighbourhood level. "
 "Third, measurement points are few (typically 4-70) and concentrated in western cities, so the "
 "fine-grained spatial heterogeneity of dense East Asian cities goes unobserved. Fourth, almost all "
 "studies compare absolute sensor readings directly, leaving room for network calibration bias and "
 "multi-year drift to contaminate the results. Compared along these four axes, graded exposure, measured "
 "mobility, high-density observation and drift-aware identification, no representative study satisfies "
 "more than two of them at once (Table 1), and none brings all four together in a dense East Asian "
 "city. This study addresses that gap."),

("최근 두 가지 데이터 혁신이",
 "Two recent developments in urban data make the first three requirements attainable. The first is "
 "de-facto population estimated from mobile-network signals, which records where people actually are at a "
 "given moment, at fine spatial and temporal resolution, rather than where they reside [35,36,37,38,39]. "
 "The second is the low-cost acoustic sensor network. SONYC [40], low-cost monitoring devices [41], "
 "wireless acoustic sensor networks [42,43], the participatory NoiseCapture platform [44] and smart-city "
 "sensing programmes [45,46] now sustain continuous observation of urban sound at hundreds to thousands "
 "of points. Korea's graded distancing, the citywide S-DoT noise network in Seoul (about 1,100 sensors) "
 "and dong-level de-facto population provide these ingredients together in a single city, a rare "
 "combination."),

("그러나 저가 IoT 소음망에는",
 "Low-cost noise networks, however, carry a structural problem that the COVID noise literature has "
 "largely ignored: each sensor has its own calibration offset, and readings drift slowly with age and "
 "environmental exposure. The same problem is well documented for low-cost air-quality networks [47]. "
 "These devices are simply not calibrated to measure absolute sound pressure levels. Ignoring this and comparing "
 "absolute levels over time or across space produces spurious patterns or spurious conclusions, and this "
 "study demonstrates a concrete case. What remains is the fourth requirement: removing this contamination "
 "at the design stage of identification rather than at the stage of interpretation."),

("본 연구는 한국의 단계적 사회적 거리두기를",
 "This study combines Korea's graded social distancing as a natural-experiment setting, Seoul's S-DoT "
 "noise sensors as the outcome, and dong-level de-facto population as the measured mobility dose, and "
 "estimates the mobility dose-response of urban noise. Specifically, we adopt a two-way fixed-effects "
 "strategy [48,49] within a natural-experiment framework [50] that identifies the slope only from "
 "within-sensor changes and from variation across dong within the same date, never from absolute levels, "
 "and we check the drift and resolution limits of the low-cost network against official monitoring "
 "stations. In doing so we extend the descriptive work on the spatiotemporal variation of Seoul's sound "
 "environment [16] towards quantitative, quasi-experimental estimation; the position of this study relative "
 "to prior work is summarised in the last row of Table 1. We ask three questions, and answering them "
 "doubles as an answer to a methodological one: under what design can low-cost IoT noise networks be "
 "trusted at all?"),

("RQ1.",
 "RQ1. How much does urban noise fall when measured mobility falls, and with what dose-response slope?"),

("RQ2.",
 "RQ2. How does the response differ by land use, day and night, day type and season?"),

("RQ3.",
 "RQ3. How does noise rebound as restrictions are relaxed and lifted, and does the relationship also "
 "appear in long-run neighbourhood averages (the spatial cross-section)?"),

("본 연구의 기여는 네 가지다",
 "The study makes four contributions. First, to our knowledge it provides the first quantitative estimate "
 "of a graded dose-response of urban noise to activity that is measured, dong by dong, across an entire "
 "city rather than assumed. Second, it combines a high-density IoT network of about 1,100 sensors with "
 "administrative-neighbourhood de-facto population to identify the response at fine urban resolution. "
 "Third, it extends a literature concentrated on western cities to Seoul, a high-density East Asian "
 "megacity. Fourth, it quantifies the calibration and drift limits of low-cost noise sensing and derives "
 "a design principle: rely on within-sensor change and same-day comparisons across neighbourhoods, not on "
 "absolute levels."),

("연구지역은 서울특별시",
 "The study area is the whole of Seoul, covering all 25 autonomous districts. With about 9.6 million "
 "residents on 605 km², Seoul is among the densest megacities in the world, and its fine grid of 421 "
 "administrative neighbourhoods (dong) over intensely mixed land use makes it well suited to capturing "
 "within-city contrasts (Fig. 1). The analysis runs from 2020-04-01, the first day of S-DoT noise "
 "provision, to 2023-12-31, extending well past the full lifting of distancing on 2022-04-18 so that the "
 "return of noise to its usual level, the rebound, is observed in full. The design rests on three "
 "elements. Korean distancing tightened and relaxed restrictions repeatedly, moving mobility gradually "
 "and for reasons external to the noise environment, and we treat this variation as a natural-experiment "
 "setting. Because each low-cost sensor reads with its own constant bias, as explained below, absolute "
 "levels are never used; the outcome is always the change of each sensor relative to its own norm. And "
 "the response is identified from how much mobility differed across neighbourhoods within the same day."),

("결과변수는 서울시가",
 "The outcome variable is noise from the Smart Seoul Data of Things (S-DoT), the city's permanently "
 "operated IoT sensing network. About 1,100 stations record hourly broadband sound pressure levels in dB, and "
 "all data are openly downloadable without authentication (Table 2). We aggregate the hourly values into "
 "daily energy-equivalent levels, computing the full-day Leq,24h, the daytime Lday (06-21 h) and the "
 "night-time Lnight (22-05 h) as 10·log10 of the mean of 10^(L/10), and retain only sensor-days with at "
 "least 12 observed hours and values in a physically plausible range (20-95 dB). In 2023 the data schema "
 "changed and noise was split into maximum, mean and minimum fields; we use the mean. One caveat matters "
 "throughout. S-DoT is a general-purpose urban sensing node rather than an environmental-noise monitor: "
 "it reports uncalibrated broadband decibels (the values carry no A-weighted LAeq label, and the "
 "microphone specification, the frequency and time weighting, and the rule aggregating 2-minute raw "
 "records into hourly values are not documented in the public specification), so absolute readings cannot "
 "be compared across sensors. Day and night windows follow the calendar date; Lnight combines the 00-05 h "
 "and 22-23 h records of the same calendar day. We therefore use only within-sensor changes over time, "
 "and we check their reliability against the official calibrated LAeq networks (Fig. 8, Supplementary "
 "Fig. S2)."),

("이동량 노출은 서울",
 "Mobility exposure is measured with Seoul's de-facto ('living') population: the number of people "
 "actually present in each dong at a given time, estimated from mobile-network signals rather than from "
 "residence records (Table 2). Aggregating the hourly stock to daily means (daytime 06-21 h and "
 "night-time 22-05 h averages) yields 424 dong × 1,461 days = 619,464 dong-days with no missing values, "
 "of which the 421 dong hosting S-DoT sensors enter the analysis. One feature of these data shapes the "
 "dose definition. The citywide total of de-facto population is almost conserved, because people who stay "
 "home are still counted where they are; distancing shows up not in the total but in spatial "
 "redistribution. Under the strongest restrictions the daytime population of commercial and business dong "
 "fell by up to about 35%, depending on the dong, while residential dong gained up to 37% (averaged over "
 "the whole distancing era, neighbourhood reductions mostly lie in the 0-13% range; Fig. 7a). We "
 "therefore define the dose as each dong's change relative to its own norm, lp = log(daily de-facto "
 "population / the dong's post-lifting baseline mean), so that it captures how far a dong departs from "
 "its own normal and is unaffected by level differences across dong or sensors. Two clarifications are in "
 "order. First, de-facto population is a stock of presence, not a flow of movement, so 'mobility' in this "
 "paper is an operational term for changes in activity presence. Second, the baseline period is 2022-07 "
 "to 2023-12, chosen to exclude the transitional weeks immediately after full lifting (2022-04 to 06). "
 "Within a given date, the standard deviation of the dose across dong averages 0.115 log-units, ample "
 "variation for identification under date fixed effects."),

("센서 위치정보에는",
 "Sensor metadata contain street addresses but no administrative-neighbourhood name, so each sensor was "
 "assigned to a dong by point-in-polygon matching of its coordinates against public boundary data (the "
 "vuski/admdongkor GeoJSON, 2022-01 vintage). The de-facto population data and the boundary file use "
 "slightly different dong-code systems in a few reorganised areas of Gangbuk and Gangdong; these were "
 "reconciled by name matching. In the result, 1,165 of 1,170 sensors (99.6%) were assigned to a dong, and "
 "the assigned dong agreed with the district recorded in the street address in 99.6% of cases. After the "
 "noise quality filters (at least 12 h per day, 20-95 dB) and the merge with daily mobility, the final "
 "panel comprises 1,123 sensors in 421 dong (1,248,794 sensor-days); the main two-way fixed-effects model "
 "uses the 1,122 sensors and 420 dong with non-missing daytime mobility (1,247,546 sensor-days)."),

("거리두기 '단계'",
 "Distancing 'tiers' (Levels 1, 2, 2.5 and 4, among others) are unsuitable as an exposure variable: the "
 "tier system was redefined twice (2020-11 and 2021-07), the spacing between tiers is not uniform, and "
 "the same tier corresponded to different mobility at different times. We instead re-coded the daily "
 "regime into a continuous stringency index from 0 (no restriction) to 7 (strongest), built from two "
 "concrete enforceable components, the business closing hour and the private-gathering cap (Table 3). "
 "This index remains auxiliary; the primary exposure throughout is measured mobility (de-facto "
 "population). Weather comes from the Open-Meteo ERA5 daily reanalysis (temperature, precipitation, "
 "wind). The calendar controls entering M1 are weekend and Korean public-holiday indicators; day-of-week "
 "and season are available in the data, but under M2 the date fixed effects absorb all calendar effects, "
 "so separate terms are unnecessary."),

("저가 S-DoT 센서는",
 "A low-cost S-DoT sensor is not precision-calibrated: presented with the same sound, each unit records a "
 "level that is systematically high or low by its own constant amount. Because this offset differs across "
 "sensors and is unknown, comparing absolute decibels between two sensors is meaningless. If we look only "
 "at how a sensor changes over time, however, for instance how today compares with that sensor's own "
 "average, the constant offset cancels in the subtraction. Every analysis in this paper therefore rests "
 "on within-sensor relative change rather than on absolute levels. Statistically this takes the form of a "
 "fixed-effects model: subtracting each sensor's mean removes its idiosyncratic offset (we use the "
 "mathematically equivalent within-transformation rather than 1.25 million dummy variables). Standard "
 "errors are made robust to within-cluster correlation and are clustered by dong, the level at which the "
 "mobility dose is assigned, in both the sensor fixed-effects and the two-way models [48]. Fig. 2 "
 "summarises the identification logic."),

("주된 분석은 두 단계의",
 "The main analysis proceeds in two steps. (M1) The sensor fixed-effects model removes each sensor's mean "
 "and explains its noise variation with dong mobility, weather (temperature, temperature squared, "
 "precipitation, wind, a rain indicator) and weekend and holiday indicators. This controls the visible "
 "temporal confounders while leaving temporal variation available for estimation. (M2), the core model of "
 "the paper, adds date fixed effects. A date fixed effect absorbs everything the whole city experienced "
 "in common on a given day: that day's weather, the day-of-week and holiday pattern, secular trends in "
 "city noise, network-wide sensor drift and nationwide distancing itself. What remains is a single source "
 "of information: how much mobility differed across dong within the same day. M2 therefore answers one "
 "question. On a given day, did sensors in neighbourhoods that emptied more than their norm become "
 "quieter than sensors in neighbourhoods that emptied less? One consequence of the design is that any "
 "variable taking a single value per day citywide (total subway ridership, the stringency index) is "
 "absorbed by the date effects and cannot be estimated; only mobility measured per dong remains a valid "
 "exposure. Formally, M2 is Y_sdt = α_s + λ_t + β·x_dt + ε_sdt, and the identifying assumption is that "
 "residual noise shocks within a date are unrelated to dong mobility. Shocks at the dong-by-date level "
 "that move mobility and noise together, such as construction or local events, can violate this "
 "assumption, so we read β conservatively as a within-date conditional association rather than a causal "
 "effect, and we keep that reading throughout. The two sets of fixed effects are removed by iterating the "
 "alternating subtraction of sensor and date means to convergence."),

("이 두 모형 위에 보조 분석을",
 "Several auxiliary analyses sit on top of these two models. (M3) Segmented estimation: M2 is "
 "re-estimated by outcome (daytime, night-time, day-night gap), by activity type (commercial, mixed and "
 "residential terciles of each dong's daytime-to-night-time population ratio; Supplementary Fig. S3; the "
 "classification uses only post-lifting level information rather than the daily exposure variation, which "
 "limits circularity, and although it is in substance an activity-profile grouping we call it 'land use' "
 "for brevity), by weekday versus weekend, and by season, to locate where the response arises. (M4) "
 "High-loss versus low-loss trajectories: dong that emptied strongly (mostly commercial) are compared "
 "week by week with dong that barely emptied (mostly residential). Common factors such as weather, season "
 "and sensor drift largely cancel in the difference. Because the groups are defined from realised "
 "mobility during the restriction era and both series are normalised to the post-lifting window, "
 "convergence to zero after lifting is partly built in, and we therefore interpret this as a descriptive "
 "comparison rather than a formal event study. (M5) Official-network comparison: the reliability of S-DoT "
 "absolute levels and multi-year trends is checked against the calibrated national environmental-noise "
 "networks (automatic and manual stations) and four permanent roadside LAeq stations, using level offsets "
 "at nearby station pairs and 2020-2023 trends to diagnose sensor drift and the 2023 schema change. (M6) "
 "Robustness and sensitivity: a permutation check that reshuffles the dong-level dose across dong within "
 "each date (300 shuffles, applied identically to all sensors of a dong because the dose is assigned at "
 "the dong-day level); an equal-weight re-estimation on the dong-day panel, to confirm that sensor-rich "
 "dong do not dominate; a quadratic term, to probe nonlinearity; and Benjamini-Hochberg FDR adjustment "
 "over the segmented tests of M3. Finally, all dong-level spatial analyses use outlier-robust statistics "
 "(Theil-Sen regression, Spearman rank correlation) and are restricted to dong with at least two sensors, "
 "so that no small set of extreme neighbourhoods drives the results. All computations use Python 3.13 "
 "(pandas, NumPy, SciPy, statsmodels); the two-way fixed effects are absorbed by iterated demeaning, and "
 "the robust estimators come from SciPy."),

("분석에 사용한 변수들의 기술통계는",
 "Table 4 reports descriptive statistics. In the sensor fixed-effects model (M1, Table 5), a dong's "
 "daytime mobility and the daytime noise of its sensors move together (β=+1.130 dB per log-unit, SE "
 "0.292, p<0.001; standard errors are clustered at the dong level, where the dose is assigned). The "
 "control variables behave physically: temperature is U-shaped (louder when cold or hot), precipitation "
 "adds +0.045 dB/mm and rainy days +0.10 dB, weekends run 0.83 dB quieter and holidays 1.18 dB quieter "
 "(all p<0.001). Adding date fixed effects, which strip out everything the city shared on a given day "
 "(weather, trends, drift), shrinks the response: in the core model (M2, Table 5) the daytime slope is "
 "β=+0.648 dB per log-unit (95% CI 0.17-1.13, p=0.008) and the full-day slope is β=+0.628 (p=0.038). In "
 "concrete terms, when a dong's daytime mobility falls 30% below its norm, its daytime noise falls by "
 "about 0.23 dB; a 50% drop corresponds to about 0.45 dB. The dose-response estimated purely from "
 "same-day differences across neighbourhoods is statistically solid but small (RQ1)."),

("이 효과의 기능적 위치를 분해하면",
 "Splitting the estimate by segment (Fig. 3, Table 6, RQ2) shows the individually significant responses "
 "concentrating in the daytime. Across outcomes, only daytime noise is significant; night-time noise and "
 "the day-night gap are not (their interpretation follows in the next section). Across activity types, "
 "commercial (+0.34), mixed (+0.86) and residential (+0.85) dong all show positive slopes, but the split "
 "samples lose power, none is individually significant, and an interaction test finds no significant "
 "difference across types; the response appears broadly uniform. Weekdays (+0.49) and weekends (+0.64) "
 "are similar. All four seasons carry the same positive sign (DJF +0.66, MAM +0.69, JJA +0.53, SON "
 "+0.76), with winter, spring and autumn nominally significant; after FDR adjustment, spring and autumn "
 "remain so. The response is consistent in direction through the year rather than an artefact of one "
 "season."),

("주간과 야간을 나눠 보면",
 "Separating day and night shows that the significant response is confined to the daytime outcome (Fig. "
 "3b). Night-time noise appeared to respond strongly to night-time mobility under sensor fixed effects "
 "alone (β=+2.28, p=0.002), but the association loses significance once date fixed effects are added "
 "(β=+0.50, 95% CI -0.73 to +1.74, p=0.42), which indicates that much of the apparent night-time response "
 "reflected time-varying confounding shared across the city (drift, season, secular trends). The daytime "
 "association survives the same test. The wide night-time interval does, however, contain the daytime "
 "coefficient (+0.65), so a difference between the day and night responses cannot be claimed. Consistent "
 "with this, the day-night gap, regressed on the same daytime dose, has a coefficient of essentially zero "
 "(β=+0.02, p=0.92), which also fits daytime and night-time noise moving together in similar measure with "
 "daytime activity. Individually detectable signals are thus confined to the daytime outcome, while the "
 "day-night difference itself remains statistically unresolved."),

("이동량이라는 노출이 시간과 공간에서",
 "How the exposure moved through time and space is central to the design (Fig. 4). The citywide total of "
 "de-facto population barely changes (staying home still counts as presence), but where that population "
 "sits is reallocated sharply with each phase of distancing. In the strongest phases, commercial and "
 "business dong lost up to about 35% of their daytime population, depending on the dong "
 "(phase-window averages; the distribution of distancing-era average reductions appears in Fig. 7a), "
 "while residential dong gained. Central commercial dong emptied most under the strictest regimes (the "
 "5-person ban of December 2020, capital-area Level 4 of July 2021) and recovered through the with-COVID "
 "relaxation (November 2021) and full lifting (April 2022)."),

("이 기능적 차등은 시간축에서도",
 "The same functional contrast is visible over time (Fig. 5, RQ3). Weekly daytime mobility in commercial "
 "dong stayed below baseline throughout the restriction era while residential dong stayed above, and the "
 "two converged after the 2022 lifting (Level 4 period medians: commercial 0.98 versus residential 1.05). "
 "Drift-robust relative noise (group mean minus city mean) likewise shows commercial dong running quieter "
 "than the city average during restrictions. The dynamics of regime switches are read from the weekly gap "
 "between high-loss and low-loss dong (a descriptive comparison, Fig. 6; the groups are defined ex post "
 "from realised mobility and both series are normalised to the post-lifting window, so this is not a "
 "formal event study). The weekly noise gap tilts negative during restrictions and converges to zero in "
 "the normal period (shown with dong-clustered 95% confidence bands), and the weekly correlation between "
 "the mobility gap and the noise gap is Pearson +0.44 and Spearman +0.46."),

("반면 동을 하나의 점으로 보고",
 "By contrast, the long-run spatial cross-section, which treats each dong as a single point and compares "
 "distancing-era averages, shows no robust association between mobility loss and noise change (Fig. 7). "
 "Across all analysable dong (those with at least two sensors), the rank correlation between mobility "
 "reduction and noise change relative to the city mean is essentially zero (Spearman ρ ≈ -0.05). Among "
 "commercial dong the ordinary correlation looks positive (Pearson r=+0.38), but that value is driven by "
 "a few extreme neighbourhoods; robust statistics on the same dong leave almost nothing (Spearman "
 "ρ=+0.08, Theil-Sen slope near zero; Fig. 7c, detail in Supplementary Fig. S1). The mobility-noise "
 "signal, in other words, lives in same-day differences across neighbourhoods (M2, M4) and does not "
 "survive being painted onto a map of long-run averages."),

("각 센서의 자기 대비 변화",
 "Plotting each sensor's change from its own post-lifting norm (ΔLday) produces an apparent paradox: "
 "daytime noise during the strictest restrictions comes out higher than in the normal period (+0.85 in "
 "the strictest phases versus -0.04 after lifting). The diagnosis is not a COVID effect but multi-year "
 "movement of absolute levels (Fig. 8a). Among the 842 sensors operating through all four years, the "
 "annual mean Leq,24h declined steadily from 49.4 dB (2020) to 47.6, 47.2 and 47.1 dB (2023), with 85% of "
 "sensors trending downward; that decline is what makes 2020-21 look relatively loud. (There is no level "
 "jump at the 2023 schema change: +0.02 dB.)"),

("이 하락이 실제 소음 변화가 아니라",
 "We checked this decline against the calibrated official environmental-noise network (the national "
 "noise information system). Absolute levels disagree substantially: near official stations, S-DoT reads "
 "on average 11.7 dB below the standard daytime LAeq, and about 16 dB below at roadside sites (pairs "
 "matched within 500 m across non-simultaneous years, 2022 versus 2024; Fig. 8b, Supplementary Fig. S2), "
 "which quantifies why S-DoT absolute values cannot be read as standard noise indicators. The calibrated "
 "networks' own multi-year trends, meanwhile, are flat or rising (Fig. 8a): the automatic daily roadside "
 "network (9 stations) moved +0.04 dB from 2020 to 2023, while the quarterly manual networks rose +1.93 "
 "dB at 91 general sites and +1.91 dB at 60 roadside sites (the separate four-station roadside LAeq "
 "series moves the same way). Constant offsets cancel in trend differences, so a 2.3 dB decline unique to "
 "S-DoT while calibrated networks hold or rise is a pattern consistent with downward sensor drift rather "
 "than a real environmental change (residual confounding, such as differing seasonal composition across "
 "years, cannot be excluded entirely). That the S-DoT decline concentrates in quiet non-roadside areas, "
 "where the calibrated residential network actually rose, strongly suggests that the decline was not "
 "real. The calibrated rise "
 "deserves note for another reason: calibrated noise was lowest in 2020, when mobility was most "
 "depressed, and climbed through 2023 as activity recovered, matching the direction of our "
 "dose-response. The networks that can be trusted in multi-year absolute terms thus support the sign of "
 "our estimate; the series that moves the other way is the drift-contaminated S-DoT trend. Above all, our "
 "dose-response is identified from same-day differences across dong, with date fixed effects absorbing "
 "all multi-year and seasonal variation, and makes no use of the S-DoT multi-year trend. This is "
 "precisely why the analysis relies on within-sensor change and same-day comparison (M2) instead of "
 "absolute or time-series contrasts."),

("주력 추정치가 기계적 산물이 아님을",
 "Five sensitivity checks establish that the headline estimate is not a mechanical artefact (Fig. 9, "
 "Table 7). First, reshuffling the dong-level dose across dong within each date (300 shuffles, applied "
 "identically to all sensors of a dong) yields placebo coefficients tightly distributed around zero (mean "
 "-0.001, SD 0.023); the actual estimate of +0.65 lies entirely outside that distribution (two-sided "
 "p=(B+1)/(N+1)=0.003). Because the date-by-date shuffle does not preserve the temporal persistence of "
 "each dong's dose, we read this p-value as a diagnostic check rather than formal inference. Second, "
 "re-estimating on the dong-day panel with equal dong weights gives β=+0.83 (SE 0.42, p=0.047), "
 "preserving sign and magnitude (the main estimate is a sensor-location-weighted slope). Third, a "
 "quadratic mobility term is not significant (β²=-0.59, p=0.44), supporting the linear approximation "
 "(Fig. 9b). Fourth, Benjamini-Hochberg FDR adjustment across the 12 segmented tests of Table 6 leaves "
 "the daytime, spring and autumn responses significant (FDR<0.05). Finally, excluding sensor-days with "
 "more than 24 recorded hours (5.9% of rows, a possible duplicate-transmission artefact) leaves the "
 "estimate essentially unchanged at β=+0.651 (SE 0.248). The same-day cross-neighbourhood signal is "
 "therefore robust to implementation, weighting, functional form and aggregation rules."),

("본 연구는 측정 이동량에 대한 도시소음의",
 "This study provides, to our knowledge, the first quantitative estimate of a graded dose-response of "
 "urban noise to measured mobility. Four findings summarise the results. First, a statistically solid "
 "positive conditional association links a dong's daytime mobility to its daytime noise, but it is small: "
 "a 30% fall in mobility corresponds to roughly 0.2-0.3 dB, and the association carries through to the "
 "full-day Leq,24h (β=+0.628). Second, individually significant responses appear only for the daytime "
 "outcome (the day-night difference itself remains unresolved), with the same positive sign in all four "
 "seasons and significance surviving FDR adjustment in spring and autumn; the apparent night-time "
 "association dissolves under strict identification as a time confound. Third, the signal appears only in "
 "same-day cross-neighbourhood contrasts and in the high-loss versus low-loss trajectories; it vanishes "
 "in the long-run spatial cross-section. Fourth, the multi-year absolute levels of the low-cost network "
 "are contaminated by sensor drift and uneven recovery, so neither time-series nor absolute comparisons "
 "can be trusted; only designs built on within-sensor change and cross-neighbourhood comparison remain "
 "valid. The upward multi-year trend of the calibrated networks, which tracks recovering activity, "
 "independently supports the direction of the estimated response."),

("우리의 추정치는 봉쇄기",
 "Our estimate stands in contrast to the several-decibel reductions reported for lockdowns "
 "[19,20,21,22,23,24]. Those studies mostly compare absolute noise at particular busy locations before "
 "and during confinement, an approach that readily generalises the extreme changes of the loudest, "
 "most-emptied spots to the whole city. Here the marginal association (the average noise change per unit "
 "change in mobility) is estimated across roughly 1,100 locations, from within-sensor changes, along the "
 "whole range of mobility, over the full period; a citywide average response is bound to come out "
 "smaller. One qualification is essential. The core model identifies the slope only from same-day "
 "differences across dong, so the effect of the common citywide component of reduced mobility, everyone "
 "moving less at once, is not identified in this design. We offer neither an upper nor a lower bound for "
 "that component, and the estimate should be read strictly as the response to relative, "
 "neighbourhood-level changes. The small magnitude also fits two physical facts: much urban noise comes "
 "from background sources such as through-traffic and arterials, which empty little, and the 1 dB "
 "resolution of S-DoT adds measurement noise that widens standard errors. The popular impression that "
 "lockdowns made cities several decibels quieter is therefore likely an overstatement when read as a "
 "mobility dose-response; the graded response that survives strict identification is real but small."),

("두 추정량은 서로 다른",
 "The two kinds of estimate answer different questions, so rather than comparing them head-on it is more "
 "useful to push our slope through the exposure of the earlier studies. Taking as the dose the largest "
 "daytime emptying observed for commercial dong in the strictest phases (about 35%, Fig. 4), the "
 "predicted reduction at β=+0.648 is about 0.28 dB. The drop in the daytime coefficient from +1.13 to "
 "+0.65 when date fixed effects are added reflects the removal of date-common confounding (citywide "
 "trends, network-wide drift, weather); it does not by itself measure any common mobility effect. The "
 "gap between such figures and Madrid's 4-6 dB [19] then decomposes into three parts: (1) the effect of "
 "the citywide common decline in mobility, which date fixed effects absorb and this design does not "
 "identify; (2) site selection, since monitoring points concentrated in the busiest streets carry doses "
 "and responsiveness far above the city average; and (3) background trends and sensor changes folded "
 "into absolute before-and-after comparisons. Of the several decibels reported, the share explained by "
 "relative cross-neighbourhood mobility differences is around 0.3 dB; the remainder belongs to "
 "components this design deliberately does not identify, and to site selection. In perceptual terms a "
 "0.65 dB shift lies below what an individual listener would notice, though as a movement of a long-run "
 "exposure indicator for an entire urban population it retains meaning."),

("효과가 나타나는 양상은 그 메커니즘",
 "The pattern of the response is consistent with its likely mechanism. A significant response appears in "
 "the daytime outcome and keeps its direction through the year, which fits daytime noise being tied "
 "directly to human activity (commerce, commuting, visits), whereas night-time noise carries a larger "
 "share of fixed installations and arterial through-traffic and may respond weakly to changes in "
 "occupancy (with the caveat, again, that the day-night difference is not statistically settled). The "
 "disappearance of the apparent night-time association under stricter identification is a warning "
 "against reading simple time trends as causal. The rough uniformity of the response across activity "
 "types suggests that the mobility-noise link is not confined to one kind of place but spread thinly "
 "across the city. The absence of a clean long-run spatial gradient deserves emphasis. It does not "
 "indicate an absent effect; it indicates that long-run neighbourhood averages are dominated by "
 "idiosyncratic factors, construction, road changes and sensor-specific drift. The signal lives in fast "
 "same-day variation between neighbourhoods and is invisible to simple spatial mapping. Indeed the "
 "briefly positive correlation among commercial dong turned out to be an artefact of a few extreme "
 "neighbourhoods under Pearson correlation, and it disappears under robust statistics: a compact "
 "demonstration of how easily spatial analysis of dense IoT data can mislead."),

("다만 야간의 비유의를",
 "Reading the night-time null as 'night noise is insensitive to activity' would nevertheless be "
 "premature; three alternative explanations run in parallel. One is statistical power. The night-time "
 "standard error (0.63) is about 2.6 times the daytime one (0.245), and at that precision an estimated "
 "coefficient would need to reach about 1.2 dB to attain two-sided 5% significance, so a night-time "
 "response equal in size to the daytime one would go undetected in these data. A second is exposure "
 "variation. Because the practical instruments of distancing were closing hours and gathering caps, the "
 "relative variation of night-time presence across dong may be narrower than in the daytime, and a "
 "smaller dose variance estimates the same true slope less precisely. A third is measurement error. The "
 "daily quality filter imposes no minimum on within-window hours, so some Lnight values rest on few "
 "observations. If the resulting outcome error is conditionally mean-zero, it inflates the residual "
 "variance and widens standard errors rather than attenuating the coefficient; non-random hour coverage "
 "could, however, bias it. A similar precision problem extends to the long-run spatial cross-section: "
 "sensor-specific drift adds a large random component to the outcome, while error in the estimated "
 "mobility change can attenuate the exposure coefficient. The spatial null is therefore inconclusive "
 "rather than evidence of an absent effect; it is what a cross-section with error in both exposure and "
 "outcome would be expected to produce."),

("본 연구의 방법론적 기여는",
 "The methodological lessons matter as much as the estimates. Attempts to measure urban environmental "
 "effects with dense low-cost noise networks run into three pitfalls: (1) each sensor's idiosyncratic bias must "
 "be cancelled by designs that use only within-sensor change; (2) multi-year drift, staggered "
 "installation and uneven recovery make absolute levels incomparable over time and space, which forces "
 "same-day cross-neighbourhood comparison; and (3) exposure variables that take one value per day "
 "citywide are fragile to time confounding, so exposures must be measured below the city level. The "
 "stable spatial null of the long-run cross-section is itself instructive: painting uncalibrated "
 "absolute levels onto maps can produce spurious patterns, or spurious blankness. These lessons apply to "
 "the spreading families of low-cost noise monitoring, SONYC [40], wireless acoustic sensor networks "
 "[42], NoiseCapture [44] and low-cost devices generally [41], and to policy uses of smart-city sensor "
 "data at large [45,46]. The procedures shown here (fixed effects, outlier-robust statistics, "
 "cross-checks against official networks) are offered as a template for future IoT-based urban "
 "environmental research."),

("계획·정책 측면에서",
 "For planning and policy, the direct noise co-benefit of managing mobility and travel demand is "
 "probably smaller than commonly assumed: even lockdown-scale collapses in mobility produced differences "
 "of only about 0.3 dB at the level of relative neighbourhood change (the citywide common component lies "
 "outside this design). The estimate applies to relative dong-level changes; it does not transfer "
 "directly to policies that cut total citywide travel, or to street-scale interventions that close a "
 "single road. Fifteen-minute-city policies in particular redistribute activity rather than reduce it, "
 "so their noise effect differs in kind from the component identified here. What survives these caveats "
 "is the conclusion that the mobility channel alone cannot be counted on for reductions of several "
 "decibels. Mobility policies remain justified on their own grounds, but noise targets call for source- "
 "and propagation-stage measures in parallel: quieter pavements, quieter vehicles, barriers [17,51]. The "
 "rough uniformity of the response across activity types suggests scope for broad approaches rather than "
 "regulation differentiated by land use. Whether these results generalise to Tokyo, Hong Kong, Singapore and other dense East "
 "Asian cities remains a hypothesis to be tested."),

("본 연구에는 몇 가지 한계가",
 "Several limitations qualify these conclusions. (1) S-DoT is uncalibrated, coarse (1 dB steps) and "
 "subject to ageing drift, so absolute levels and multi-year series cannot be taken at face value; our "
 "workaround is within-sensor change plus same-day comparison. (2) De-facto population records presence, "
 "not the activities that generate noise (traffic above all), and dong-by-date shocks such as "
 "construction or events, which move both presence and noise, survive date fixed effects; this is why we "
 "read the estimate as a conditional association. Linking stop-level transit boardings to dong [35,37] "
 "would sharpen interpretation. (3) The exact effective dates in the stringency coding warrant further "
 "checking against KDCA source records. (4) Weather comes from reanalysis (ERA5) and could be replaced "
 "with official station observations. (5) The effect of the citywide common mobility component is "
 "absorbed by date fixed effects and is not identified here, with no bound on its size or direction; "
 "recovering absolute levels by fusion with calibrated reference networks is a task for follow-up work. "
 "(6) The activity classification approximates land use by the day-night population ratio (with "
 "circularity limited by using only post-lifting levels); official zoning data would refine it. (7) The "
 "daily quality filter (12 h or more) sets no within-window minimum, so some Lnight values rest on few "
 "hours; window-specific thresholds remain to be examined. (8) The analysis window opens with S-DoT "
 "provision (2020-04) and misses the strongest mobility shock, the first wave and the first intensive "
 "distancing of March 2020; with the extremes of the dose distribution truncated, power to detect "
 "saturation or thresholds is limited, and the linear approximation is claimed only within the observed "
 "range. (9) De-facto population does not separate indoor from outdoor presence. In phases when people "
 "stayed indoors, equal presence produced less street-level source activity, overstating exposure and "
 "attenuating the coefficient; noise-source changes that ran against the fall in mobility, such as the "
 "pandemic-era growth in delivery motorcycles, may likewise be folded into the observed slope. (10) Installation context (setback from roads, mounting "
 "height, facades) is unobserved, so heterogeneous responsiveness across sites is averaged into β. (11) "
 "The analysis covers a single city, and extensions to the fine spectral and temporal structure of noise "
 "and to human soundscape perception [11,16,33,34] remain open."),

("우리는 한국의 단계적 거리두기를 자연실험",
 "Using Korea's graded distancing as a natural-experiment setting, we estimated, to our knowledge for "
 "the first time, a graded dose-response of urban noise to measured mobility on a city-scale IoT "
 "network. The association is statistically solid but small: a 30% fall in mobility corresponds to about "
 "0.23 dB of daytime noise, and a 50% fall to about 0.45 dB. Detectable responses appear only for the "
 "daytime outcome (the day-night difference itself is unresolved), keep the same sign through all four "
 "seasons (spring and autumn surviving multiplicity adjustment), and are estimated under the conservative "
 "same-day cross-neighbourhood comparison; in that strict sense they are conditional associations. The "
 "signal vanishes in long-run spatial cross-sections, and the multi-year absolute levels of the low-cost "
 "network are corrupted by drift. For policy, at least at the level of relative neighbourhood change, "
 "mobility-reducing urban policies (car-free streets, fifteen-minute cities) are unlikely on their own "
 "to lower urban noise by much, and noise targets will require source- and propagation-stage "
 "interventions alongside them. The results also temper the popular reading of the lockdown literature: "
 "citywide quieting of several decibels, generalised from a few hotspots, may overstate what a graded "
 "mobility dose-response supports, and correcting for this helps recalibrate expectations about the "
 "costs and benefits of noise policy."),

("방법론적으로, 본 연구는",
 "Methodologically, the study sets out a procedure for using uncalibrated high-density IoT noise "
 "networks with confidence: within-sensor relative change, two-way fixed effects, outlier-robust "
 "statistics and cross-checks against official networks, with the signal confirmed by dong-level "
 "permutation and weighting sensitivity checks. Because absolute levels mislead under multi-year drift, "
 "smart-city noise monitoring should adopt within-sensor change designs from the outset. Future work "
 "should link more direct traffic exposures, such as stop-level boardings, to sharpen the estimate, fuse "
 "the network with calibrated reference stations to restore absolute levels, and repeat the design on "
 "other low-cost networks (SONYC, wireless acoustic sensor networks, NoiseCapture) and in other dense "
 "East Asian cities (Tokyo, Hong Kong, Singapore) to test its generality."),

("모든 원시데이터는",
 "All raw data are public: S-DoT noise (Seoul Open Data Plaza OA-15969), de-facto population (OA-14991), "
 "subway ridership (OA-12921), the distancing implementation history (data.go.kr 15106451), weather "
 "(Open-Meteo ERA5), administrative boundaries (vuski/admdongkor), the four roadside LAeq stations "
 "(OA-15473) and the calibrated environmental-noise networks (noiseinfo.or.kr). The subway data are "
 "distributed under KOGL Type 3 (attribution, no modification)."),
]
