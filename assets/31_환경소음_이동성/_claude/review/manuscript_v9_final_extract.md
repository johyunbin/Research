# v8 최종본 추출 (Manuscript_20260726_231342.docx)
- 문단 224 · 표 7 · 임베드 이미지 9
- 이미지: #1:image1.png(1468KB), #2:image2.png(334KB), #3:image3.png(394KB), #4:image5.png(991KB), #5:image6.png(565KB), #6:image7.png(430KB), #7:image8.png(739KB), #8:image10.png(278KB), #9:image11.png(241KB)
- 잔존 § 참조: 0

## The dose-response of urban noise to human mobility measured with a city-scale IoT sensor network in Seoul
Hyun In Jo*
Department of Architectural Engineering, Hanyang University, Seoul 04763, Korea
Corresponding author:
Hyun In Jo (best2012@hanyang.ac.kr)
Architectural Acoustics Lab (Room 605–1)
Department of Architectural Engineering
Hanyang University
222 Wangsimni-ro, Seongdong-gu
Seoul 04763, Korea
Phone: +82 2 2220 1795
Fax: +82 2 2220 4794
## Highlights
• Graded distancing, measured mobility and 1,123 IoT sensors give a dose–response
• A 30% fall in neighbourhood mobility is tied to only ~0.23 dB less daytime noise
• Only daytime noise shows a detectable association, consistent in sign across seasons
• The 2.3 dB multi-year S-DoT decline is consistent with drift, not quieter streets
• Within-sensor, same-day designs are essential for low-cost noise networks
## ABSTRACT
How strongly the urban acoustic environment responds to human mobility is a first-order question for mobility, land-use and noise policy, yet mobility rarely varies exogenously. We exploit Korea's graded COVID-19 social distancing, which repeatedly tightened and relaxed activity over 2020-2023, together with a city-scale low-cost IoT noise network, as a graded natural-experiment setting. We assembled 1,248,794 sensor-days from 1,123 Smart Seoul Data of Things (S-DoT) sensors across 421 neighbourhoods (dong) and matched each sensor-day to neighbourhood daytime de-facto population. Because low-cost sensors carry calibration offsets and multi-year drift, we estimate the dose-response only from within-sensor variation and from cross-neighbourhood variation within each date (two-way fixed effects), and read it as a within-date conditional association. Daytime noise rises and falls with daytime mobility (β = +0.65 dB per log-unit; a 30% mobility reduction corresponds to about 0.23 dB), is robust to a dong-level permutation check (p = 0.003), and is positive across seasons; only daytime noise is individually significant, and the high-low neighbourhood gap tracks regime tightening and closes after full lifting. No robust gradient survives in the long-run spatial cross-section, and a calibrated network shows that S-DoT reads about 12 dB below nearby calibrated LAeq stations, its multi-year decline consistent with sensor drift. The response is statistically robust but modest: at the neighbourhood scale, mobility-demand management alone is unlikely to deliver large noise reductions and should be paired with source- and propagation-stage measures. The design offers a transferable, drift-aware template for smart-city noise monitoring.
Keywords: urban noise; environmental acoustics; human mobility; dose-response; IoT sensor network; smart-city monitoring; de-facto population; Seoul; natural experiment
## 1. Introduction
## 1.1. Environmental noise and urban sustainability
도시가 사람들의 이동을 줄이면 도시는 얼마나 조용해지는가? 재택근무, 교통수요 관리, 토지이용 개편처럼 지속가능한 도시 정책의 상당수가 이동량을 겨냥하지만, 이 단순한 질문의 기울기는 정량적으로 알려져 있지 않다. 환경소음은 대기오염에 이어 도시 환경이 인간 건강에 부과하는 두 번째로 큰 부담으로 평가된다. 세계보건기구(WHO)는 환경소음을 유럽 도시의 주요 질병부담 원인으로 규정하며, 서유럽에서만 매년 100만 건강수명연수(healthy life-years) 이상이 교통소음으로 손실되는 것으로 추정한다 [1,2]. 소음의 건강영향은 단순한 불쾌를 넘어선다. 야간소음은 수면을 단편화하고 [3], 만성적 노출은 자율신경·내분비 스트레스 경로를 통해 고혈압·허혈성 심질환 등 심혈관계 질환 위험을 높이며 [4,5], 주간에는 소음 성가심(annoyance)과 인지수행 저하를 유발한다 [6,7]. 이러한 영향은 상당한 사회·경제적 비용으로 이어진다 [8].
따라서 소음은 도시의 지속가능성과 거주적합성(livability)을 좌우하는 핵심 물리적 스트레서이며, 그 부담이 사회경제적 약자에게 불균등하게 분포한다는 점에서 환경정의(environmental justice)의 문제이기도 하다. 소음을 단순 데시벨이 아니라 맥락 속 인간 경험으로 다루는 사운드스케이프(soundscape) 패러다임 [9,10,11,12,13,14]은 이 인식을 확장해, 도시 공공공간의 음향쾌적성 [15]과 토지이용·도시형태에 따른 음환경의 시공간 변이 [16]를 함께 다룬다(본 연구의 범위는 물리적 수준 변화에 한정된다). 그러나 소음을 '관리'하려면 한 가지 근본 질문에 답해야 한다. 지속가능성 정책이 실제로 쥐고 있는 레버는 차량 그 자체가 아니라 통행·활동의 총량과 배치다. 그렇다면 도시소음은 그 레버, 곧 인간의 활동과 이동량에 얼마나 민감하게 반응하는가? 이 반응성(responsiveness)은 모든 이동·토지이용 기반 소음정책의 효과를 가늠하는 출발점이지만, 정량적으로 규명된 바 많지 않다.
## 1.2. Mobility as a driver of urban noise
이 반응성을 기대할 물리적 근거는 분명하다. 도시소음의 지배적 원천은 도로교통이며, 소음 수준은 교통량과 함께 증가하기 때문이다 [17]. 그러나 정책이 실제로 다루는 대상은 교통량 하나가 아니라 사람들의 활동과 이동 전체이고, 이 '도시 전체의 인간 활동·이동량'과 소음의 관계가 인과적 기울기로 추정된 적은 사실상 없다. 그 이유는 식별(identification)의 어려움에 있다. 정상적인 도시에서 이동량은 외생적으로 변하지 않으며, 토지이용·시간대·요일·날씨 등 소음과 이동량을 동시에 좌우하는 공통 요인에 묶여 있다. 따라서 관측 자료의 단순한 이동량-소음 상관은 이 공통 요인들에 의해 교란되어, 이동량이 '원인'으로서 소음을 얼마나 변화시키는지를 말해 주지 못한다.
기존의 교통소음 예측 모델, 예컨대 CNOSSOS-EU 같은 표준식이나 토지이용회귀(land-use regression) 모델 [17,18,29]은 교통량·도로폭·토지이용으로부터 소음을 추정하지만 기본적으로 횡단적·정적(static)이다. 즉 '교통이 많은 곳이 시끄럽다'는 공간 패턴은 잘 재현하나, '이동량이 줄면 소음이 얼마나 줄어드는가'라는 동적 용량-반응(dynamic dose-response)은 답하지 못한다. 이 정적 모델의 공백을 메우려면 이동량의 외생적·점진적(graded) 변동과, 그에 대응하는 고밀도 소음 관측이 동시에 필요하다. 이는 정상적인 도시에서는 좀처럼 주어지지 않는 조건이다.
## 1.3. COVID-19 as a natural experiment
코로나-19 대유행은 이 조건을 이례적으로 충족시킨 자연실험(natural experiment)이었다. 봉쇄·거리두기는 인간 이동량을 외생적으로 급감시켰고, 다수 연구가 봉쇄기 도시소음의 수 dB 감소를 계측으로 보고했다(마드리드는 4-6 dB 감소 [19], 더블린 [20], 런던 [21], 스톡홀름 [22], 몬트리올 [23], 그라나다 [24], 칸푸르 [28]). 지각 조사에서도 아르헨티나 [25]·리마 [26]·이탈리아 [27]에서 봉쇄기 음환경의 변화가 보고되었다. 같은 시기 대기질에서도 이산화질소·이산화탄소가 급감해 [30,31,32] 봉쇄가 도시 환경 전반에 미친 충격을 보였다.
그러나 이 빠르게 축적된 문헌은 네 가지 구조적 한계를 공유하며, 이는 본 연구의 동기와 직결된다. 첫째, 대부분 '봉쇄 vs 비봉쇄'의 이항적(binary) 비교로, 규제 강도가 단계적으로 달라지는 graded 용량-반응 곡선을 추정할 수 없다. 그 결과 소음이 이동량 변화에 비례적으로 반응하는지, 그 기울기가 얼마인지 알 수 없다. 둘째, 대부분의 연구가 소음 감소를 '교통이 줄어서'라고 서술적으로 귀인할 뿐, 이동·활동을 실제로 측정해 소음과 정량적으로 결합하지 않는다. 차량 교통량 계측과 결합한 예외적 연구들이 있으나(포르투의 소음센서 4점×유도루프 [52], 로마의 FCD 기반 교통·배출 시뮬레이션 [53], 보훔의 단일 지점 장기측정 [54]), 소수 지점의 차량 교통류에 한정되며, 도시 전역의 사람 활동량을 동 단위로 소음과 결합한 예는 우리가 아는 한 없다. 셋째, 측정점이 소수(대개 4-70점)이고 서구 도시에 편중되어, 고밀도·동아시아 도시의 미시적 공간 이질성과 토지이용별 차등을 포착하지 못한다. 넷째, 거의 모든 연구가 센서의 절대 음압을 그대로 비교해, 측정망 자체의 검교정 편의와 다년 드리프트가 결과에 섞여 들 여지를 남긴다. 대표적 선행연구를 이 네 축으로 비교하면 어느 연구도 두 축 이상을 함께 충족하지 못한다(Table 1). 요약하면 필요한 것은 단계적으로 변하는 노출(graded), 실측된 이동량(measured), 수백 점 이상의 고밀도 관측(high-density), 그리고 센서 편의를 설계로 제거하는 식별 전략(drift-aware)이며, 동아시아 고밀도 도시에서 이 조건들을 동시에 충족하는 증거는 아직 없다. 이것이 본 연구가 필요한 이유다.
Table 1. Design features of representative COVID-19 urban-noise studies and the present study.

[[TABLE 1]]
  | Study || City (measurement points) || Exposure contrast || Mobility measured? || Sensor-bias handling
  | Asensio et al. [19] || Madrid (municipal noise network) || Binary: lockdown vs. pre-lockdown || No || Absolute levels
  | Basu et al. [20] || Dublin (12 fixed stations) || Binary: lockdown phases vs. baseline || No || Absolute levels
  | Aletta et al. [21] || London (short-term site measurements) || Binary: lockdown vs. pre-lockdown || No || Absolute levels
  | Rumpler et al. [22] || Stockholm (central station) || Binary: recommendation period vs. before || No || Absolute levels
  | Manzano et al. [24] || Granada (urban measurement points) || Binary: lockdown vs. reference || No || Absolute levels
  | Mishra et al. [28] || Kanpur (campaign sites) || Binary: lockdown vs. before || No || Absolute levels
  | Traffic-count studies [52,53,54] || Porto (4 sensors + loop detectors); Rome (FCD simulation); Bochum (1 long-term site) || Lockdown era, vehicle traffic only || Partly: vehicle counts, no de-facto population || Absolute levels / emission model
  | This study || Seoul (1,123 IoT sensors, 421 neighbourhoods) || Graded: repeated tightening and relaxation, 2020-2023 || Yes: daily neighbourhood de-facto population || Within-sensor change; within-date two-way FE; calibrated-network cross-comparison

Note: measurement-point descriptions are indicative; see the cited papers for details.
## 1.4. Measured mobility and IoT noise sensing
최근 두 가지 데이터 혁신이 이 요건들 가운데 앞의 세 가지를 비로소 충족 가능하게 했다. 첫째는 모바일 통신·신호 기반의 생활인구(de-facto population)다. 이는 거주지가 아니라 특정 시점에 사람들이 '실제로 어디에 있는가'를 도시 미세공간·고빈도로 측정한다 [35,36,37,38,39]. 둘째는 저가 IoT 음향 센서망이다. SONYC [40], 저가 모니터링 기기 [41], 무선음향센서망(WASN) [42,43], 참여형 NoiseCapture [44], 스마트시티 센싱 [45,46] 등은 도시 음환경을 수백~수천 점에서 상시 관측할 수 있게 했다. 한국의 단계적 거리두기, 서울 전역의 S-DoT 소음센서망(약 1,100점), 행정동 단위 생활인구는 이 요건들을 한 도시 안에서 동시에 제공하는 드문 조합이다.
그러나 저가 IoT 소음망에는 기존 COVID-소음 문헌이 대체로 간과해 온 구조적 문제가 있다. 바로 센서마다 다른 검교정 오프셋과, 노후·환경요인으로 측정값이 서서히 변하는 드리프트(표류)다 — 같은 문제가 저가 대기질 센서망에서 이미 잘 알려져 있다 [47]. 저가 센서는 애초에 절대 음압을 정밀 측정하도록 검교정되지 않기 때문이다. 이를 무시하고 절대레벨을 시계열·공간으로 단순 비교하면 거짓 패턴 또는 거짓 결론에 이르며, 본 연구가 그 실례를 직접 보인다. 따라서 남는 것은 넷째 요건, 곧 이 교란을 결과 해석 단계가 아니라 식별 설계 단계에서 제거하는 일이다.
## 1.5. The present study
본 연구는 한국의 단계적 사회적 거리두기를 graded 자연실험적 맥락으로, 서울의 S-DoT IoT 소음센서를 결과변수로, 행정동 생활인구를 측정 이동량 dose로 결합해 도시소음의 이동량 용량-반응을 추정한다. 구체적으로, 우리는 절대레벨이 아닌 센서-내 상대변화와 같은 날짜 안에서 동(洞) 간 변이만으로 기울기를 추정하는 양방향 고정효과(two-way fixed effects) 전략 [48,49]과 자연실험 틀 [50]을 채택하고, 저가 IoT의 드리프트·해상도 한계를 공식 측정망으로 교차검증한다. 이로써 서울 사운드스케이프의 시공간 변이 연구 [16]를 준실험 맥락의 정량적 이동량-소음 추정으로 확장하며, 본 연구의 위치는 Table 1의 마지막 행에 정리했다. 본 연구는 아래 세 가지 질문에 답하며, 그 과정 자체가 저가 IoT 소음망을 어떤 설계로 써야 하는가라는 방법론적 질문에 대한 답을 겸한다.
RQ1. 측정 이동량 감소는 도시소음을 얼마나 낮추는가(용량-반응 기울기)?
RQ2. 그 효과는 토지이용·주야·요일·계절에 따라 어떻게 다른가?
RQ3. 거리두기 완화·해제 국면에서 소음은 어떻게 반등하며, 이 관계가 동별 장기 평균(공간 단면)에서도 관찰되는가?
본 연구의 기여는 네 가지다. 첫째, 가정이 아니라 도시 전역에서 동 단위로 실측된(measured) 활동량에 대한 도시소음의 단계적(graded) 용량-반응을 우리가 아는 한 처음으로 정량 추정한다. 둘째, 약 1,100개 지점의 고밀도 IoT 센서망과 행정동 생활인구를 결합해 도시 미시공간 해상도에서 그 효과를 식별한다. 셋째, 서구 도시에 편중되어 온 기존 문헌을 서울이라는 고밀도 동아시아 메가시티로 확장한다. 넷째, 저가 IoT 소음망의 검교정·드리프트 한계를 정량적으로 드러내고, 절대 수준 대신 센서 내 상대변화와 '같은 날 동 사이 비교'에 기반해야 한다는 식별 설계 원칙을 제시한다.
## 2. Materials and methods
## 2.1. Study area, period and study design
연구지역은 서울특별시 25개 자치구 전역이다. 서울은 약 960만 명이 605 km²에 거주하는 세계 최고 밀도의 메가시티 중 하나로, 고밀도 혼합토지이용과 조밀한 행정동(行政洞, administrative neighbourhood) 체계(421개)를 갖추어 도시 안의 미세한 공간 차이를 포착하기에 이상적이다(Fig. 1). 분석 기간은 S-DoT 소음 자료가 처음 제공된 2020-04-01부터 2023-12-31까지이며, 거리두기가 전면 해제된 2022-04-18 이후의 정상기를 충분히 포함해 소음이 평상시 수준으로 되돌아가는 과정(rebound)까지 추적한다. 연구 설계는 세 가지 축 위에 서 있다. 첫째, 한국의 거리두기는 규제를 단계적으로 조였다 풀기를 반복하며 사람들의 이동량을 외부 요인에 의해(외생적으로) 점진적으로 변화시켰는데, 우리는 이 변동을 일종의 통제된 실험처럼 활용한다(자연실험적 맥락). 둘째, 뒤에서 설명하듯 저가 센서는 저마다 측정값에 일정한 치우침이 있어 절대 수준을 그대로 쓸 수 없으므로, 결과변수로는 각 센서가 자기 평소 대비 얼마나 달라졌는지(센서 내 상대변화)만 사용한다. 셋째, 효과는 '같은 날, 도시 안에서 동마다 이동량이 얼마나 달랐는가'라는 차이로부터 식별한다.
[[IMAGE 1: image1.png 1468KB]]
Fig. 1. Study area and the S-DoT sensing network. (a) 1,165 geolocated S-DoT noise sensors (1,123 enter the analysis panel after quality control) across the 421 administrative neighbourhoods (dong) of Seoul, coloured by neighbourhood activity type (terciles of the daytime/night-time de-facto population ratio: commercial, mixed, residential; Supplementary Fig. S3). (b) Baseline daytime noise level (post-lifting mean Lday, July 2022 to December 2023) by neighbourhood.
## 2.2. Outcome: urban noise from the S-DoT network
결과변수는 서울시가 도시 전역에 상시 운영하는 IoT 도시데이터 센서망(Smart Seoul Data of Things, S-DoT)의 소음 자료다. 약 1,100개 지점에서 시간별 광대역 음압(dB)을 측정하며, 모든 자료는 별도 인증 없이 공개 다운로드된다(Table 2). 우리는 시간별 값을 하루 단위의 에너지 등가소음도로 합쳐, 전일 Leq,24h, 주간 Lday(06-21시), 야간 Lnight(22-05시)을 각각 10·log10(평균(10^(L/10)))으로 계산하고, 하루 최소 12시간 이상 관측되고 값이 물리적으로 타당한 범위(20-95 dB)에 드는 자료만 남겼다. 2023년에는 자료 구조가 바뀌어 소음이 최대·평균·최소 세 값으로 나뉘었는데, 이 가운데 평균값을 사용했다. 여기서 한 가지를 유의해야 한다. S-DoT는 도시 전반을 감지하는 범용 IoT 노드로, 정밀 검교정을 거치지 않은 광대역 데시벨을 제공하며(소음 분야의 표준 측정량인 A-가중 등가소음도 LAeq라는 라벨이 붙어 있지 않고, 마이크 사양·주파수/시간 가중·2분 원시값에서 시간별 값으로의 집계 규칙도 공개 명세에 명시되어 있지 않다), 서로 다른 센서의 절대값을 맞비교하는 데에는 쓸 수 없다. 주간·야간 구간은 달력일 기준이며, 야간 Lnight는 같은 달력일의 00-05시와 22-23시 관측을 묶은 것이다. 따라서 우리는 절대 수준이 아니라 각 센서의 시간에 따른 변화만 분석에 쓰고, 그 신뢰성은 공식 LAeq 측정망과의 대조 진단으로 점검한다(Fig. 8, Supplementary Fig. S2).
Table 2. Data sources and processing.

[[TABLE 2]]
  | Component || Dataset (provider) || Resolution / processing || Period || Use in analysis
  | Outcome: / urban noise || S-DoT city-wide IoT sensing network (Seoul Open Data Plaza, OA-15969) || ~1,100 fixed sensors; hourly broadband SPL (dB) aggregated to daily Leq,24h, Lday (06-21 h), Lnight (22-05 h); QC: ≥12 valid h/day, 20-95 dB; 2023 schema change reconciled ('mean' field) || 2020-04 to 2023-12 || Outcome; within-sensor change only (absolute levels never compared across sensors)
  | Exposure: / mobility || De-facto ('living') population per administrative dong (Seoul Open Data Plaza, OA-14991) || Hourly counts per 424 dong, averaged to daily daytime (06-21 h) and night-time (22-05 h) values; 619,464 dong-days, no missing || 2020-01 to 2023-12 || Dose lp = log(daily / dong's post-lifting baseline mean); within-dong relative change
  | Policy / covariate || Social-distancing implementation history (KDCA) || Daily regime re-coded to continuous stringency 0-7 from business curfew hour and gathering cap || 2020-01 to 2022-04 || Auxiliary descriptor (Table 3); absorbed by date FE in M2
  | Weather / covariate || Open-Meteo ERA5 reanalysis at Seoul city centre (37.57° N, 126.97° E) || Daily mean/max/min temperature, precipitation, maximum wind speed || 2020-2023 || M1 covariates; absorbed by date FE in M2
  | Calendar || Day of week, weekend, Korean public holidays (incl. substitutes), season || Daily indicators || 2020-2023 || Weekend/holiday indicators as M1 covariates; absorbed by date FE in M2
  | Comparison / networks || Calibrated official stations: national environmental-noise network (noiseinfo.or.kr; automatic daily road stations, manual quarterly general/road stations) and four roadside LAeq stations (OA-15473) || Station-level standard LAeq || 2020-2024 || Drift diagnosis and absolute-offset comparison (Fig. 8, Supplementary Fig. S2)

## 2.3. Exposure: human mobility from de-facto population
이동량 노출은 서울 '생활인구(living / de-facto population)'로 측정했다. 이는 거주 인구가 아니라 모바일 통신 신호로 추정한, 특정 시점에 각 행정동에 실제로 체류하는 인구수다(Table 2). 시간별 체류수(stock)를 일 평균(및 주간 06-21 / 야간 22-05 평균)으로 집약해 424개 동 × 1,461일 = 619,464 동-일을 얻었다(결측 0; 주간 06-21시 / 야간 22-05시 평균). 이 가운데 S-DoT 센서가 배정된 421개 동이 분석 대상이 된다. dose 구성에서 한 가지 통찰이 중요하다. 도시 전체 생활인구 총량은 거의 보존된다(사람들이 외출을 줄여도 자택 체류로 계수되기 때문). 즉 거리두기의 신호는 총량이 아니라 공간 재분포에 있으며, 최강 제한 국면에는 상업·업무 동의 주간 체류가 동에 따라 최대 약 35%까지 비워지고 주거 동은 최대 37%까지 늘어난다(거리두기 기간 전체 평균으로는 동별 감소가 대체로 0-13% 범위다; Fig. 7a). 따라서 우리는 dose를 동의 자기기준 대비 상대변화로 정의한다. 곧 lp = log(동의 일별 생활인구 / 동의 post-lift 정상기 평균)이다. 이렇게 하면 dose는 각 동이 자기 정상에서 얼마나 벗어났는가를 나타내며, 동·센서별 절대수준 차이에 영향받지 않는다. 두 가지를 명확히 해 둔다. 첫째, 생활인구는 이동 흐름이 아니라 체류 인구(stock)이므로, 본 논문에서 '이동량(mobility)'은 활동 존재량(activity presence)의 변화를 가리키는 조작적 용어다. 둘째, 기준기간은 전면 해제 직후의 과도기(2022-04~06)를 제외하고 행동이 안정된 2022-07~2023-12로 정했다. 같은 날짜 안에서 동 간 dose의 표준편차는 평균 0.115 log-unit로, 날짜 고정효과 아래에서도 식별에 쓸 변동이 충분하다.
## 2.4. Linking sensors to neighbourhoods
센서 위치정보에는 도로명주소만 있고 행정동 이름이 없어, 각 센서를 그 좌표가 어느 행정동 경계 안에 들어가는지로 동에 배정했다(점-다각형 포함 판정). 행정동 경계는 공개 GeoJSON(vuski/admdongkor, 2022-01 버전)을 사용했다. 한편 생활인구 자료와 경계 자료의 행정동 코드 체계가 일부 지역(강북·강동의 행정구역 개편 동)에서 서로 다르게 매겨져 있었는데, 동 이름을 기준으로 짝지어 바로잡았다. 그 결과 1,170개 센서 중 1,165개(99.6%)가 동에 배정되었고, 주소상의 자치구와 배정된 동의 자치구가 일치하는 비율도 99.6%로 매우 높았다. 이후 소음 QC(하루 12시간 이상·20–95 dB)와 일별 이동량 결합을 거친 최종 분석패널은 1,123개 센서·421개 동(1,248,794 sensor-days)이며, 주력 양방향 고정효과 모형은 주간 이동량이 결측이 아닌 1,122개 센서·420개 동(1,247,546 sensor-days)을 사용한다.
## 2.5. Covariates: social distancing, weather and calendar
거리두기 '단계'(1·2·2.5·4단계 등)를 노출변수로 그대로 쓰는 것은 적절치 않다. 단계 체계가 두 차례 통째로 재정의되었고(2020-11, 2021-07), 단계 사이의 간격이 일정하지 않으며, 같은 단계라도 시기에 따라 실제 이동량이 달랐기 때문이다. 대신 우리는 매일의 규제를 두 가지 구체적 수치, 곧 식당·카페 영업종료시각과 사적모임 허용 인원으로부터 0(규제 없음)에서 7(최강)까지 이어지는 연속적 강도(stringency) 지수로 다시 코딩했다(Table 3). 다만 이 지수는 보조 변수일 뿐이며, 본 연구의 주된 노출변수는 어디까지나 실측 이동량(생활인구)이다. 기상은 Open-Meteo ERA5 일자료(기온·강수·풍속)를 쓰고, 달력 통제로는 주말·한국 공휴일 지표를 M1에 투입한다(요일·계절 정보도 자료에 있으나, M2에서는 날짜 고정효과가 모든 달력 효과를 흡수하므로 별도 투입이 불필요하다).
Table 3. Seoul capital-area social-distancing timeline and its quasi-continuous re-coding (selected regime changes).

[[TABLE 3]]
  | Effective from || Regime || Business curfew (h) || Gathering cap || Stringency (0-7)
  | 2020-01 || Pre-COVID (normal) || 24 || none || 0
  | 2020-03-22 || 1st intensive distancing || 21 || ≤10 || 4
  | 2020-05-06 || Daily-life distancing || 24 || none || 0
  | 2020-08-30 || Capital Level 2.5 (first 21 h cap) || 21 || ≤50 || 5
  | 2020-12-23 || Level 2.5 + 5-person ban || 21 || ≤4 || 6
  | 2021-02-15 || Level 2 + 5-person ban || 22 || ≤4 || 5
  | 2021-07-12 || Capital Level 4 || 22 || ≤4 || 5
  | 2021-11-01 || With-COVID recovery || 24 || ≤10 || 1
  | 2021-12-18 || Special measures || 21 || ≤4 || 6
  | 2022-03 || Gradual easing || 23 || ≤8 || 3
  | 2022-04-18 || Full lifting || 24 || none || 0

Note: the stringency index re-codes the two enforceable components (business curfew hour and private-gathering cap) onto a 0-7 scale. The official tier system itself was redefined in 2020-11 and 2021-07, so tier labels are not comparable over time and are never used as the dose; the measured mobility of each dong is the exposure throughout.
## 2.6. Statistical analysis and identification
저가 S-DoT 센서는 정밀 검교정을 거치지 않아, 똑같은 소리를 들려주어도 센서마다 일정량 높거나 낮게 기록하는 고유의 치우침(상수 오프셋)을 갖는다. 이 치우침의 크기는 센서마다 다르고 알 수 없으므로, 서로 다른 두 센서의 절대 데시벨을 맞비교하는 것은 의미가 없다. 그러나 한 센서가 시간에 따라 얼마나 달라졌는지(예컨대 오늘이 그 센서의 평소보다 얼마나 큰지)만 보면, 그 센서에 늘 똑같이 들어 있는 치우침은 빼는 과정에서 저절로 지워진다. 그래서 우리의 모든 분석은 절대 수준이 아니라 각 센서의 자기 대비 변화(센서 내 상대변화)에 기반한다. 통계적으로 이는 고정효과(fixed-effects) 모형으로 구현된다. 각 센서의 전체 평균을 빼 줌으로써 그 센서 고유의 치우침을 제거하는 것이다(125만 건의 자료에 센서별 더미변수를 일일이 넣는 대신, 집단 평균을 차감하는 수학적으로 동등한 방법을 썼다). 또한 같은 군집에서 나온 관측치들은 서로 닮아 있을 수 있으므로 표준오차를 군집에 강건하게(cluster-robust) 보정했다. 센서 고정효과 모형은 센서 기준, 양방향 고정효과 모형은 동 기준으로 군집화했다 [48]. 전체 식별 논리는 Fig. 2에 요약했다.
주된 분석은 두 단계의 고정효과 모형이다. (M1) 센서 고정효과 모형은 각 센서의 평균을 제거한 뒤, 그 센서의 소음 변화를 동 이동량과 날씨(기온·기온²·강수·풍속·강수 여부)·주말·공휴일로 설명한다. 이렇게 하면 눈에 보이는 시간적 교란을 통제하면서도 시간에 따른 변이는 남겨 두어 이동량의 효과를 추정할 수 있다. (M2)는 본 연구의 핵심 모형으로, 여기에 날짜 고정효과를 더한다. 날짜 고정효과는 특정 날짜에 도시 전체가 공통으로 겪은 모든 것을 한꺼번에 흡수한다. 곧 그날의 날씨, 요일·공휴일, 해가 갈수록 변해 온 도시 전반의 소음 추세, 센서망 공통의 표류, 전국적 거리두기가 모두 여기에 포함된다. 그러면 남는 정보는 오직 '같은 날, 같은 도시 안에서 동(洞)마다 이동량이 얼마나 달랐는가'뿐이다. 따라서 M2는 다음 질문에 답한다. 같은 날, 자기 평소보다 더 비워진 동에 있는 센서가 덜 비워진 동의 센서보다 더 조용해졌는가? 이 설계의 한 가지 귀결로, 도시 전체에 하루 하나의 값으로만 변하는 변수(예: 그날의 지하철 총승객 수나 거리두기 강도)는 날짜 고정효과에 완전히 흡수되어 따로 효과를 추정할 수 없다. 동마다 값이 다른 실측 이동량만이 유효한 노출변수다. 형식적으로 M2는 Y_sdt = α_s + λ_t + β·x_dt + ε_sdt이며, 식별 가정은 같은 날짜 안에서 잔여 소음 충격이 동별 이동량과 무관하다는 것이다. 국지 공사·행사·사업장 운영처럼 동×날짜 수준에서 이동량과 소음을 함께 움직이는 충격은 이 가정을 위협할 수 있으므로, 우리는 β를 인과효과가 아니라 '같은 날짜 안의 조건부 연관(within-date conditional association)'으로 읽는 보수적 독법을 전 해석에서 유지한다. 두 종류의 고정효과는 센서 평균과 날짜 평균을 번갈아 빼는 계산을 수렴할 때까지 반복해 처리했다.
이 두 모형 위에 보조 분석을 더했다. (M3) 기능 분할: M2를 결과변수(주간·야간·주야 차이), 토지이용(동의 주간 대비 야간 인구비를 3분위로 나눈 상업·혼합·주거; Supplementary Fig. S3 — 이 분류는 노출의 일별 변동이 아니라 해제 후 정상기의 수준 정보만 쓰므로 노출과의 순환성이 제한적이며, 실질은 활동 프로파일 분류이나 이하 편의상 '토지이용'으로 부른다), 평일/주말, 계절별로 따로 추정해 효과가 어디서 나타나는지 본다. (M4) 고영향-저영향 궤적 비교: 이동량이 크게 줄어든 동(주로 상업지구)과 거의 줄지 않은 동(주로 주거지구)을 나눠, 두 그룹의 소음 변화를 매주 비교한다. 두 그룹은 같은 도시·같은 시기를 공유하므로 날씨·계절·센서 표류 같은 공통 요인은 차이에서 상당 부분 상쇄된다. 다만 그룹이 제한기의 실현 이동량으로 사후 정의되고 소음·이동량 모두 해제 후 기준창에 정규화되어 있으므로(해제 후 0 수렴이 부분적으로 내장), 이는 정식 event-study가 아니라 기술적(descriptive) 비교로 해석한다. (M5) 공식망 대조 진단: S-DoT의 절대레벨과 다년 시계열의 신뢰성을 검교정된 공식 환경소음 측정망(국가소음정보시스템의 자동·수동 측정망)과 도로교통 4점 상시측정소(표준 LAeq)에 대조해 점검한다. 측정점 인근 S-DoT와의 레벨 차(offset) 및 2020-2023 연추세를 비교해 센서 표류와 2023년 자료구조 변경의 영향을 진단한다. (M6) 강건성·민감도 점검: 추정이 기계적 산물이 아님을 확인하기 위해 ① 같은 날짜 안에서 동별 이동량 dose를 동들 사이에 무작위로 재배치하는 순열 민감도 점검(300회; dose가 동-일 수준에서 배정되므로 셔플도 동 단위로 수행하고 그 동의 모든 센서에 동일 적용), ② 센서가 많은 동이 결과를 지배하지 않는지 확인하는 동-일 동일가중 재추정, ③ 이동량 2차항을 넣은 비선형성 점검을 수행했고, 분할추정(M3)의 다중비교는 Benjamini-Hochberg FDR로 보정했다. 끝으로 동 단위 공간 분석에서는 소수의 극단적인 동에 결과가 휘둘리지 않도록 일반 회귀·상관 대신 극단값에 강한(robust) 방법(Theil-Sen 회귀와 Spearman 순위상관)을 쓰고, 추정이 불안정한 단일 센서 동을 제외해 센서가 2개 이상인 동만 사용했다. 모든 분석은 Python 3.13(pandas·NumPy·SciPy·statsmodels)으로 수행했으며, 양방향 고정효과는 센서·날짜 평균을 번갈아 차감하는 반복 within-변환으로, 극단값에 강한 추정은 SciPy의 Theil-Sen·Spearman으로 계산했다.
[[IMAGE 2: image2.png 334KB]]
Fig. 2. Identification strategy. (a) Conceptual structure: graded social distancing shifts neighbourhood mobility, whose effect on urban noise (solid arrow) must be separated from weather and calendar confounding and from sensor-side artefacts (calibration offset and multi-year drift; dashed arrows). (b) Fixed-effects ladder: sensor fixed effects remove each sensor's calibration offset; date fixed effects remove all date-common factors (weather, calendar, city-wide trend, network-wide drift); the remaining identifying variation is the same-day difference in mobility across neighbourhoods.
## 3. Results
## 3.1. A robust but modest daytime dose-response
분석에 사용한 변수들의 기술통계는 Table 4에 정리했다. 센서 고정효과 모형(M1, Table 5)에서 한 동의 주간 이동량과 그 동의 주간 소음은 같은 방향으로 움직였다(β=+1.130 dB/log-unit, SE 0.292, p<0.001; 표준오차는 dose가 배정되는 동 수준으로 군집화). 통제변수들은 모두 물리적으로 타당했다. 기온은 U자형(춥거나 더울 때 소음이 커짐)이었고, 강수 +0.045 dB/mm, 비 오는 날 +0.10 dB, 주말 −0.83 dB, 공휴일 −1.18 dB였다(모두 p<0.001). 그러나 날짜 고정효과까지 더해 도시 전체가 그날 공통으로 겪은 변화(날씨·추세·표류 등)를 모두 걷어 낸 핵심 모형(M2, Table 5)에서는 효과가 작아졌다(주간 β=+0.648 dB/log-unit, 95% CI 0.17-1.13, p=0.008; 전일 β=+0.628, p=0.038). 풀어 보면, 한 동의 주간 이동량이 평소보다 30% 줄면 그 동의 주간 소음은 약 0.23 dB, 50% 줄면 약 0.45 dB 낮아진다. 즉 도시 전체에 공통된 시간 변화를 엄밀히 제거하고 '같은 날 동 사이의 차이'만으로 추정한 용량-반응은 통계적으로 견고하지만 그 크기는 작다(RQ1).
Table 4. Descriptive statistics of the analysis panel (1,248,794 sensor-days; 1,123 sensors; 421 dongs; 2020-2023).

[[TABLE 4]]
  | Variable || Mean || SD || P5 || P95 || N
  | Lday (dB) || 50.52 || 7.20 || 40.87 || 65.54 || 1,248,793
  | Lnight (dB) || 47.58 || 6.81 || 38.46 || 61.80 || 1,248,623
  | Leq,24h (dB) || 49.93 || 7.07 || 40.51 || 64.76 || 1,248,794
  | Daytime mobility (log relative) || 0.01 || 0.13 || -0.20 || 0.17 || 1,247,547
  | Mean temperature (°C) || 12.53 || 10.45 || -5.80 || 26.50 || 1,248,794
  | Precipitation (mm) || 4.40 || 12.18 || 0.00 || 27.80 || 1,248,794
  | Max wind (m/s) || 4.61 || 1.68 || 2.50 || 7.81 || 1,248,794

Note: mobility is the within-dong log change relative to the dong's post-lifting baseline; N = sensor-days.
Table 5. Mobility dose-response of urban noise: sensor fixed-effects (M1) and two-way fixed-effects (M2) estimates. Coefficient (SE), dong-clustered. *p<0.05, **p<0.01, ***p<0.001.

[[TABLE 5]]
  |  || M1: sensor FE || M1: sensor FE || M1: sensor FE || M2: sensor + date FE || M2: sensor + date FE || M2: sensor + date FE
  | Term || Lday || Lnight || Leq,24h || Lday || Lnight || Leq,24h
  | Mobility (log rel., own period) || +1.130*** / (0.292) || +2.283** / (0.748) || +1.258*** / (0.352) || +0.648** / (0.245) || +0.504 / (0.631) || +0.628* / (0.303)
  | Mean temperature || -0.113*** / (0.004) || -0.098*** / (0.003) || -0.113*** / (0.004) || - || - || -
  | Temperature² || +0.007*** / (0.000) || +0.004*** / (0.000) || +0.006*** / (0.000) || - || - || -
  | Precipitation (mm) || +0.045*** / (0.001) || +0.048*** / (0.001) || +0.047*** / (0.001) || - || - || -
  | Max wind (m/s) || +0.052*** / (0.005) || +0.055*** / (0.005) || +0.052*** / (0.005) || - || - || -
  | Rain day (0/1) || +0.098*** / (0.010) || +0.155*** / (0.009) || +0.116*** / (0.010) || - || - || -
  | Weekend (0/1) || -0.833*** / (0.029) || -0.022 / (0.019) || -0.686*** / (0.027) || - || - || -
  | Holiday (0/1) || -1.176*** / (0.032) || -0.186*** / (0.040) || -0.999*** / (0.033) || - || - || -
  | Date fixed effects || No || No || No || Yes || Yes || Yes
  | Weather + calendar || Yes || Yes || Yes || absorbed || absorbed || absorbed
  | Within-R² || 0.117 || 0.060 || 0.112 || - || - || -
  | Clusters (dong) || 420 || 420 || 420 || 420 || 420 || 420
  | N (sensor-days) || 1,247,546 || 1,247,376 || 1,247,547 || 1,247,546 || 1,247,376 || 1,247,547

Note: all columns use own-period doses (daytime dose for Lday, night-time dose for Lnight, whole-day dose for Leq,24h). M1 controls for weather and weekend/holiday with sensor fixed effects; M2 adds date fixed effects, which absorb all date-common terms. SEs are clustered by dong, the level at which the dose is assigned.
이 효과의 기능적 위치를 분해하면(Fig. 3, Table 6, RQ2) 개별적으로 유의한 반응은 주간에 집중된다. 결과변수별로 주간만 유의하고 야간·주야 gap은 비유의했다(주야 효과 차이의 해석은 다음 절에서 다룬다). 토지이용별로는 상업(+0.34)·혼합(+0.86)·주거(+0.85) 모두 양(+)이었으나 표본 분할로 검정력이 떨어져 개별 유의성은 약했고, 상호작용 검정에서도 토지이용 차이는 유의하지 않았다(효과는 대체로 균일). 평일(+0.49)·주말(+0.64)은 유사했다. 특히 계절별로는 네 계절 모두 같은 양(+) 방향이며(DJF +0.66·MAM +0.69·JJA +0.53·SON +0.76), 겨울·봄·가을이 명목 유의했다. 다만 다중비교(FDR) 보정 후에는 봄(MAM)·가을(SON)이 유의하게 남아, 효과가 특정 계절의 아티팩트가 아니라 연중 같은 방향으로 일관됨을 보인다.
Table 6. Function-segmented dose-response (M2, two-way fixed effects): β (dB per log-unit mobility), 95% CI, nominal and BH-FDR-adjusted p-values. *p<0.05, **p<0.01 (nominal).

[[TABLE 6]]
  | Segment || Group || β || 95% CI || p || FDR p || N
  | By outcome || Daytime Lday || +0.648** || [+0.17, +1.13] || 0.008 || 0.033 || 1,247,546
  |  || Nighttime Lnight || +0.504 || [-0.73, +1.74] || 0.42 || 0.46 || 1,247,376
  |  || Day-night gap || +0.019 || [-0.35, +0.39] || 0.92 || 0.92 || 1,247,375
  | By land use (daytime) || Commercial || +0.335 || [-0.41, +1.08] || 0.38 || 0.45 || 430,160
  |  || Mixed || +0.858 || [-0.09, +1.81] || 0.076 || 0.15 || 425,584
  |  || Residential || +0.847 || [-0.42, +2.12] || 0.19 || 0.25 || 391,802
  | By day type (daytime) || Weekday || +0.486 || [-0.16, +1.13] || 0.14 || 0.21 || 899,272
  |  || Weekend/holiday || +0.641 || [-0.17, +1.45] || 0.12 || 0.21 || 348,274
  | By season (daytime) || Winter (DJF) || +0.658* || [+0.09, +1.22] || 0.022 || 0.066 || 276,254
  |  || Spring (MAM) || +0.695** || [+0.26, +1.12] || 0.002 || 0.018 || 277,158
  |  || Summer (JJA) || +0.527 || [-0.02, +1.08] || 0.060 || 0.14 || 344,417
  |  || Autumn (SON) || +0.758** || [+0.20, +1.31] || 0.007 || 0.033 || 349,717

Note: all segments use the two-way (sensor + date) fixed-effects specification with dong-clustered SEs; FDR p = Benjamini-Hochberg adjusted over the pre-specified exploratory family of 12 tests.
[[IMAGE 3: image3.png 394KB]]
Fig. 3. Mobility dose-response of urban noise (two-way sensor + date fixed effects). (a) Segment-specific estimates (markers = β, bars = 95% CI, dong-clustered SEs) by outcome, land use, day type and season; filled markers survive Benjamini-Hochberg FDR correction (FDR < 0.05), open blue markers are nominally significant (p < 0.05), grey markers are not significant; the right-hand column lists β [95% CI]. (b) Sensor-FE versus two-way FE estimates: the night-time association under sensor FE alone (+2.28) loses significance once date fixed effects absorb common time-varying confounds (+0.50, ns), whereas the daytime estimate survives (+0.65); note that the wide night-time CI overlaps the daytime estimate, so a day-night difference in the coefficients is not itself established.
## 3.2. Daytime versus night-time responses
주간과 야간을 나눠 보면 유의한 반응은 주간 결과변수에서만 나타난다(Fig. 3b). 야간 소음은 센서 고정효과만 둔 모형에서는 야간 이동량에 강하게 반응하는 것처럼 보였으나(β=+2.28, p=0.002), 날짜 고정효과까지 더하자 그 관계가 유의성을 잃었다(β=+0.50, 95% CI −0.73~+1.74, p=0.42). 이는 야간의 겉보기 연관의 상당 부분이 시간에 공통으로 작용한 교란(센서 표류·계절·도시 추세)에 기인했음을 보여 준다. 반면 주간 연관은 날짜 고정효과 아래에서도 살아남는다. 다만 야간 추정치의 넓은 신뢰구간이 주간 계수(+0.65)를 포함하므로, 주간과 야간의 효과 크기가 서로 다르다고 단정할 수는 없다. 실제로 같은 주간 dose에 대한 주야 차이(Lday − Lnight)의 계수는 사실상 0이어서(β=+0.02, p=0.92), 주간 활동 변화에 주간·야간 소음이 비슷한 폭으로 함께 움직였을 가능성과도 부합한다. 요약하면 개별적으로 유의한 신호는 주간에서만 관찰되지만, 주야 효과의 차이 자체는 통계적으로 확립되지 않는다.
## 3.3. Spatiotemporal mobility and functional differentiation
이동량이라는 노출이 시간과 공간에서 어떻게 움직였는지가 본 설계의 핵심이다(Fig. 4). 도시 전체의 생활인구 총량은 거의 변하지 않지만(외출을 줄여도 자택 체류로 계수되므로), 그 인구가 어디에 머무는지는 거리두기 국면마다 크게 재배치된다. 상업·업무 동은 최강 제한 국면에 주간 체류가 동에 따라 최대 약 35%까지 비워지는 반면 주거 동은 오히려 늘어난다(국면 창 평균 기준; 거리두기 기간 전체 평균의 감소 분포는 Fig. 7a). 규제가 가장 강했던 시기(2020-12 5인 이상 모임금지, 2021-07 수도권 4단계)에 도심 상업동이 가장 비워졌고, 위드코로나(2021-11)와 전면해제(2022-04)를 거치며 평소 수준으로 회복했다.
[[IMAGE 4: image5.png 991KB]]
Fig. 4. Spatiotemporal evolution of the mobility dose across four graded distancing phases (phase-window averages of neighbourhood daytime de-facto population relative to the post-lifting baseline): (a) December 2020, Level 2.5 with the 5-person gathering ban; (b) July 2021, capital-area Level 4 (strongest); (c) November 2021, with-COVID relaxation; (d) March 2022, the weeks before full lifting. The city-wide median stays near 1.0 in every phase, but central commercial/business neighbourhoods empty by up to about 35% under the strongest restrictions (blue) while residential neighbourhoods fill (red), and the contrast fades through (c) and (d).
이 기능적 차등은 시간축에서도 뚜렷하다(Fig. 5, RQ3). 상업 동의 주간 이동량은 제한기 내내 기준선 아래로, 주거 동은 기준선 위로 벌어졌다가 2022년 해제 후 수렴한다(수도권 4단계기 중앙값 상업 0.98 vs 주거 1.05). 드리프트를 제거한 상대 소음(동 그룹평균 − 도시평균)에서도 상업 동이 제한기에 도시평균보다 상대적으로 조용한 경향이 관찰된다. 거리두기 전환의 동역학은 고영향 동과 저영향 동의 주별 궤적 차이로 본다(기술적 비교, Fig. 6; 그룹이 제한기의 실현 이동량으로 사후 정의되고 두 변수 모두 해제 후 기준창에 정규화되어 있어 정식 event-study는 아니다). 고영향 동과 저영향 동의 주별 ΔLday 차이는 제한기에 음(−)으로 기울고 정상기에 0으로 수렴했으며(동 클러스터 95% 신뢰구간과 함께 표시), 주별 (이동량 차이 vs 소음 차이) 상관은 Pearson +0.44 · Spearman +0.46으로 일치했다.
[[IMAGE 5: image6.png 565KB]]
Fig. 5. Functional differentiation of mobility and noise over time. (a) Weekly daytime mobility by neighbourhood activity type: commercial neighbourhoods fall below and residential neighbourhoods rise above the post-lifting baseline during restrictions (shaded bands, stringency >= 4), converging after the April 2022 lifting. (b) Daytime noise relative to the city-wide mean (drift-robust, 4-week moving average) by activity type.
[[IMAGE 6: image7.png 430KB]]
Fig. 6. High- versus low-mobility-loss neighbourhood trajectories (descriptive). Weekly difference between the two groups in (a) within-sensor ΔLday and (b) daytime mobility; shading shows pointwise 95% confidence bands from dong-clustered standard errors (groups treated as independent), and grey bands mark strong-restriction periods. Date-common drift, season and city-wide trends cancel in the difference. Because groups are defined from realised mobility during the Level-4 period and both series are normalised to the post-lifting window, near-zero differences after lifting are partly built in; the panel is read as a descriptive trajectory rather than a formal event study. The noise gap turns negative during strong restrictions, tracking the mobility gap (weekly correlation r = +0.44, ρ = +0.46).
## 3.4. No robust long-run spatial gradient
반면 동을 하나의 점으로 보고 거리두기 기간 전체의 평균을 비교하는 '장기 공간 단면'에서는 이동량과 소음의 뚜렷한 경향이 나타나지 않았다(Fig. 7). 분석 대상(센서 2개 이상) 동 전체에서, 극단값에 강한 순위상관으로 보면 (이동량 감소 vs 도시평균 대비 소음변화)의 관계는 사실상 0이었다(Spearman ρ≈−0.05). 상업 동만 보면 일반 상관계수(Pearson r=+0.38)가 마치 관계가 있는 듯 보였지만, 이 Pearson 값은 소수의 극단적인 동에 민감해 부풀려진 것이다. 같은 동들을 극단값에 강한 방법으로 다시 보면 관계는 0에 가까웠다(Spearman ρ=+0.08, Theil-Sen 기울기≈0; Fig. 7c, 상세는 Supplementary Fig. S1). 정리하면, 이동량-소음 신호는 '같은 날 동 사이의 차이'(M2·M4)에서만 안정적으로 나타나며, 동별 장기 평균을 단순히 지도에 칠하는 방식으로는 잡히지 않는다.
[[IMAGE 7: image8.png 739KB]]
Fig. 7. No robust long-run spatial gradient. (a) Neighbourhood daytime mobility reduction (distancing-era average; peak-phase reductions are larger, cf. Fig. 4). (b) Drift-removed relative noise change over the same period. (c) Neighbourhood-level association (dongs with >= 2 sensors; point size proportional to sensor count; Theil-Sen fit): the robust correlation is near zero overall, and the apparently positive Pearson correlation among commercial dongs is driven by a few extreme neighbourhoods (Supplementary Fig. S1).
## 3.5. Sensor drift and official-station comparison
각 센서의 자기 대비 변화(ΔLday, 해제 후 정상기 기준)를 그대로 그려 보면 언뜻 모순처럼 보이는 결과가 나온다. 규제가 가장 강했던 시기의 주간 소음이 오히려 해제 후 정상기보다 높게 나오는 것이다(최강 제한기 평균 +0.85 vs 해제 후 −0.04 dB). 진단해 보니 이는 코로나 효과가 아니라 여러 해에 걸친 절대 수준의 변동(표류) 때문이었다(Fig. 8a). 4년 내내 가동된 842개 센서의 연평균 Leq,24h는 49.4(2020)→47.6→47.2→47.1 dB(2023)로 해마다 꾸준히 낮아졌고(85% 센서가 하락 추세), 이 하락이 2020-21년을 상대적으로 '시끄럽게' 보이게 만든 것이다(2023년 자료구조 변경 지점에서는 수준 도약이 없었다; +0.02 dB).
이 하락이 실제 소음 변화가 아니라 센서 드리프트와 부합함을, 검교정된 공식 환경소음 측정망(국가소음정보시스템)과 대조해 점검했다. 먼저 절대레벨이 크게 어긋난다. 검교정망 인근의 S-DoT는 표준 LAeq보다 주간 평균 11.7 dB, 도로변에서는 약 16 dB 낮게 읽혀(비동시 연도[2022 vs 2024]·500 m 이내 비동일 지점 대응 기준, Fig. 8b, Supplementary Fig. S2), S-DoT 절대값을 표준 소음도로 해석할 수 없음을 보여 준다. 반면 검교정망 자체의 다년 추세는 안정적이거나 오히려 상승한다(Fig. 8a). 일별 자동측정망(도로 9점)은 2020→2023에 +0.04 dB로 사실상 일정했고, 분기 수동측정망은 주거·일반지역 91점이 +1.93 dB, 도로 60점이 +1.91 dB 상승했다(별도의 도로교통 4점 상시 LAeq도 같은 방향이다). 절대레벨의 치우침은 추세 차분에서 상쇄되므로, 검교정망이 안정·상승하는 동안 S-DoT만 2.3 dB 하락했다는 사실은 그 하락이 실제 환경 변화라기보다 센서의 하향 드리프트와 부합함을 뜻한다(연도 간 계절 구성 차이 등 잔여 교란 가능성은 남는다). 특히 S-DoT 하락이 집중된 비(非)도로 정온지역에서 검교정 주거망이 오히려 상승했다는 점은 그 하락이 실재가 아님을 강하게 시사한다. 한편 이 검교정망의 상승은 우리 용량-반응과 같은 방향임에 주목할 만하다. 검교정 소음은 이동량이 가장 크게 줄었던 2020년에 가장 낮았다가 활동이 회복되며 2023년까지 올라갔는데, 이는 '이동량이 늘면 소음도 는다'는 우리 추정과 부합한다. 즉 다년 절대 수준에서 신뢰할 수 있는 검교정망은 오히려 우리 결과의 방향을 뒷받침하며, 반대로 움직인 것은 드리프트에 오염된 S-DoT의 다년 추세뿐이다. 무엇보다 우리의 용량-반응은 다년 비교가 아니라 '같은 날 동 사이의 차이'로 식별되며(날짜 고정효과가 다년·계절 변동을 모두 흡수한다), S-DoT의 다년 절대 추세를 전혀 사용하지 않는다. 이것이 우리가 절대·시계열 비교 대신 센서 내 상대변화와 '같은 날 동 사이의 차이'(M2)에 의존하는 이유다.
[[IMAGE 8: image10.png 278KB]]
Fig. 8. Sensor drift and calibrated-network comparison. (a) Annual network-mean level change from 2020: the S-DoT network declines by 2.3 dB over 2020-2023 while calibrated environmental-noise stations (automatic daily road stations; manual quarterly general and road stations) stay flat or rise – a pattern consistent with a sensor-drift artefact rather than a real citywide quieting. (b) Nearby S-DoT sensors read on average 11.7 dB below calibrated daytime LAeq (n = 60 station pairs within 500 m; non-simultaneous years, 2022 S-DoT vs 2024 survey; larger at roadside), so absolute S-DoT levels cannot be interpreted as standard noise indicators (Supplementary Fig. S2).
## 3.6. Robustness and sensitivity
주력 추정치가 기계적 산물이 아님을 다섯 가지 민감도 점검으로 확인했다(Fig. 9, Table 7). 첫째, 같은 날짜 안에서 동별 이동량 dose를 동들 사이에 무작위로 재배치한 순열 민감도 점검(300회, 동 단위 셔플)에서 위약 계수는 0 근처에 좁게 분포했고(평균 −0.001, SD 0.023), 실제 추정치 +0.65는 그 분포를 완전히 벗어났다(양측 p=(B+1)/(N+1)=0.003). 다만 날짜별 셔플은 동별 dose의 시계열 지속성을 보존하지 않으므로, 이 p값은 공식적 추론이 아니라 진단적 점검으로 읽는다. 둘째, 센서가 많은 동이 결과를 지배하지 않는지 확인하기 위해 동-일 단위로 동일가중 재추정하면 β=+0.83(SE 0.42, p=0.047)으로 부호와 크기가 유지되었다(주 추정치는 센서 지점 가중 기울기다). 셋째, 이동량 2차항은 유의하지 않아(β²=−0.59, p=0.44) 선형 용량-반응 근사가 타당했다(Fig. 9b). 넷째, 분할추정(Table 6)의 12개 검정을 Benjamini-Hochberg FDR로 보정해도 주간 효과와 봄(MAM)·가을(SON) 계절 효과는 유의하게 남았다(FDR<0.05). 끝으로, 시간 집계의 중복 전송 가능성을 점검하기 위해 기록 시간 수가 하루 24시간을 초과하는 관측(전체의 5.9%)을 제외하고 재추정해도 β=+0.651(SE 0.248)로 사실상 동일했다. 이로써 '같은 날 동 사이의 차이'에서 추정된 신호가 구현·가중·함수형·집계 규칙에 강건함을 확인한다.
[[IMAGE 9: image11.png 241KB]]
Fig. 9. Robustness and sensitivity. (a) Permutation sensitivity check: distribution of the dose-response coefficient when the dong-level mobility dose is reshuffled across neighbourhoods within each date and broadcast to all sensors in the dong (300 shuffles); the actual estimate (+0.65, vertical line on the broken axis) lies far outside the null (two-sided p = (B+1)/(N+1) = 0.003). (b) Binned dose-response after two-way demeaning: decile means (error bars = 95% CI, dong-clustered) lie close to the linear fit, supporting the linear approximation.
Table 7. Robustness and sensitivity checks for the headline two-way fixed-effects dose-response.

[[TABLE 7]]
  | Check || Specification || Result || Conclusion
  | Permutation sensitivity || Dong-level dose reshuffled across dongs within each date, broadcast to all sensors in the dong (300 shuffles) || Null β = −0.001 ± 0.023; actual β = +0.648; two-sided p = (B+1)/(N+1) = 0.003 || Signal is not mechanical
  | Weighting sensitivity || Dong-day panel re-estimated with equal dong weights (main estimate is sensor-location weighted) || β = +0.826 (SE 0.416, p = 0.047) || Sign and magnitude preserved
  | Nonlinearity || Quadratic mobility term added to M2 || β₂ = −0.59 (p = 0.44) || Linear approximation adequate
  | Multiple comparisons || BH-FDR over the 12 segment tests of Table 6 || Daytime, spring (MAM) and autumn (SON) remain significant (FDR < 0.05) || Daytime effect robust
  | Hour-count filter || Sensor-days with more than 24 recorded hours excluded (5.9% of rows) || β = +0.651 (SE 0.248) || Estimate unchanged after exclusion

## 4. Discussion
## 4.1. Summary of findings
본 연구는 측정 이동량에 대한 도시소음의 단계적(graded) 용량-반응을 우리가 아는 한 처음으로 정량 추정했다. 핵심 결과는 네 가지로 요약된다. 첫째, 한 동의 주간 이동량과 그 동의 주간 소음 사이에 통계적으로 견고한 정(+)의 조건부 연관이 존재하지만 그 크기는 작다(이동량 30% 감소 ≈ 소음 −0.2~0.3 dB; 전일 Leq,24h에서도 β=+0.628로 같은 연관이 이어진다). 둘째, 개별적으로 유의한 반응은 주간 결과변수에서만 관찰되며(주야 효과 차이 자체는 미확정), 네 계절 모두 같은 방향으로 나타난다(다중비교 보정 후 봄·가을에서 유의). 야간의 겉보기 연관은 엄밀한 식별에서 시간 교란으로 사라진다. 셋째, 신호는 '같은 날 동 사이의 차이'와 고영향-저영향 궤적 비교에서만 안정적으로 드러나며, 동별 장기 평균을 비교하는 공간 단면에서는 사라진다. 넷째, 저가 IoT 망의 여러 해에 걸친 절대 수준은 센서 표류와 지역별 회복 차이로 교란되어 시계열·절대 비교를 믿을 수 없고, 오직 센서 내 변화와 동 사이 비교에 기반한 설계만이 유효하다. 아울러 다년 절대 수준에서 신뢰할 수 있는 검교정망의 상승 추세는 '이동량이 늘면 소음도 는다'는 본 연관의 방향을 독립 자료로 지지한다.
## 4.2. The magnitude in context
우리의 추정치는 봉쇄기 도시소음이 수 dB 줄었다고 보고한 기존 문헌 [19,20,21,22,23,24]과 대비된다. 그 보고들은 대개 특정 번화가의 절대 소음을 팬데믹 전후로 단순 비교한 것이라, 가장 시끄럽고 가장 크게 비워진 소수 지점의 극단적 변화를 도시 전체로 일반화하기 쉽다. 반면 본 연구는 약 1,100개 지점 전역에서, 센서 내 변화로·이동량의 정도에 따라·전 기간에 걸쳐 한계적 연관(이동량이 한 단위 변할 때의 평균적 소음 변화)을 추정했기에, 도시 평균의 반응은 그보다 더 작게 나타난다. 한 가지 중요한 한정도 있다. 핵심 모형은 '같은 날 동 사이의 차이'만으로 기울기를 추정하므로, 모두가 동시에 덜 움직인 '도시 전체 공통의 이동량 감소' 성분의 효과는 본 설계로 식별되지 않는다. 그 성분에 대해 우리는 상한도 하한도 제시하지 않으며, 본 추정치는 동별 상대 변화에 대한 반응으로 한정해 읽어야 한다. 효과가 작은 것은, 도시소음이 통과 교통이나 간선도로처럼 쉽게 줄지 않는 배경 소음원의 비중이 크다는 점, 그리고 S-DoT의 1 dB 단위 거친 해상도가 측정 잡음으로서 표준오차를 넓힌다는 점과도 들어맞는다. 결국 '봉쇄가 도시를 수 dB 조용하게 만들었다'는 통념은 엄밀한 이동량-소음 관점에서는 과장일 가능성이 높고, 실제 단계적 용량-반응은 견고하지만 작다.
두 추정량은 서로 다른 질문에 답하므로, 직접 비교보다는 우리 기울기를 선행연구의 노출 조건에 대입해 보는 것이 유익하다. 최강 제한 국면에 가장 크게 비워진 상업·업무 동의 주간 체류 감소폭(약 35%, Fig. 4)을 dose로 환산하면 β=+0.648 기준 예상 감소는 약 0.28 dB이다. 날짜 고정효과를 더할 때 주간 계수가 +1.13에서 +0.65로 줄어드는 것은 날짜 공통 교란(도시 추세·망 전체 표류·날씨 등)이 제거된 결과이며, 그 자체가 도시 공통 이동량 효과의 크기를 말해 주지는 않는다. 마드리드의 4-6 dB [19] 같은 선행 보고와의 격차는 따라서 세 성분으로 분해된다. ① 날짜 고정효과가 흡수하는 도시 공통 이동량 감소의 효과(본 설계 미식별), ② 최번화가에 편중된 측정점 선택(도시 평균보다 훨씬 큰 dose와 반응성), ③ 절대 전후 비교에 섞여 드는 배경 추세·센서 변화. 즉 선행 보고의 수 dB 가운데 '동별 상대 이동량 차이'로 설명되는 몫은 0.3 dB 안팎이며, 나머지는 본 설계가 원리적으로 식별하지 않는 성분과 지점 선택에 기인한다. 지각의 관점에서 0.65 dB는 개인이 알아차리기 어려운 크기이지만, 도시 전체 인구가 노출되는 장기 지표의 이동으로서는 의미가 남는다.
## 4.3. Mechanisms behind the patterns
효과가 나타나는 양상은 그 메커니즘과 부합한다. 유의한 반응이 주간 결과변수에서 관찰되고 연중 같은 방향을 유지하는 것은, 주간 소음이 상거래·통근·방문 같은 사람의 활동에 직접 연동되는 반면 야간 소음은 고정 설비나 간선교통 같은 배경원의 비중이 커 이동량 변화에 둔감할 수 있기 때문이다(다만 주야 효과 차이 자체는 통계적으로 확정되지 않았다). 야간의 겉보기 연관이 엄밀한 식별에서 사라진 사례는, 단순한 시간 추세를 인과로 오인할 위험을 보여 준다. 효과가 토지이용에 비교적 균일했다는 점은, 이동량-소음 반응이 특정 장소 유형에 몰려 있지 않고 도시 전반에 얇게 퍼져 있음을 시사한다. 특히 주목할 것은 동 단위 장기 공간 단면에서 깨끗한 경향이 나타나지 않은 점이다. 이는 효과가 없어서가 아니라, 동별 장기 평균 소음 변화가 국지적 공사·도로 변화·센서별 표류 같은 특이 요인에 가려지기 때문이다. 신호는 같은 날 동 사이의 빠른 변동에 있고, 단순한 공간 지도화로는 드러나지 않는다. 실제로 상업 동에서 잠깐 보였던 양의 상관조차 소수의 극단적인 동에 민감한 Pearson이 만든 착시였고, 극단값에 강한 방법으로 보면 0이었다. 이는 고밀도 IoT 자료의 공간 분석이 얼마나 쉽게 잘못된 결론으로 이어질 수 있는지 보여 주는 사례다.
다만 야간의 비유의를 '야간 소음이 활동에 둔감하다'는 결론으로 읽는 것은 성급하며, 세 가지 대안 설명이 함께 간다. 첫째는 검정력이다. 야간 계수의 표준오차(0.63)는 주간(0.245)의 약 2.6배이고, 이 정밀도에서 5% 수준으로 탐지 가능한 최소 효과는 약 1.2 dB이므로, 야간에 주간과 같은 크기의 반응이 있더라도 본 자료로는 탐지되지 않는다. 둘째는 노출 변동이다. 거리두기의 실질 수단이 영업종료시각과 모임 인원이었던 만큼 야간 체류의 동 간 상대 변동은 주간보다 좁을 수 있고, dose 분산이 작을수록 같은 참값도 더 부정확하게 추정된다. 셋째는 측정오차다. 일 단위 품질관리가 구간별 최소 유효시간을 요구하지 않아 일부 Lnight가 소수의 시간으로 계산되며, 이런 고전적 측정오차는 계수를 0 쪽으로 끌어당긴다. 같은 희석 논리가 장기 공간 단면에도 적용된다. 동별 다년 평균에는 센서별 표류라는 큰 무작위 성분이 실려 있고 이는 동 단위 dose와 무관하므로 회귀계수를 체계적으로 0 쪽으로 희석한다. 따라서 공간 null은 '효과 부재'의 증거가 아니라, 노출과 결과 모두에 오차가 실린 단면 설계의 낮은 신호대잡음비와 부합하는 결과다.
## 4.4. Methodological implications
본 연구의 방법론적 기여는 결과 못지않게 중요하다. 저가 고밀도 IoT 소음망으로 도시 환경효과를 측정하려는 시도는 세 가지 함정을 드러낸다. ① 센서마다 다른 고유의 치우침은 센서 내 변화만 쓰는 설계로 상쇄해야 하고, ② 여러 해에 걸친 표류·센서 설치 시점 차이·지역별 회복은 절대 수준의 시계열·공간 비교를 불가능하게 하므로 '같은 날 동 사이 비교'가 필수이며, ③ 도시 전체에 하나의 값으로만 변하는 노출변수는 시간 교란에 취약하므로 동 단위로 측정된 노출이 필요하다. 특히 장기 공간 단면이 안정적으로 영(null)이라는 결과는, 검교정 없이 절대 수준을 그대로 지도에 칠하는 접근이 거짓 패턴(또는 거짓 '무패턴')을 낳을 수 있음을 보여 준다. 이는 SONYC [40]·무선음향센서망 [42]·NoiseCapture [44]처럼 확산 중인 저가 소음 모니터링 [41], 나아가 스마트시티 센서 데이터 [45,46]의 정책 활용 전반에 적용된다. 본 연구가 보여 준 고정효과·극단값에 강한 통계·공식망 교차대조의 절차는, 앞으로의 IoT 기반 도시 환경연구에 참고가 될 수 있다.
## 4.5. Policy implications and limitations
계획·정책 측면에서, 이동량·교통수요 관리가 소음에 주는 직접적 공편익(co-benefit)은 통념보다 작을 가능성이 높다. 봉쇄급 이동량 급감조차 동 단위 상대 변화로는 0.3 dB 안팎의 차이를 낳았을 뿐이다(도시 공통 성분의 효과는 본 설계 밖이다). 또한 본 추정치는 동 단위 상대 이동량 변화에 대한 반응이므로, 도시 전역의 총량 감소를 노리는 정책이나 가로 하나를 차단하는 국지 개입에 그대로 외삽할 수는 없다. 특히 '15분 도시'류 정책은 활동 총량을 줄이기보다 재배치하므로, 그 소음 효과는 본 설계가 식별한 성분과도 성격이 다르다. 그럼에도 어느 방향으로든 이동량 경로만으로 수 dB급 저감을 기대하기 어렵다는 결론은 유지된다. 따라서 이동량 저감 정책은 그 자체로 정당하나, 소음 저감 수단으로는 노면 포장·저소음차량·차폐 등 음원·전파 단계 개입 [17,51]과 병행되어야 한다. 효과가 토지이용에 비교적 균일했다는 점은 상업/주거 구분에 따른 차등 규제보다 광역적 접근의 여지를 시사한다. 서울은 세계 최고 밀도의 메가시티 중 하나이며, 본 결과가 도쿄·홍콩·싱가포르 등 고밀도 동아시아 도시로 일반화되는지는 검증이 필요한 가설로 남긴다.
본 연구에는 몇 가지 한계가 있다. (1) S-DoT는 검교정되지 않은 데다 1 dB 단위의 거친 해상도와 노후에 따른 표류를 가져 절대 수준과 여러 해에 걸친 시계열을 그대로 믿기 어렵다. 그래서 우리는 센서 내 상대변화와 '같은 날 동 사이 비교'로 우회했다. (2) 생활인구는 사람이 그곳에 '있다'는 정보일 뿐, 소음을 실제로 만드는 활동(특히 교통)을 정확히 대변하지는 못하며, 공사·행사처럼 동×날짜 수준에서 이동량과 소음을 함께 움직이는 국지 충격은 날짜 고정효과로도 제거되지 않는다 — 이것이 우리가 추정치를 조건부 연관으로 읽는 이유다. 앞으로 버스·지하철 정류장별 승하차를 동 단위로 연결한 더 직접적인 교통 노출 자료 [35,37]를 쓰면 해석이 더 또렷해질 수 있다. (3) 거리두기 강도 코딩의 정확한 시행 일자는 질병관리청 원자료와의 추가 대조가 필요하다. (4) 기상은 재분석 자료(ERA5)로, 기상청 공식 관측으로 대체할 수 있다. (5) 도시 전체 공통 이동량 성분의 효과는 날짜 고정효과에 흡수되어 본 설계로 식별되지 않으며, 그 크기·방향에 대한 경계도 제시할 수 없다. 검교정된 참조망과 결합해 절대 수준을 복원하는 후속연구가 필요하다. (6) 토지이용은 주간 대비 야간 인구비로 근사했으며, 실제 토지이용 분류로 더 정밀화할 수 있다. (7) 일 단위 품질관리(하루 12시간 이상)는 주간·야간 구간별 최소 유효시간을 따로 요구하지 않으므로 일부 Lnight가 소수의 시간으로 계산될 수 있다(구간별 임계 민감도 분석이 남은 과제다). (8) 분석 창이 S-DoT 개시일(2020-04)부터여서, 2020년 2-3월 1차 유행과 3월 22일 1차 강력 거리두기라는 가장 강한 이동량 충격 구간은 포착하지 못한다. dose 분포의 극단이 절단된 만큼 비선형성(포화·문턱) 탐지력은 제한적이며, 선형 근사의 타당성은 관측된 dose 범위 안에서만 주장된다. (9) 생활인구는 실내·실외 체류를 구분하지 못한다. 자택 대기로 실내 체류가 늘어난 국면에서는 같은 체류 수가 가로의 음원 활동으로 이어지지 않아 노출이 과대 측정되고 계수는 0 쪽으로 감쇠할 수 있으며, 팬데믹기에 늘어난 배달 이륜차처럼 이동량 감소와 반대로 작용한 소음원 변화도 관측된 기울기에 섞여 있을 수 있다. (10) 센서의 설치 맥락(도로 이격거리·설치 고도·벽면 여부)은 관측되지 않으므로, 같은 이동량 변화에 대한 반응성의 지점 간 이질성은 포착되지 않고 β는 이질적 반응의 평균으로 읽어야 한다. (11) 분석은 서울 한 도시에 한정되며, 소음의 주파수·시간 미세구조와 사람의 사운드스케이프 지각 [11,16,33,34]으로의 확장이 남은 과제다.
## 5. Conclusions
우리는 한국의 단계적 거리두기를 자연실험적 맥락으로 삼아, 도시 규모의 IoT 센서망에서 측정 이동량에 대한 도시소음의 graded 용량-반응을 우리가 아는 한 처음으로 정량 추정했다. 그 연관은 통계적으로 견고하지만 작다(이동량 30% 감소 ≈ 주간 소음 0.23 dB, 50% 감소 ≈ 0.45 dB). 개별적으로 유의한 반응은 주간 결과변수에서 관찰되고(주야 효과 차이 자체는 미확정), 네 계절에 걸쳐 같은 방향으로 나타나며(다중비교 보정 후 봄·가을에서 유의), '같은 날 동 사이의 차이'라는 보수적 비교에서 추정된 조건부 연관이다. 동별 장기 평균을 비교하는 공간 단면에서는 신호가 사라지고, 저가 IoT 망의 여러 해에 걸친 절대 수준은 센서 표류로 교란된다. 정책적으로 이는, 적어도 동 단위 상대 변화의 수준에서는 이동수요를 줄이는 도시정책(차 없는 거리·15분 도시)만으로 도시소음을 크게 낮추기 어려울 가능성이 높으며, 소음 목표 달성에는 노면 포장·저소음 차량·차폐 같은 음원·전파 단계 개입이 병행되어야 함을 시사한다. 또한 봉쇄가 도시를 수 dB 조용하게 만들었다는 통념이 소수 핫스팟의 극단치를 일반화한 과대평가일 수 있음을 보여, 소음정책의 비용-편익 기대치를 보정한다.
방법론적으로, 본 연구는 검교정되지 않은 고밀도 IoT 소음망을 도시 환경연구에 신뢰성 있게 쓰는 절차(센서 내 상대변화, 양방향 고정효과, 극단값에 강한 통계, 공식망 교차검증)를 제시하고, 동 단위 순열·가중 민감도 점검으로 그 신호를 확인했다. 절대 수준을 그대로 정책지표로 삼으면 다년 표류에 오도되므로, 스마트시티 소음 모니터링은 within-sensor 변화량 기반 설계를 채택해야 한다. 향후 과제는 세 가지다. 첫째, 정류장별 승하차 같은 더 직접적인 교통 노출을 동 단위로 결합해 효과를 정밀화하고, 둘째, 검교정된 참조망과 결합해 절대 수준을 복원하며, 셋째, 본 설계를 SONYC·WASN·NoiseCapture 등 다른 저가 소음망과 도쿄·홍콩·싱가포르 같은 고밀도 동아시아 도시로 확장해 일반화를 검증하는 것이다.
## CRediT authorship contribution statement
Hyun In Jo: Conceptualization, Methodology, Software, Formal analysis, Data curation, Writing – original draft, Writing – review & editing, Visualization.
## Declaration of competing interest
The author declares no competing financial interests or personal relationships that could have appeared to influence the work reported in this paper.
## Funding
This research received no specific grant from any funding agency in the public, commercial, or not-for-profit sectors.
## Ethics
This study analysed only publicly available, aggregated and de-identified data and did not involve human participants directly; institutional review board approval was therefore not required.
## Data Availability Statement
모든 원시데이터는 공개 공공데이터이다(S-DoT 소음 OA-15969, 생활인구 OA-14991, 지하철 OA-12921, 거리두기 시행연혁 공공데이터포털 15106451, 기상 Open-Meteo ERA5, 행정동 경계 vuski/admdongkor, 검증 4점 도로교통 LAeq OA-15473, 검교정 환경소음 자동·수동 측정망 국가소음정보시스템 noiseinfo.or.kr). 지하철 데이터는 공공누리 제3유형(출처표시+변경금지). 분석 코드는 [저장소 TBD]에 공개 예정이다.
## References
[1] World Health Organization. Environmental Noise Guidelines for the European Region. Copenhagen: WHO Regional Office for Europe; 2018.
[2] World Health Organization. Burden of disease from environmental noise: quantification of healthy life years lost in Europe. Copenhagen: WHO Regional Office for Europe; 2011.
[3] Basner M, McGuire S. WHO Environmental Noise Guidelines for the European Region: A Systematic Review on Environmental Noise and Effects on Sleep. International Journal of Environmental Research and Public Health. 2018;15:519. doi:10.3390/ijerph15030519
[4] Basner M, Babisch W, Davis A, Brink M, Clark C, Janssen S, et al. Auditory and non-auditory effects of noise on health. The Lancet. 2014;383:1325-1332. doi:10.1016/s0140-6736(13)61613-x
[5] Münzel T, Gori T, Babisch W, Basner M. Cardiovascular effects of environmental noise exposure. European Heart Journal. 2014;35:829-836. doi:10.1093/eurheartj/ehu030
[6] Guski R, Schreckenberg D, Schuemer R. WHO Environmental Noise Guidelines for the European Region: A Systematic Review on Environmental Noise and Annoyance. International Journal of Environmental Research and Public Health. 2017;14:1539. doi:10.3390/ijerph14121539
[7] Miedema H, Oudshoorn C. Annoyance from Transportation Noise: Relationships with Exposure Metrics DNL and DENL and Their Confidence Intervals. Environmental Health Perspectives. 2001;109:409. doi:10.2307/3454901
[8] Hammer M, Swinburn T, Neitzel R. Environmental Noise Pollution in the United States: Developing an Effective Public Health Response. Environmental Health Perspectives. 2014;122:115-119. doi:10.1289/ehp.1307272
[9] Schafer RM. The Soundscape: Our Sonic Environment and the Tuning of the World. Rochester (VT): Destiny Books; 1994.
[10] International Organization for Standardization. ISO 12913-1:2014 Acoustics - Soundscape - Part 1: Definition and conceptual framework. Geneva: ISO; 2014.
[11] Aletta F, Kang J, Axelsson Ö. Soundscape descriptors and a conceptual framework for developing predictive soundscape models. Landscape and Urban Planning. 2016;149:65-74. doi:10.1016/j.landurbplan.2016.02.001
[12] Axelsson Ö, Nilsson M, Berglund B. A principal components model of soundscape perception. The Journal of the Acoustical Society of America. 2010;128:2836-2846. doi:10.1121/1.3493436
[13] Kang J, Aletta F, Gjestland T, Brown L, Botteldooren D, Schulte-Fortkamp B, et al. Ten questions on the soundscapes of the built environment. Building and Environment. 2016;108:284-294. doi:10.1016/j.buildenv.2016.08.011
[14] Kang J. Noise Management: Soundscape Approach. Encyclopedia of Environmental Health. 2019:683-694. doi:10.1016/b978-0-12-409548-9.10933-9
[15] Yang W, Kang J. Acoustic comfort evaluation in urban open public spaces. Applied Acoustics. 2005;66:211-229. doi:10.1016/j.apacoust.2004.07.011
[16] Hong J, Jeon J. Relationship between spatiotemporal variability of soundscape and urban morphology in a multifunctional urban area: A case study in Seoul, Korea. Building and Environment. 2017;126:382-395. doi:10.1016/j.buildenv.2017.10.021
[17] Adulaimi A, Pradhan B, Chakraborty S, Alamri A. Traffic Noise Modelling Using Land Use Regression Model Based on Machine Learning, Statistical Regression and GIS. Energies. 2021;14:5095. doi:10.3390/en14165095
[18] Gharehchahi E, Hashemi H, Yunesian M, Samaei M, Azhdarpoor A, Oliaei M, et al. Geospatial analysis for environmental noise mapping: A land use regression approach in a metropolitan city. Environmental Research. 2024;257:119375. doi:10.1016/j.envres.2024.119375
[19] Asensio C, Pavón I, de Arcas G. Changes in noise levels in the city of Madrid during COVID-19 lockdown in 2020. The Journal of the Acoustical Society of America. 2020;148:1748-1755. doi:10.1121/10.0002008
[20] Basu B, Murphy E, Molter A, Sarkar Basu A, Sannigrahi S, Belmonte M, et al. Investigating changes in noise pollution due to the COVID-19 lockdown: The case of Dublin, Ireland. Sustainable Cities and Society. 2021;65:102597. doi:10.1016/j.scs.2020.102597
[21] Aletta F, Oberman T, Mitchell A, Tong H, Kang J. Assessing the changing urban sound environment during the COVID-19 lockdown period using short-term acoustic measurements. Noise Mapping. 2020;7:123-134. doi:10.1515/noise-2020-0011
[22] Rumpler R, Venkataraman S, Göransson P. An observation of the impact of CoViD-19 recommendation measures monitored through urban noise levels in central Stockholm, Sweden. Sustainable Cities and Society. 2020;63:102469. doi:10.1016/j.scs.2020.102469
[23] Steele D, Guastavino C. Quieted City Sounds during the COVID-19 Pandemic in Montreal. International Journal of Environmental Research and Public Health. 2021;18:5877. doi:10.3390/ijerph18115877
[24] Manzano J, Pastor J, Quesada R, Aletta F, Oberman T, Mitchell A, et al. The 'sound of silence' in Granada during the COVID-19 lockdown. Noise Mapping. 2021;8:16-31. doi:10.1515/noise-2021-0002
[25] Maggi A, Muratore J, Gaetán S, Zalazar-Jaime M, Evin D, Pérez Villalobo J, et al. Perception of the acoustic environment during COVID-19 lockdown in Argentina. The Journal of the Acoustical Society of America. 2021;149:3902-3909. doi:10.1121/10.0005131
[26] Montano W, Gushiken E. COVID-19 and soundscape changes due to the lockdown. The case of Lima, Peru. Akustika. 2021;39. doi:10.36336/akustika20213946
[27] Bartalucci C, Bellomini R, Luzzi S, Pulella P, Torelli G. A survey on the soundscape perception before and during the COVID-19 pandemic in Italy. Noise Mapping. 2021;8:65-88. doi:10.1515/noise-2021-0005
[28] Mishra A, Das S, Singh D, Maurya AK. Effect of COVID-19 lockdown on noise pollution levels in an Indian city: a case study of Kanpur. Environmental Science and Pollution Research. 2021. doi:10.1007/s11356-021-13872-z
[29] Sonaviya D, Tandel B. Integrated road traffic noise mapping in urban Indian context. Noise Mapping. 2020;7:99-113. doi:10.1515/noise-2020-0009
[30] Le Quéré C, Jackson R, Jones M, Smith A, Abernethy S, Andrew R, et al. Temporary reduction in daily global CO2 emissions during the COVID-19 forced confinement. Nature Climate Change. 2020;10:647-653. doi:10.1038/s41558-020-0797-x
[31] Brancher M. Increased ozone pollution alongside reduced nitrogen dioxide concentrations during Vienna’s first COVID-19 lockdown: Significance for air quality management. Environmental Pollution. 2021;284:117153. doi:10.1016/j.envpol.2021.117153
[32] Mahato S, Pal S. Revisiting air quality during lockdown persuaded by second surge of COVID-19 of megacity Delhi, India. Urban Climate. 2022;41:101082. doi:10.1016/j.uclim.2021.101082
[33] Erfanian M, Mitchell A, Aletta F, Kang J. Psychological well-being and demographic factors can mediate soundscape pleasantness and eventfulness: A large sample study. Preprint, bioRxiv; 2020. doi:10.1101/2020.10.16.341834
[34] Torresin S, Albatici R, Aletta F, Babich F, Oberman T, Siboni S, et al. Indoor soundscape assessment: A principal components model of acoustic perception in residential buildings. Building and Environment. 2020;182:107152. doi:10.1016/j.buildenv.2020.107152
[35] Deville P, Linard C, Martin S, Gilbert M, Stevens F, Gaughan A, et al. Dynamic population mapping using mobile phone data. Proceedings of the National Academy of Sciences. 2014;111:15888-15893. doi:10.1073/pnas.1408439111
[36] Paez A. Using Google Community Mobility Reports to investigate the incidence of COVID-19 in the United States. Findings. 2020. doi:10.32866/001c.12976
[37] Kalleitner F, Schiestl DW, Heiler G. Varieties of mobility measures: Comparing survey and mobile phone data during the COVID-19 pandemic. Preprint, SocArXiv; 2021. doi:10.31235/osf.io/r78fk
[38] Romanillos Arroyo G. Urban population dynamics during the COVID-19 pandemic based on mobile phone data. Datasets. 2021. doi:10.36443/10259/6864
[39] Yim B, Lee J, Park M. A Two-Timescale Typology of Neighborhood-Scale Commercial Districts in Seoul: Evidence from Mobile Phone De Facto Population Data. Sustainability. 2026;18:4326. doi:10.3390/su18094326
[40] Bello J, Silva C, Nov O, Dubois R, Arora A, Salamon J, et al. SONYC. Communications of the ACM. 2019;62:68-77. doi:10.1145/3224204
[41] Mydlarz C, Salamon J, Bello J. The implementation of low-cost urban acoustic monitoring devices. Applied Acoustics. 2017;117:207-218. doi:10.1016/j.apacoust.2016.06.010
[42] Alías F, Alsina-Pagès R. Review of Wireless Acoustic Sensor Networks for Environmental Noise Monitoring in Smart Cities. Journal of Sensors. 2019;2019:1-13. doi:10.1155/2019/7634860
[43] Sevillano X, Socoró J, Alías F, Bellucci P, Peruzzi L, Radaelli S, et al. DYNAMAP - Development of low cost sensors networks for real time noise mapping. Noise Mapping. 2016;3. doi:10.1515/noise-2016-0013
[44] Picaut J, Bocher E, Aumond P, Petit G, Fortin N. Exploiting data from the NoiseCapture application for environmental noise measurements with a smartphone. INTER-NOISE and NOISE-CON Congress and Conference Proceedings. 2021;263:3149-3159. doi:10.3397/in-2021-2316
[45] Boumchich A, Picaut J, Bocher E. Using a Clustering Method to Detect Spatial Events in a Smartphone-Based Crowd-Sourced Database for Environmental Noise Assessment. Sensors. 2022;22:8832. doi:10.3390/s22228832
[46] Peng B, Wang K, Abdulla W. An Integrated Hierarchical Wireless Acoustic Sensor Network and Optimized Deep Learning Model for Scalable Urban Sound and Environmental Monitoring. Applied Sciences. 2025;15:2196. doi:10.3390/app15042196
[47] Cui H, Zhang L, Li W, Yuan Z, Wu M, Wang C, et al. A new calibration system for low-cost Sensor Network in air pollution monitoring. Atmospheric Pollution Research. 2021;12:101049. doi:10.1016/j.apr.2021.03.012
[48] Abadie A, Athey S, Imbens GW, Wooldridge JM. When should you adjust standard errors for clustering? The Quarterly Journal of Economics. 2023;138:1-35. doi:10.1093/qje/qjac038
[49] Guo J, Qu X. Fixed effects spatial panel data models with time-varying spatial dependence. Economics Letters. 2020;196:109531. doi:10.1016/j.econlet.2020.109531
[50] Craig P, Campbell M, Deidda M, Dundas R, Green J, Katikireddi S, et al. Using natural experiments to evaluate population health interventions: a framework for producers and users of evidence. Public Health Research. 2025:1-59. doi:10.3310/jtyw6582
[51] Vogiatzis K, Remy N. Soundscape design guidelines through noise mapping methodologies: An application to medium urban agglomerations. Noise Mapping. 2017;4:1-19. doi:10.1515/noise-2017-0001
[52] Pascale A, Mancini S, d'Orey PM, Guarnaccia C, Coelho MC. Correlating the effect of Covid-19 lockdown with mobility impacts: A time series study using noise sensors data. Transportation Research Procedia. 2022;62:115-122. doi:10.1016/j.trpro.2022.02.015
[53] Aletta F, Brinchi S, Carrese S, Gemma A, Guattari C, Mannini L, et al. Analysing urban traffic volumes and mapping noise emissions in Rome (Italy) in the context of containment measures for the COVID-19 disease. Noise Mapping. 2020;7:114-122. doi:10.1515/noise-2020-0010
[54] Hemker F, Haselhoff T, Brunner S, Lawrence BT, Ickstadt K, Moebus S. The role of traffic volume on sound pressure level reduction before and during COVID-19 lockdown measures - a case study in Bochum, Germany. International Journal of Environmental Research and Public Health. 2023;20:5060. doi:10.3390/ijerph20065060