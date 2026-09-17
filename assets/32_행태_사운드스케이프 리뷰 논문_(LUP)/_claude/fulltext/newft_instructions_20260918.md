# 추가 확보 전문 평가 지시문 (2026-09-18) — 전문 적격성 판정 · 자료 추출 · MMAT · 효과 통계

리뷰: "Soundscape and human behaviour in urban open space: a systematic review and meta-analysis of forward and reverse pathways"
(OSF 등록 osf.io/7ew8q). 8월에 전문 평가한 코퍼스(포함 98편)와 **같은 기준·같은 양식**으로 평가한다.

## 0. 공통 원칙
- 배정된 레코드의 전문 텍스트 `fulltext/txt/TXT_<id>.txt` 를 **처음부터 끝까지 전부** 읽는다(표·그림 캡션·부록 포함). 앞부분만 읽고 판정하지 않는다.
  파일이 길면 Read 도구의 offset/limit 로 나눠 끝까지 읽는다.
- 근거는 원문 위치(절 번호·표 번호·쪽 `[page n]`)와 함께 적는다. 원문에 없는 수치·사실을 만들지 않는다. 보고되지 않았으면 `NR`.
- 판정 사유·메모는 **한국어**, 통제어휘 칸은 아래 **영어 어휘 그대로**, 추출 자유서술(exposure·behaviour_measure·key_finding·effect_stats)은 **영어**로 쓴다.
- CSV 는 반드시 Python `csv` 모듈(`encoding="utf-8-sig"`, `newline=""`)로 쓴다(쉼표·따옴표 이스케이프).

## 1. 적격 기준 (등록본 그대로)
포함: 아래를 **모두** 충족
1. 사람 대상 실증 연구, **영어로 게재된 동료심사 학술지 논문**
2. 옥외 또는 반옥외의 도시·경관 공공공간 이용자: 공원, 가로, 광장, 수변, 캠퍼스 옥외공간, 주거단지 옥외공간(공용 마당·커뮤니티 활동공간), 레크리에이션 공간(도시 숲·교외 숲·국립공원 탐방로 포함)
3. 노출: 음환경(소음, 교통음, 자연음, 새소리, 물소리, 음악·부가음, 사람 소리, 사운드스케이프 질·구성, 음향 개입). **역방향**은 이용자 행태·활동이 음환경 형성이나 사운드스케이프 지각·평가에 미치는 영향
4. 결과: **관찰 가능한 행태**(측정·관찰·자기보고 모두 인정) — 5도메인
   movement(보행속도·경로·횡단·접근/회피·통행수단 선택) · staying(체류시간·머무름·착석·멈춤) · space-use(방문·이용빈도·자리 선택·공간 분포) · social(대화·집단 상호작용·도움/반사회 행동·군중 행동) · activity(신체활동·여가 활동 유형)
   역방향 연구의 결과는 음환경(측정 음향) 또는 사운드스케이프 지각·평가
5. 설계: 현장·실험실·VR 실험, 현장 관찰, 자연실험, 설문, 센서·빅데이터 관찰, 질적 연구

배제 코드(번호 순서대로 적용하고 **처음 해당하는 코드 하나**):
| 코드 | 내용 |
|---|---|
| X1 | 동물·야생 대상(사운드스케이프 생태학) |
| X2 | 옥외·반옥외 도시·경관 공공공간이 아님(실내 상업시설·직장·병원·차량 내부·주택 실내, 공간을 특정하지 않은 거주지 단위 노출) |
| X3 | 음환경 노출·조작 없음(역방향: 행태→음환경 경로 없음). 음환경이 **통제변수로만** 들어간 경우 포함 |
| X4 | 관찰 가능한 행태 결과 없음(지각·성가심·선호·회복감·건강·생리 결과만) |
| X5 | 실증 연구 아님(리뷰·논평·순수 시뮬레이션·방법론 제안만) |
| X6 | 동료심사 학술지 논문 아님(학회 초록·워킹페이퍼·학위논문·보고서·업계지) |
| X7 | 본문이 영어가 아님(제목·초록만 영어인 경우 포함 — 반드시 본문으로 확인) |

## 2. 경계 규칙 (8월에 정한 규칙 — 그대로 적용하고 `boundary_rule` 칸에 코드 기록)
| 코드 | 상황 | 판정 |
|---|---|---|
| R1 | 거주지 소음 노출 × 특정 공공공간과 연결되지 않은 신체활동·건강행태(역학 연구) | FINAL_EXCLUDE (X2) |
| R2 | 행태 결과가 **행동 의도·의향**(재방문 의향, 지불 의사, 친환경 행동 의도 등)뿐 | SENS_ONLY |
| R3 | 역방향 연구이고 결과가 사운드스케이프 지각·평가 | FINAL_INCLUDE |
| P1 | 옥외 장면을 실험실·VR 로 재현한 연구 | 데이터베이스 경로(branch=DB) = FINAL_INCLUDE / 인용 추적·보조 검색(branch=CT·OAS) = SENS_ONLY |
| P3 | 음환경이 통제변수·공변량으로만 들어감 | FINAL_EXCLUDE (X3) |
- 8월 선례: 주거단지 공용 옥외공간(중정·커뮤니티 활동공간) 연구 포함 / 운전자 경적 사용이 가로 음환경에 미치는 영향 = 역방향 포함 / 소음과 통행수단 선택 = movement 포함 / 거주지 소음 성가심·정신건강 설문 = X4 또는 R1 배제.
- 판단이 정말 갈리면 판정은 내리되 `confidence=low` 로 두고 rationale 에 갈리는 지점을 적는다(UNCERTAIN 은 쓰지 않는다).
- 같은 표본을 쓴 기존 코퍼스 연구가 있는지 확인: `fulltext/corpus_v4_extraction.csv` 의 제목·저자·국가·표본으로 대조하고, 겹치면 rationale 에 `표본 공유 의심: <uid>` 를 적는다(판정은 적격성 기준대로).
- 원문 버전이 서지와 다르면(예: 학술지 논문 레코드인데 파일이 워킹페이퍼판) rationale 에 적고, 내용으로 판정한다.

## 3. 출력 파일 (배치 이름 `nft_XX` — 지시받은 이름 사용) → `fulltext/newft_results/`

### 3.1 `nft_XX_verdict.csv` — 배정된 모든 레코드 1행씩
`id, branch, verdict, reason_code, boundary_rule, confidence, rationale`
- verdict: `FINAL_INCLUDE` / `SENS_ONLY` / `FINAL_EXCLUDE`
- reason_code: 배제일 때 X1–X7, 아니면 빈칸 · boundary_rule: R1/R2/R3/P1/P3 또는 빈칸 · confidence: high/medium/low
- rationale: 한국어 1–3문장, 원문 위치 포함(예: "§2.3·Table 2 에서 …")

### 3.2 `nft_XX_extract.csv` — FINAL_INCLUDE·SENS_ONLY 만
`id, branch, final_verdict, year, journal, title, country, setting, design, sample_n, exposure, behaviour_domain, behaviour_measure, measurement_method, direction, key_finding, effect_stats`
- year·journal·title: 텍스트 머리의 `### RECORD`/`### TITLE` 값 그대로
- country: 자료 수집 국가 영문(복수면 `; `), 미보고 `NR`
- setting: 선두어 하나 `street` / `park` / `square` / `campus` / `residential` / `recreation` / `waterfront` / `VR-lab` / `mixed` + 괄호 설명(장소명·규모). 실험실 재현은 `VR-lab (…)`
- design: 선두어 하나 `survey` / `mixed` / `field-experiment` / `natural-experiment` / `quasi-experiment` / `field-observation` / `observational` / `sensor-bigdata` / `lab-VR-experiment` / `qualitative` / `NR` + 필요시 괄호 설명
  (`mixed` = 관찰+설문처럼 둘 이상의 자료수집 결합. 객관 음향측정 + 설문은 `survey`)
- sample_n: 참가자 수 우선, 분석 단위가 다르면 괄호로(예: `412 participants (analytic unit: 29 street segments)`)
- exposure: 음환경 노출·조작 내용(음원, 레벨, 측정 방법), 역방향이면 행태 노출과 음환경 결과를 함께
- behaviour_domain: `movement` / `staying` / `space-use` / `social` / `activity` 를 `; ` 로(해당 전부)
- behaviour_measure: 행태 지표 정의
- measurement_method: `self-report` / `observation` / `sensor-GPS-video` 를 `; ` 로(해당 전부) + 괄호 설명. 소음계처럼 **노출만** 잰 장비는 sensor 로 보지 않는다
- direction: `forward`(음환경→행태) / `reverse`(행태→음환경·사운드스케이프 평가) / `both`
- key_finding: 행태와 음환경의 관계에 관한 주요 결과 1–3문장(방향·크기)
- effect_stats: 행태–음환경 관계의 통계를 원문 그대로(표 번호 포함), 없으면 `NR`

### 3.3 `nft_XX_mmat.csv` / `nft_XX_detail.csv` — FINAL_INCLUDE·SENS_ONLY 만 (MMAT 2018)
mmat: `id, mmat_category, S1, S2, Q1, Q2, Q3, Q4, Q5, n_yes, quality_tier, note`
detail: `id, item, item_no, verdict, rationale` — 연구당 **5행**(선택한 범주의 5문항)
- mmat_category: `Qualitative` / `Quantitative RCT` / `Quantitative non-randomised` / `Quantitative descriptive` / `Mixed methods`
- 범주 배정(8월 교훈): 객관 음향측정 + 설문 = 정량(혼합 아님) / 조건 배정이 무작위가 아닌 현장실험 = non-randomised / 사례 기술만 있고 질적 자료수집·분석이 없으면 qualitative 아님 / 혼합은 질적 성분과 정량 성분이 모두 실제로 있을 때만 / 상관·회귀 관찰 연구 대부분 = descriptive 또는 non-randomised(노출군 비교·교란 보정이 핵심이면 non-randomised)
- 문항 판정 `Y` / `N` / `CT`(판단불가: 보고 없음). S1(명확한 연구질문)·S2(자료가 질문에 답할 수 있음) 도 Y/N/CT
- item_no 와 item(영문, 아래 그대로):
  - 1.1 qualitative approach appropriate to answer research question · 1.2 qualitative data collection methods adequate · 1.3 findings adequately derived from the data · 1.4 interpretation of results sufficiently substantiated by data · 1.5 coherence between data sources, collection, analysis and interpretation
  - 2.1 randomization appropriately performed · 2.2 groups comparable at baseline · 2.3 complete outcome data · 2.4 outcome assessors blinded to the intervention · 2.5 participants adhered to the assigned intervention
  - 3.1 participants representative of the target population · 3.2 measurements appropriate regarding outcome and exposure · 3.3 complete outcome data · 3.4 confounders accounted for in design and analysis · 3.5 exposure/intervention administered as intended
  - 4.1 sampling strategy relevant to address research question · 4.2 sample representative of the target population · 4.3 measurements appropriate · 4.4 risk of nonresponse bias low · 4.5 statistical analysis appropriate to answer research question
  - 5.1 adequate rationale for using a mixed methods design · 5.2 components effectively integrated to answer research question · 5.3 outputs of integration adequately interpreted · 5.4 divergences and inconsistencies adequately addressed · 5.5 components adhere to quality criteria of each tradition
- Q1–Q5 = 선택 범주 문항 1–5 판정, n_yes = Y 개수, quality_tier: Y 4–5 `high` · 3 `moderate` · 0–2 `low`
- rationale·note: 한국어, 원문 위치 포함

### 3.4 `nft_XX_es.csv` — FINAL_INCLUDE·SENS_ONLY 중 아래 **메타분석 클러스터 후보**가 있을 때만
`id, cluster, outcome_measure, comparison, statistic_type, values_verbatim, n_info, location, quote, computable`
현재 메타분석 클러스터(해당 여부만 판단해 후보를 기록한다 — 포함 결정은 하지 않는다):
| cluster | 정의 |
|---|---|
| walking_speed | 음환경 조건 간(예: 자연음 vs 교통소음, 음악 vs 무음) 보행속도 비교 |
| staying | 음환경 조건 간 체류시간·머무름 비교 |
| social | 음환경 조건 간 사회적 상호작용·도움 행동 빈도 비교 |
| correlation | 음 지표(측정 음향 또는 지각)와 행태 지표 간 상관계수(r, rho) |
| other | 위에 없지만 효과크기 계산이 가능한 행태–음환경 비교 |
- statistic_type: `M_SD_n` / `test_stat`(t, F, χ²) / `correlation` / `2x2_counts` / `OR` / `regression` / `other` / `none`
- values_verbatim: 원문 수치 그대로(단위·표기 포함) · n_info: 조건별 n, 참가자 수와 관측 수 구분 · location: 표·그림·절 · quote: 해당 문장 원문 짧게
- computable: 조건별 평균·SD·n, 또는 t/F(분자 df 1)+n, 또는 r+n, 또는 2×2 빈도가 모두 있으면 `YES`, 아니면 `NO`

## 4. 끝내기 전 자가 점검
- verdict 행 수 = 배정 레코드 수, id 누락·중복 없음
- INCLUDE/SENS 인 id 는 extract·mmat 에 모두 있고 detail 은 id 당 정확히 5행, item_no 첫 자리가 mmat_category 와 맞음
- CSV 를 다시 읽어(`csv.DictReader`) 열 이름·행 수를 출력해 확인
- 최종 보고(부모에게): 레코드별 `id | verdict | reason_code | 한 줄 사유` 표 + 판단이 갈린 레코드 목록. 한국어, 800 단어 이내
