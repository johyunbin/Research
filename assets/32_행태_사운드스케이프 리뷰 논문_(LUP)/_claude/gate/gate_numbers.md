# Gate — 수치·데이터 정합성 독립 검증 (paper32)

- **검증 대상**: `01_논문작업/Manuscript_EN_20260806_draft.md`
- **검증 축**: 정량 주장의 근거파일 재현성 (수치·백분율·분모·정본버전)
- **검증일**: 2026-08-06 · 검증자: 독립 게이트(원고 미작성자)
- **재산출 방법**: `corpus_v3_verdicts.csv`(154행) · `corpus_v3_extraction.csv`(100행) · `table1_v2.csv`(100행) ·
  `quality_v2.csv` · `quality_detail_v2.csv` · `evidence_counts_v2.csv` · `ct_verdicts_final.csv`(54행) ·
  `ma/ma_v2_summary.md` · `ma/ma_sensitivity_v2.md` · `ma/ma_*_input.csv` · `design_matrix_v2.csv` ·
  `quality_truncation_effect.md` · `prisma_flow.md` 를 pandas 로 직접 로드해 원고 서술과 1:1 대조.
  집계는 `final_verdict == 'FINAL_INCLUDE'`(96편) 필터를 명시적으로 걸어 재계산했고,
  국가 집계는 `make_fig7_geo_time.py` 의 `norm_countries()` 정규화기를 그대로 재사용해 96편/100편 두 기준으로 각각 산출.

---

## 판정

**FAIL** — 결론 서사의 핵심 수치("fifty years" 반복 간격)가 근거파일의 연도(1975·1988 = 13년)와
어긋나고, Methods 가 선언한 적격 규칙(랩 재현 → 민감도 전용)이 코퍼스에서 실제로 집행되지
않았으며(랩 6편이 본분석 96편 안에 있음), MMAT 문항 백분율 2건이 재산출값과 불일치한다.

---

## 불일치

| 원고 위치(절) | 원고 서술 | 근거 파일 값 | 심각도 |
|---|---|---|---|
| Abstract · §3.4 · §4.3 · §5 (4회 반복) | "independent field experiments **fifty years apart**" / "separated by fifty years" / "replicated across fifty years" | MA3 기여 2편 = Mathews & Canon **1975**(CT0025) · Moser **1988**(uid 14) → 간격 **13년**. `table1_v2.csv` year 필드·`ma_v2_new_inputs.csv`. (`ma_v2_summary.md` 도 "50년 간격"으로 잘못 적혀 있어 원고가 이를 승계) | **critical** |
| §2.2 · §2.4 (경계규칙) | "Laboratory studies … **reserved for sensitivity analysis**" / "laboratory reproduction of outdoor scenes → **sensitivity analysis only**" | 규칙이 집행되지 않았다. `table1_v2.csv` 에서 `setting == lab(outdoor scene)` 6편(**CT0025·661·666·999·1122·1272**, 그중 5편은 `design == lab experiment`)이 **FINAL_INCLUDE 96편 안에** 있다. 반대로 SENS_ONLY 4편(469·818·1009·CT0356)은 전부 행동의향 사유(`ruling_audit.csv` rule=R2)이고 랩 사유는 **0편**. `deviation_log.md` D1-2 의 P1 규칙과도 모순 | **critical** |
| §4.2 | "the **pooled studies** and the two new high-quality walking papers **are all Czech**" | MA1 풀 기여 = 461(Czech) · 532 Exp1·Exp2(Czech) · **617 = Berkouk 2020, Algeria**. 근거파일(`design_implications_v2.md`)이 "모두 체코"라고 한 목록은 461·532·CT0090·CT0166 이며 **617을 포함하지 않는다**. 또한 §3.4가 말한 고품질 3편 중 **CT0335 = Germany** | **critical** |
| §3.3 · §4.4 | "low **non-response bias in 41%**" / "non-response bias in **two in five**" | MMAT 4.4 `Y` = **16/51 = 31.4%**(96편 기준) · 18/54 = 33.3%(100편 기준). ≈ 1/3 이지 2/5 가 아니다. `quality_v2.csv` Q4 · `quality_detail_v2.csv` item 4.4 두 경로로 동일값 | **major** |
| §3.3 | "adequate control of confounding in **37%** of non-randomised studies" | MMAT 3.4 `Y` = **9/22 = 40.9%**(96편) · 9/23 = 39.1%(100편). 37% 는 어느 분모로도 재현되지 않음. **원고의 41%/37% 는 4.4↔3.4 가 서로 뒤바뀐 뒤 한쪽이 추가로 틀어진 형태**(41% ≈ 3.4의 40.9%) | **major** |
| §3.3 · §4.4 | "the **two randomised studies score lowest** in the corpus" | RCT 2편 = 461(`n_yes`=2) · 532(`n_yes`=1). 그러나 **`n_yes`=0 인 연구가 4편**(uid 7·123·791·1122)으로 더 낮고, `n_yes`=1 은 532 포함 15편이 동률. "최하위"는 성립하지 않음 (`quality_v2.csv`) | **major** |
| §3.1 | 초록 확보군이 "converted at **46% (17 of 37)**" | 정본 판정표 `ct_verdicts_final.csv`(보정 후) 교차표 = RETRIEVE 37편 중 INCLUDE 15 + SENS 1 = **16 → 43.2%**. 46%(17/37)는 **REC 410 을 X7(중국어)로 재판정하기 전 값**이며 `ct_fulltext_summary.md`·`prisma_flow.md` 에만 남아 있다(`ct_corrections.md` 보정1 미반영). B군 0%(0 of 17)는 일치 | **major** |
| §2.5 | 절단 재평가에서 "seven changed at least one item and **four changed tier**" | `quality_truncation_effect.md` 표에서 등급이 바뀐 것은 **5편** — CT0090(low→high) · CT0166(low→high) · CT0184(**high→moderate**) · CT0335(low→high) · CT0348(low→moderate). CT0184 누락. (§4.4의 "moved four … **down**"은 상향 4편 기준이라 정확) | **major** |
| §4.4 | "The bias is **one-directional** and invisible without a deliberate audit" | 같은 감사표의 **CT0184 는 high → moderate 로 하향**(절단본이 오히려 과대평가). 단방향이 아님 | **major** |
| §4.2 | "**the only study that altered the environment itself** is rated low and reports a difference of **0.06 m/s**" (앞 문장 = "Three of the four pooled effects … headphones") | 0.06 m/s(1.24→1.18)는 **uid 941**(`design_implications_v2.md` L9)이고 **941은 MA1 풀에 없다**. 네 번째 풀 기여는 **617**(Algeria, field-**observation**)이며 원자료는 0.84 vs 0.81 = **0.03 m/s**(`ma_walking_input.csv`). 문장 구조가 "풀 안의 네 번째 효과"로 읽혀 오도 | **major** |
| §4.5-5 | "**only two studies** give rules for predicting the sound a design will generate" | 81편 구코퍼스 수치. `design_implications_v2.md` §A.3 = "480·341 **두 개뿐이던** … CT0001·CT0348·CT0137 **로 늘었다**"(총 5편), `design_matrix_v2.csv` L6 = "밀도→음압 **예측식 1종에서 3종으로**"(480·CT0001·CT0348) | **major** |
| §3.5 | "110 forward / 74 reverse / 22 bidirectional: **40%** of the evidence is not about sound acting on people" | 제시된 세 수(합 206)로는 40%가 나오지 않는다. 74/206 = **35.9%**, (74+22)/206 = **46.6%**. 40%의 출처는 `rebuild_evidence_map.py:168` = `74/(110+74)` = **40.2%** 로 **both 22건을 분모에서 뺀** 값이며 원고는 이 분모를 밝히지 않는다. 같은 절·§5의 "nearly half"(46.6%)와 분모가 서로 다름 | **major** |
| §3.2 vs Fig. 7 | "**38 of 96** studies (40%) … China, Spain (7), UK (6), Germany (5)" | 본문 수치는 96편 기준으로 **재현됨**(38/96=39.6%). 그러나 Fig 7의 산출물 `geo_time_counts.csv` 는 **100편 전체** 기준이라 **China 40 · UK 7 · Germany 6**. 그림과 본문이 서로 다른 숫자를 보여준다 | **major** |
| §3.2 | 표본크기 "IQR **30–400**" | `table1_v2.csv` n(67편 보고) → Q1 **30.5** · Q3 **377.5**. 상한이 ~6% 올려 표기됨(중앙값 120·최대 13,322은 정확히 일치) | minor |
| §1.2 | "**Two-thirds** of the eligible literature appeared **after 2020** (68 of 96)" | 68/96 = **70.8%**("two-thirds"=66.7%). 또 68은 `year >= 2020` 값이고 `year > 2020` 은 63. §3.2·§4.6은 "71%"·"2020 or later"로 맞게 씀 — 절 간 불일치 | minor |
| §3.7 | "**Every** behavioural domain × sound source cell … **is populated**, with one exception" | `evidence_map_v2.md` §1에서 항공기소음 열의 **세 셀이 0**(movement·staying·social). 나머지 셀은 모두 채워짐. "빈 셀 없음"은 문자 그대로는 거짓 | minor |
| §3.4 | 고품질 보행 3편 중 "one **analysed steps rather than participants**" | CT0335 추출 기록은 `sample_n = 16(참가자)` · `effect_stats = **NR(본문 절단으로 보행속도 통계 미수록)**`. 풀 진입 불가 사유가 원고 서술과 다름 | minor |
| §2.6 vs MA1 | "Where a study contributed **more than one estimate to the same cluster and contrast frame, one effect was selected**" | MA1 k=4 는 **3편**에서 나왔다 — uid **532가 동일 대비 프레임("birdsong vs crowded city noise")으로 Exp1·Exp2 두 효과를 기여**(`ma_walking_input.csv`, `ma_sensitivity_v2.md` "기여: 461·532·532·617"). `design_matrix_v2.csv` L9 는 "MA1 k=4 (**3편**·532는 2개 실험)"이라 명시하는데 원고는 이 사실을 어디에도 쓰지 않음 | minor |
| §4.2 vs §3.4 | "the **two** new high-quality walking papers" | §3.4는 "**Three** high-quality walking studies exist in the corpus"(CT0090·CT0166·CT0335). 절 간 개수 불일치 | minor |
| Table 1 (CT0025 행) | — | `table1_v2.csv` 의 CT0025 `setting = lab(outdoor scene)` 이지만 `corpus_v3_extraction.csv` 는 `street (Exp2 … 보도; Exp1은 실내 실험실 대기실)`. MA3에 들어간 효과는 **현장(Exp2)** 이므로 Table 1 표기가 원고 서술(field experiment)과 어긋남 | minor |

---

## 근거 없는 주장

정량 서술 중 근거 파일에서 **확인 불가**한 것:

1. **"two countries"** (§4.3 "independent replication across fifty years, **two countries** and two noise sources" · §5 동일) — Moser 1988 = France 로 기록돼 있으나 **CT0025 의 country 는 `NR (본문에 수집국 미명시; 저자 소속 University of New Hampshire, USA)`**. 코퍼스는 수집국을 명시적으로 "미보고"로 판정했으므로 "두 나라"는 추출 기록에서 지지되지 않는다(저자 소속에서 유추한 값). *"two noise sources"(잔디깎기·도로공사)는 확인됨.*
2. **"~50 to ~87 dB(**C**)"** (§3.4) — `corpus_v3_extraction.csv` CT0025 원문 발췌는 `87 dB` · `50 dB` 로 **가중치 표기가 없다**. dB(C)는 `ma_v2_summary.md`·`design_matrix_v2.csv` 의 2차 서술에만 존재.
3. **§1.4 위치설정 주장 전부** — "The nearest adjacent review (Cities & Health, 2026) covers physical activity only", "Zhang et al. (2025, *Landscape and Urban Planning*) meta-analysed landscape → perception" — 두 문헌 모두 코퍼스·근거파일 어디에도 없다(코퍼스는 실증연구만 포함). 별도 출처 확인 필요.
4. **§4.2 "three of them on the same 1.75–1.8 km circuit"** — `design_implications_v2.md` 이 지지하는 것은 "CT0090·CT0166 이 **동일** 순환로를 명시" + "461·532 도 **루트 길이가** 1.8 km 로 같다"이다. **동일 회로 3편**은 근거파일에 없다(동일 확인 2편 + 길이 일치 2편). *"two sharing an identical belt-camera and annotation protocol"(532·CT0166)은 확인됨.*
5. **§4.6 "Two publication pairs share samples; only one member of each pair entered any pool"** — 쌍 자체는 확인(CT0126↔CT0322 · CT0007↔761). 다만 CT0007↔761 은 `ct_corrections.md` 보정2에서 **"표본 중복 *의심*"(플래그만, 판정 변경 없음)** 으로 기록돼 있어 "share samples" 단정은 근거보다 강하다.

---

## 확인된 것

재산출에 성공(원고값 = 근거파일값):

- **PRISMA DB 갈래** 2,073(WoS 1,010·Scopus 850·PubMed 213) − 757 → 1,316 − 1,127 → 189 − 89(그중 84 uncertain) → 100 − 16 − 3 → **81**. 산술 전부 성립.
- **PRISMA 인용추적 갈래** 시드 84편 → 2,073 신규(backward 413 + forward 1,660 = 2,073 ✓) → 우선순위 배제 1,645(T3 26 + T4 1,619 ✓) → 428(T1 11 + T2 417 ✓) → 146 → 79 → 54 → −38 −1 → **15**. B군 전환율 **0% (0 of 17)** = `ct_verdicts_final.csv` 교차표와 일치.
- **코퍼스** FINAL_INCLUDE **96**(db 81 + ct 15) · SENS_ONLY **4**(db 3 + ct 1) · 총 **100** — `corpus_v3_verdicts.csv` 교차표로 직접 확인. 인용추적 기여율 15/81 = **18.5% → 19%** ✓.
- **미확보 합계** 89 + 25 = **114** ✓ · A군 회수율 37/40 = **92.5% → 92%** ✓ · 미확보 제목판정분 39−17 = **22** ✓.
- **품질 22/40/34** — `quality_v2.csv` 를 96편으로 필터하면 정확히 high 22 · moderate 40 · low 34(100편이면 24/41/35). 갈래별 18/33/30 + 4/7/4 로도 교차 검증됨. MMAT 4.2 = 10/51 = **19.6% → 20%** ✓.
- **절단 감사** 인용추적 16편 중 **13편** 한도 초과 ✓ · 문항 변경 **7편** ✓ · **2등급 이동 3편**(CT0090·CT0166·CT0335) ✓ · "down 4편"(§4.4) ✓.
- **메타분석 4클러스터 전 수치 일치** — MA1 k=4 g=−0.500 [−1.411,+0.412] p=.179 I²=75.7% · MA2 k=3 g=+0.313 [−0.076,+0.702] p=.074 I²=54.4% · MA3 k=4 g=+0.646 [+0.188,+1.104] p=.021 I²=48.6% · MA4 k=7 r=+0.409 [+0.227,+0.564] p=.002 I²=90.5%. Abstract 반올림(+0.65/+0.19~+1.10, +0.41/+0.23~+0.56)도 정확.
- **MA3 민감도** CT0025 제외 시 k=3 · **p=.053** ✓ · **MA4 인용추적 제외 시 r=.426** ✓ · **MA1 low 제외 시 k=0** ✓ · 민감도 **7축** = ①LOO ②품질 ③관측 n ④rho ⑤가정 ⑥대안정의 ⑦갈래 — Methods 나열과 정확히 대응.
- **Mathews & Canon 2×2** 20/40 vs 5/40 · OR = (20/20)/(5/35) = **7.00** ✓ · Chinn d = ln7·√3/π = **1.0728 → +1.07** ✓ · **F(1,76) = 20.00** ✓ · 클러스터 내 최대효과 ✓ · 기계소음 2건 vs 자연음 2건 ✓.
- **기술통계** 68/96 = **70.8% → 71%** ✓ · 중앙 연도 **2022** ✓ · 범위 1975–2026 ✓ · 중앙 n **120** ✓ · 최대 **13,322 street segments**(uid 951, Helsinki Strava) ✓ · street **36** · park **28** ✓ · field experiment **17 = 17.7% → 18%** ✓ · mixed 31 · survey 21 ✓ · China **38/96 = 39.6% → 40%**(96편 기준) ✓.
- **방향** forward **110** · reverse **74** · both **22** ✓ · space-use 25 vs 23 ✓ · activity 18 vs 21 ✓ · social 19 vs 14 ✓ (분모 문제는 위 표 참조).
- **측정세대** G3 4 → **20** · G1 14 → **37** · G2 11 → **26** ✓ · 밴드합계 G1 52 · G2 40 · G3 24(총 116) ✓ · **2세대 이상 동시 사용 20편** ✓ (`measure_gen` 파싱으로 직접 재계수).
- **항공기 4건** ✓(space-use 2 + activity 2).
- **설계 레버 10개 · high 0 · moderate 4 · low 5 · very low 1** ✓ — `design_matrix_v2.csv` confidence 열과 정확히 일치하고, moderate 4개의 이름(L1 오디오 프로그래밍 · L2 자연음 · L3 정온경로 · L4 기계소음)도 일치.
- **L10 "quarter-point per decibel"** = CT0126 `LAeq 1 dB당 발성 증대 +0.25점(0~10 척도), R²=0.42` ✓.
- **§4.3 기타** "도착 인원엔 무효과"(323, high) ✓ · "마스킹→행태 검정 0편" ✓ · "가장 큰 데이터셋 2건이 소음-활동 정적 연관"(1226 n=4,640 · 951 13,322구간) ✓ · §4.5 "최장 한 시즌"(566) ✓ · "헤드폰 3/4" ✓.
- **§2.3 벤치마크 5편 전부 HIT** ✓ (`search_strings_20260802_205110.md` "벤치마크 리콜 5/5").
- **정본 버전 사용 여부** — 구버전 4종(`evidence_map.md` · `table1_study_characteristics.md` · `quality_summary.md` · `ma/ma_summary.md`, 전부 81편 기준·첫 줄 표식)의 고유값(등급 20/34/30, 84편 등)은 원고에 **나타나지 않음**. 다만 §4.5-5의 "only two studies"는 v2 이전 상태를 서술한 잔존값이다(위 표 참조).
