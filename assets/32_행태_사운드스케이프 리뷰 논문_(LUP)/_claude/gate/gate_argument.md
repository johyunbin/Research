# Gate — 논증 타당성 (독립 검증)

대상: `01_논문작업/Manuscript_EN_20260806_draft.md` (2026-08-06 draft, 96편)
축: 근거가 결론을 지탱하는가 · 관점: Landscape and Urban Planning 심사자
대조 근거: `ma/ma_v2_summary.md` · `ma/ma_sensitivity_v2.md` · `ma/ma_*_input.csv` · `design_matrix_v2.csv` ·
`design_implications_v2.md` · `quality_v2.csv` · `source_flags.md` · `deviation_log.md` ·
`corpus_v3_extraction.csv` · `evidence_map_v2.md` · `prisma_flow.md` · `analysis_rules.md`

---

## 판정

**REVISE** — 자료층(추출·품질평가·민감도·이탈기록)은 이 규모의 리뷰로는 이례적으로 견고하나, **원고가 헤드라인으로 내세운 두 주장("50년 간격의 독립 반복"·"MA3는 소음 제거에 귀속")이 각각 사실오류와 자기 입력자료에 의한 반증에 걸려 있어**, 현 상태로 투고하면 심사자가 두 곳만 대조해도 리뷰 전체의 신뢰가 무너진다.

> ⚠️ 현 상태 그대로면 실질 판정은 reject에 가깝다. REVISE로 두는 이유는 **결함이 전부 서술층에 있고 재분석 없이 교정 가능**하기 때문이다. 아래 C1~C9는 교정 전 투고 금지 항목이다.

---

## 치명적 결함

| 위치 | 문제 | 왜 문제인가 | 제안 |
|---|---|---|---|
| **C1** Abstract L22 · §3.4 L244 · §4.3 L351 · §5 L431 | **"fifty years"가 사실오류.** 원고가 지목한 두 현장실험은 Mathews & Canon **1975**(CT0025, USA)와 Moser **1988**(uid 14, France) — **간격은 13년**이다. `corpus_v3_extraction.csv`에서 uid 14 = year 1988·France, CT0025 = 1975·USA로 확인됨. 4회 반복된다. | 이 리뷰의 **유일한 유의 SMD 결과에 붙은 유일한 강점 서술**이 검증 1분 만에 무너지는 오류다. 심사자는 참고문헌 연도 두 개만 보면 안다. "50년"은 `deviation_log.md` D5-6에서 발원한 내부 오류가 그대로 승계된 것으로, 원고→보충자료 전체에 퍼져 있다. | 전 4개소를 "thirteen years apart" 또는 "two independent field experiments (1975, 1988)"로 교정. `deviation_log.md` D5-6도 동시 수정(보충자료로 제출되므로). "50년"을 쓰려면 대상은 1975→현재이며 그건 반복이 아니라 **경과 시간**임을 명시. |
| **C2** Abstract L21–22 · §4.2 L325–332 · §4.3 L347–351 · §5 L430–431 | **"MA3는 자연음 도입이 아니라 소음 제거에 귀속" 주장이 원고 자신의 입력값에 의해 반증된다.** `ma_social_input.csv`+`ma_v2_new_inputs.csv`의 4개 효과: CT0025(기계음) g=**+1.073**, 931(자연음) g=**+0.977**, 1069(자연음) g=**+0.547**, 14(기계음) g=**+0.432**. 기계음 쌍 평균 ≈0.75, 자연음 쌍 평균 ≈0.76 — **사실상 동일**. 클러스터 2위 효과가 자연음 연구(931)다. | 원고는 "these include the largest effect in the cluster (d = +1.07)"만 제시하는데, 같은 기계음 쌍이 **최소 효과(0.43)도** 포함한다는 사실을 누락한다. 체리피킹이다. 게다가 `ma_sensitivity_v2.md` LOO에서 **기계음 연구 14를 빼면 추정치가 +0.646→+0.767로 올라간다** — 기계음 하위군이 결과를 끌어올린다는 서술과 정면 충돌. 이 오귀속이 §4.3 "Act on operational noise first"와 §5 결론의 근거 전부다. | ① 4개 효과의 g를 Results 본문 또는 Table 2 각주에 전부 노출. ② 주장을 원자료 수준으로 후퇴: `design_implications_v2.md` L2가 실제로 지지하는 것은 "**유의화(p .053→.021)를 만든 신규 1편이 소음제거 설계였다**"는 *증분* 진술이지, "결과가 소음제거에 귀속된다"는 *총량* 진술이 아니다. ③ 아래 C2-b를 함께 반영. |
| **C2-b** §4.2 L325–332 | **애초에 식별 불가능한 대비를 이분법으로 귀속했다.** 4개 중 931의 대비는 "bird+water **vs** traffic+construction"으로 자연음 추가와 소음 제거가 분리되지 않는다(`design_implications_v2.md` L2 표: "L2 + L4 혼재 — 자연음 추가와 소음 제거가 분리되지 않음"). 1069만 SPL 등화로 음종 효과를 격리한다. | 즉 "소음 제거 vs 자연음 추가"는 이 4효과로는 **식별되지 않는다**. 근거 부족(unsupported)이 아니라 **설계상 판별 불가(unidentifiable)**이며, 후자가 훨씬 무겁다. | "두 경로는 이 코퍼스에서 분리되지 않는다. 순수 자연음-대-정온 대비는 0건이다"로 재서술. 이게 오히려 더 강한 연구어젠다 논지가 된다. |
| **C2-c** §4.2 L331 | **k=2 이중잣대.** 자연음 하위군을 "The natural-sound effects are two, which under this review's own rules is below the threshold for an interpretable pooled interval"로 기각하면서, **동일하게 k=2인 기계음 하위군**으로부터는 긍정 결론을 끌어낸다. | 자기규칙의 선택적 적용. 심사자가 가장 좋아하는 지적 유형이다. | 같은 문장에서 "machinery effects are also two" 명시 → 어느 쪽도 하위군 결론을 지지하지 않는다는 결론으로 통일. |
| **C3** §3.4 L250–252 | **MA3 민감도의 선택적 보고.** 원고는 CT0025 제거만 보고("Removing it returns the cluster to *k* = 3 and *p* = .053")한다. 실제 `ma_sensitivity_v2.md` LOO: −931 → **p=.057**, −1069 → **p=.071**, −CT0025 → **p=.053**, −14 → p=.044. 또한 **high-only(k=2) → p=.074**. | **4편 중 3편 중 어느 하나를 빼도 유의성이 사라진다.** 원고는 이를 "인용추적 갈래가 형식이 아니었다는 증거"라는 유리한 프레임 하나로만 제시해, 독자가 "그 1편만 빼면 되는" 국소적 취약성으로 오해한다. 고품질만 남기면 유의하지 않다는 사실도 누락. 이 리뷰의 유일한 유의 SMD에 대한 강건성 표상을 왜곡한다. | Results에 LOO 4행 전부(또는 "all but one leave-one-out drops below α") + high-only 결과 명시. Limitations에도 이월. 프레임을 "인용추적 성과"에서 "k=4 풀의 구조적 취약성"으로 전환. |
| **C4** §2.6 L162–163 · Table 2 L234 · §3.4 L258–263 | **MA1의 k=4는 3편에서 나왔고, 원고가 §2.6에 선언한 규칙을 위반한다.** `ma_walking_input.csv`: 461(Franek 2018), **532 Exp1**, **532 Exp2**(둘 다 Franek 2019·동일 대비 "birdsong vs crowded city noise"), 617. `design_matrix_v2.csv` L9가 "MA1 k=4 (**3편**·532는 2개 실험)"로 명시하고 `deviation_log.md` D3-5도 "461·532·617" 3편으로 적는다. 원고 §2.6은 "Where a study contributed more than one estimate to the same cluster **and contrast frame**, one effect was selected by a fixed priority"라고 선언했고 `analysis_rules.md` §3도 "클러스터×대비프레임당 논문 1효과 원칙"이다. | 예외 사유가 원고·이탈기록 어디에도 없다. 심사자가 보충자료 입력표를 열면 **선언한 규칙을 지키지 않은 것**으로 읽는다(설령 두 실험의 표본이 독립이어도, 규칙은 무조건문이고 τ²·HK 구간은 4개 독립 단위를 전제한다). 게다가 §3.4·§4.2가 "four effects"를 반복해 폭을 실제보다 넓게 보이게 한다. | ① Table 2에 "*k* = 4 effects from 3 publications" 표기. ② §2.6 규칙에 명시적 예외(별도 표본의 별개 실험은 독립 추정치로 취급) 추가 + `deviation_log.md`에 항목 신설. ③ 또는 532를 논문 내 평균으로 합성해 k=3으로 재계산(규약 §3의 원래 처리). |
| **C5** §4.2 L338–341 | **"one research programme" 문단이 풀 안팎을 뒤섞으며 사실오류를 낳는다.** (a) "the pooled studies and the two new high-quality walking papers **are all Czech**" — 풀 기여 617은 **Algeria**(corpus: country=Algeria, design=field-**observation**)다. (b) "the only study that altered the environment itself is rated low and reports a difference of **0.06 m/s**" — 이 0.06 m/s는 `design_matrix_v2.csv` L9의 **941(1.24→1.18 m/s)**이고 **941은 MA1 풀에 없다**. (c) §3.4는 "**Three** high-quality walking studies exist"라 하고 §4.2는 "**the two** new high-quality walking papers"라 한다(세 번째 CT0335는 **Germany**). | 심사자가 지적한 우려가 그대로 현실화됐다: 461·532는 풀 안, CT0090·CT0166은 풀 밖인데 문장이 이를 뭉갠다. 결과적으로 **풀의 4번째 효과(617)가 무엇인지 독자가 알 수 없게 되고**, 617이 (i)관찰연구이며 (ii)나머지 셋과 **부호가 반대**(g=+0.227 vs −1.024/−0.359/−0.896)라는 결정적 사실이 은폐된다. | 정확히 재서술: "풀 4효과 중 3효과가 동일 연구팀(Franek 2018·2019)의 동일 1.8 km 루트 헤드폰 실험이고, 나머지 1효과(617)는 알제리 오아시스의 **관찰** 연구로 부호가 반대다. 풀 밖 고품질 3편 중 2편은 같은 체코 프로그램, 1편은 독일이다." **어느 pooled 효과도 공간의 음환경을 조작하지 않았다**는 것이 정확한 진술이며, 현 서술보다 강하다. |
| **C6** Abstract L24 · §1.3 L48 · §3.5 L267–268 · §4.1 L299 · §5 L437 | **"40%"의 분모·단위가 원고 안에서 일관되지 않다.** `evidence_map_v2.md` §2: forward 110 · reverse 74 · both 22 = **206**. 74/206 = **36%**, (74+22)/206 = 47%. 40%는 74/(110+74)=40.2%, 즉 **both를 분모에서 뺀 값**이다. 그런데 §3.5는 세 수를 나란히 제시한 직후 "40%"라 쓴다. 또 결론 L437은 **"Nearly half"**로 격상된다. | 심사자가 암산 한 번으로 74/206=36%를 얻는다. 같은 지표가 40%(§1.3·§3.5·§4.1)와 "nearly half"(§3.5 L276·§5 L437)로 오간다. **단위 문제는 더 심각**하다 — 이 40%는 논문 수가 아니라 **도메인 레코드 수**인데, Abstract는 "Forty percent of the evidence runs in reverse"로 단위 없이 쓰고 §3.2는 "38 of 96 studies (**40%**) were conducted in China"로 **같은 숫자를 논문 단위로** 쓴다. 독자가 "96편 중 40%가 역방향 연구"로 읽을 여지가 크다. | ① 분모를 하나로 고정하고 명시: "74 of 206 domain-level records (36%) run in reverse; counting bidirectional records, 96 of 206 (47%) involve the reverse pathway." ② Abstract·결론에 **단위 삽입**: "of domain-level records (not studies)". ③ "nearly half"는 both 포함 정의를 쓸 때만 사용하고 §1.3·§4.1과 정의를 통일. |
| **C7** §4.3 L353–355 · §5 L442 | **귀무 결과에서 처방을 도출한다.** MA2는 g=+0.313, 95% CI **−0.076~+0.702**, p=.074(원고 Abstract도 "dwell time (*g* = +0.31) did not [exclude zero]"로 인정). 그런데 §4.3은 단정형으로 "**Music extends the stay of people already present**"라 쓰고 §5는 "Use added sound to extend stay"로 실무 지침화한다. 또한 MA2 3효과 중 1280은 음악이 아니라 **자연음지수 로지스틱 OR**(`ma_staying_input.csv`)이라 "Music"으로 귀속할 수 없다(C2와 동형 오류). | Abstract·Results ↔ Discussion·Conclusions 간 **직접적 자기모순**. 리뷰가 스스로 "0을 포함한다"고 밝힌 추정치를 두 절 뒤에서 확정 사실로 처방한다. LUP 심사자가 가장 확실하게 잡는 유형. | "Direction is consistent across three studies but the pooled interval includes zero; treat added sound as a **hypothesis to test with monitoring**, not an established dwell-time instrument." 대비 구성(음악 2 + 자연음지수 1)도 명시. |
| **C8** §2.6 · §4.6 (부재) | **출판편향을 원고 전체에서 단 한 번도 언급하지 않는다**(전문 검색 결과 "publication bias/funnel/Egger/small-study" 0건). `analysis_rules.md` §6은 "출판편향: k≥10 클러스터만 funnel+Egger, **그 외 서술**"로 사전 규정했는데, 그 **서술조차 수행되지 않았고** `deviation_log.md`에도 항목이 없다. | PRISMA 2020 item 13e/14/22 미이행. 사전규약 미실행 + 이탈 미기록의 이중 결함이다. "모두 예측 방향을 향한다"는 원고의 반복 서술은 소규모연구효과가 배제됐을 때만 강점이 되므로, 미검정 상태에서 그 서술은 근거가 없다. | Methods에 "k<10이라 funnel/Egger는 시행하지 않았다" 1문장 + Limitations에 소규모연구효과 미배제를 명시. 4클러스터 전부 k≤7이므로 정직한 진술이 곧 충분한 대응이다. |
| **C9** §2.6 · §4.1 L300–301 | **사전등록된 하위군(moderator) 분석이 수행되지도, 이탈로 기록되지도 않았다.** `analysis_rules.md` §6: "하위그룹: 세팅(공원/가로/광장)·설계(현장실험/관찰)·측정세대". `ma_sensitivity_v2.md`에 하위군 분석은 없고 `deviation_log.md`에도 해당 항목이 없다. 그런데 §4.1은 "Between them sit moderators that **determine** whether a given sound produces approach or avoidance: setting type, visual–acoustic congruence, purpose of stay, and cultural context"라고 단정한다. | **검정하지 않은 조절변수를 프레임워크의 인과 구성요소로 선언**했다. 네 조절변수 중 어느 것도 이 리뷰에서 검정된 바 없다(그중 둘은 코딩조차 되지 않았다). Figure 6의 신뢰도 전체가 걸린다. | ① 하위군 미실행을 이탈로 기록 + Methods 1문장(k가 작아 하위군 풀링 불가). ② "determine"을 "are proposed as candidate moderators, none tested in this review"로 격하. Figure 6 캡션에도 동일 표기. |

---

## 과장·인과 오용

### 1) 인과동사 — 관찰근거에 붙은 것

**(a) §4.3 L368–370** — 가장 문제되는 곳
> "And noise reduction as a route to physical activity is **contradicted** rather than merely unsupported: two of the largest datasets here find noise positively associated with activity, **because activity concentrates where cities are loud**."

- 근거: 951(Finland, Strava 13,322 street segments)·1226(Bulgaria, n=4,640). 둘 다 **횡단 생태 상관**이며, `design_matrix_v2.csv` L5 caveat는 "1226·951은 부호가 반대(**접근성 교락**)"로 적는다. 951 원문은 blue space density가 지배적 예측인자라고 보고한다.
- 두 겹의 오용: ① 교락된 상관 2건은 개입 가설을 "contradict"할 수 없다(부재증거·역인과·생태오류). ② "because activity concentrates where cities are loud"는 **검정되지 않은 인과 설명을 사실로 단정**한 것이며, 원고는 이 교락을 오히려 논거로 승격시킨다.
- 대안: "Two of the largest datasets report noise **positively** associated with activity, most plausibly because activity concentrates in dense, loud areas — a confounding pattern that these cross-sectional designs cannot resolve. The corpus therefore provides **no behavioural test** of noise reduction as a route to physical activity, and the observational sign should not be read either way."

**(b) §3.4 L239–241**
> "All four estimates point in the theoretically predicted direction: **noise accelerates passage, positive sound extends stay,** quiet and natural sound **increase** social interaction, and appraisal covaries with behaviour at moderate strength."

- 네 인과 서술 중 둘(보행속도 p=.179 · 체류 p=.074)은 CI가 0을 포함한다. 바로 다음 문장의 굵은 교정("The clusters differ sharply, however...")이 있으나, **인용될 문장은 앞 문장**이다. 게다가 MA1 내부에서는 4효과 중 617이 부호가 반대다.
- 대안: 인과동사를 방향 서술로 교체 — "point toward faster passage, longer stay, and more social interaction respectively", 그리고 "one of the four walking effects runs in the opposite direction" 삽입.

**(c) §5 L425**
> "**Sound changes what people do** in urban open space, but the evidence for that claim is unevenly distributed..."

- 결론 첫 문장이 무조건 인과문이다. 96편 중 현장실험은 17편(18%)이고, 유의한 SMD 클러스터는 1개(k=4·3편 취약)뿐이다.
- 대안: "The acoustic environment is associated with what people do in urban open space, and in a small number of field experiments it demonstrably changes it — but..."

**(d) §1.3 L48–50 / §4.1 L298–299** — "activity ... *generate* the acoustic environment"
- 역방향 74건 대부분이 관찰·상관이다(밀도–SPL 회귀 등). "generate"는 대체로 정당(물리적 음원 발생)하나, "companionship altering what is noticed"(§3.5 L275)는 **지각** 결과이지 음환경 생성이 아니다. CT0184의 note도 "역방향(행태→지각)"으로 적는다. 역방향 arc를 "sound production"으로 정의한 §4.1과 충돌.
- 대안: 역방향을 두 갈래(sound **production** vs sound **attention/appraisal**)로 나누어 서술하고, 74건의 구성비를 밝힐 것.

### 2) 증거강도에 부합하지 않는 수식어

**(e) §3.4 L244–248 / Abstract L22**
> "An independent replication at this interval, using an unrelated manipulation and population, is **stronger evidence than *k* = 4 suggests**."

- k가 과소평가라는 주장인데, C3(4편 중 3편 어느 하나 제거 시 비유의)·C1(13년)·C4류의 구조 취약성을 감안하면 **정반대**로 읽힌다. 또 Moser(14)의 g=+0.432가 풀 평균 이하라(제거 시 추정치 상승) "replication"의 강도 자체가 원고 서술만큼 크지 않다.
- 대안: "Two field experiments thirteen years apart, in different countries and with different noise sources, agree in sign — but the pooled interval is fragile: every leave-one-out except one drops below α." Moser의 효과크기를 반드시 병기할 것.

**(f) §5 L430–431**
> "The **best-established** behavioural consequence of the acoustic environment is **social**: machinery noise suppresses interaction between strangers"

- "best-established"는 상대 서열 주장인데, 근거는 (i)k=4·3편 취약(C3) (ii)기계음 귀속 반증(C2) (iii)high-only 비유의(p=.074). "established"에 값하지 않는다. 아울러 Mathews & Canon의 아웃컴은 **helping behaviour**(떨어뜨린 물건 줍기 도움)이며 "interaction between strangers"로의 일반화는 한 단계 확장이다.
- 대안: "The only cluster whose interval excludes zero for a manipulated contrast is social interaction, where two field experiments show that machinery noise reduces prosocial responding to strangers."

**(g) §4.3 L358–361 (레버 L10)**
> "The evidence supports a continuous gradient rather than a cut-off: **roughly a quarter-point increase in self-reported voice-raising per decibel**. That is enough to compare candidate locations, not enough to set a standard."

- 단일 연구(CT0126). `design_matrix_v2.csv` L10 + `ma_v2_new_inputs.csv`에 따르면 **분석단위 N=29 sampling points**(105명이 아님), 자기보고, 스페인 단일 표본, 그리고 **CT0322와 동일 데이터셋**(중복으로 다른 곳에서는 배제한 그 쌍)이다. L10 caveat는 "**관찰된 대화 행동으로 검증되지 않았다**"고 적는다.
- 결정적으로 원고는 **척도를 밝히지 않았다** — 원자료는 "0~10 척도에서 dB당 +0.25점"이다. 척도 없는 "quarter-point per decibel"은 실무자가 쓸 수 없고, "enough to compare candidate locations"라는 주장도 검증 불가다.
- 대안: 척도·n·자기보고·단일 표본을 한 문장에 넣고 "candidate locations 비교"라는 실무 주장은 삭제하거나 "hypothesis to test"로 격하.

**(h) §4.2 L318–321 · Abstract L27–28**
> "**No lever reaches high confidence.** ... this field can currently support *hypotheses to test with monitoring*, not design standards."

- ✅ 이건 과장이 아니라 **정직한 진술이며 원고의 최대 강점 중 하나**다(유지). 다만 §4.3의 개별 처방들(C7·(g))이 이 원칙을 지키지 않아 자기모순을 만든다 — §4.3을 §4.2의 기준에 맞춰 내리는 것이 교정 방향이다.

### 3) 검색 완결성 과장

**(i) Abstract L14–15 / §2.3 L110–111**
> "backward and forward citation tracking was performed on **every included study**"

- `prisma_flow.md`: "인용 추적 시드 = **84편**(FINAL_INCLUDE 81 + SENS_ONLY 3)". 즉 **DB 갈래 포함분만** 시드였고, 인용추적으로 새로 들어온 15편에는 추적이 **반복 수행되지 않았다**(단일 이터레이션).
- Abstract는 최종 코퍼스 96편을 제시한 뒤 "every included study"라고 쓰므로 독자는 96편 전부라고 읽는다. 검색 완결성 주장이라 심사자가 반드시 확인한다.
- 대안: "citation tracking was performed on the 84 studies identified by database searching (one iteration; the 15 studies it added were not re-tracked)."

---

## 논리적 비약 — 지정된 세 주장

### 주장 1. "MA3 유의화는 자연음 도입이 아니라 소음 제거에 귀속된다"

**판정: 부당 — 기각.** 세 겹으로 실패한다.

1. **최대효과 논증은 비논리(non sequitur)다.** 4개 중 하나가 가장 크다는 사실은 그 하위군이 유의성을 만들었음을 함의하지 않는다. 실제 값: 기계음 {+1.073, +0.432}, 자연음 {+0.977, +0.547} — **두 쌍의 평균이 사실상 같고**, 2위 효과가 자연음(931)이다.
2. **민감도가 반증한다.** `ma_sensitivity_v2.md` LOO에서 **기계음 연구 14를 빼면 추정치가 +0.646 → +0.767로 상승**한다. 기계음 하위군이 풀을 끌어올린다는 서술과 모순이며, "기계음"이라는 *범주*가 아니라 **CT0025 한 편**이 클 뿐임을 보여준다. 한 편을 하위군으로 승격시킨 것이 이 주장의 실체다.
3. **애초에 식별되지 않는다.** 931의 대비는 자연음 추가와 소음 제거가 뒤섞여 있고(원자료 스스로 "L2+L4 혼재"로 명시), 순수한 "자연음 vs 정온" 대비는 풀에 **0건**이다. 노출 2 vs 2라는 산술적 대칭조차 성립하지 않는다 — 실질은 기계음 2 · 혼재 1 · 등화대비 1이다.

또한 **원자료가 지지하는 명제는 원고의 명제보다 훨씬 좁다**. `design_implications_v2.md` L2는 "MA3 **유의화**는 자연음이 아니라 소음 제거 쪽 **신규 연구**에서 왔다"(=증분 진술)로 쓴다. 원고는 이를 "The social-interaction result **is attributable to** removal of machinery noise rather than addition of natural sound"(=총량 진술)로 격상했다. **증분→총량 격상이 이 비약의 정확한 발생 지점**이다.

> 살릴 수 있는 형태: "The cluster's shift to significance was produced by a newly added noise-removal experiment; the natural-sound and noise-removal contributions are of similar magnitude and cannot be separated, because the corpus contains no natural-sound-versus-quiet contrast. Prescribing a fountain from this pool would be a category error — **and so would prescribing noise removal from it.**"

### 주장 2. "보행속도 근거는 사실상 하나의 연구 프로그램"

**판정: 결론은 옳으나 근거 제시가 부정확하고, 심사자가 지적한 대로 풀 안팎이 흐려졌다 — 재서술 필수.**

- **옳은 부분**: 461(Franek 2018)·532(Franek 2019)는 같은 체코 연구팀·같은 1.8 km 루트이며, 풀 밖 CT0090(2014)·CT0166(2021)도 흐라데츠크랄로베 1.75–1.8 km 순환로다. "one research programme"이라는 진단 자체는 자료가 지지한다. 오히려 **원고보다 강하게 말할 수 있다 — 풀 4효과 중 3효과가 Franek 2편에서 나온다**(C4).
- **틀린 부분 ①**: "the pooled studies ... are all Czech"는 **거짓**이다. 617은 알제리 오아시스 **관찰**연구다(corpus: Algeria / field-observation).
- **틀린 부분 ②**: "the only study that altered the environment itself ... reports a difference of 0.06 m/s"의 0.06 m/s는 **941**(1.24→1.18 m/s)이고 941은 **풀에 없다**. 풀의 4번째는 617이며 아무것도 조작하지 않았다.
- **흐려진 구분**: 461·532 = 풀 안(모두 헤드폰) / CT0090·CT0166 = 풀 밖(고품질) / CT0335 = 풀 밖·**독일** / 941 = 풀 밖·현장조작 / 617 = 풀 안·**비체코·관찰**. 현 문장은 이 다섯 층위를 한 덩어리로 뭉쳐 "체코 단일 프로그램"이라는 인상을 만든다. **결과적으로 617이 무엇인지 독자가 알 수 없다**(부호 반대·관찰·k의 4분의 1).
- **숫자 불일치**: §3.4 "**Three** high-quality walking studies exist" ↔ §4.2 "**the two** new high-quality walking papers".

> 재서술 골자: "*k* = 4 rests on three publications. Three of the four effects come from one Czech group's two studies on the same 1.75–1.8 km circuit, all delivering sound **through headphones**; the fourth (Algeria) is an **observational** study whose effect runs in the opposite direction. Not one pooled effect manipulated the acoustic environment of the space itself. The two non-pooled high-quality studies come from the same Czech circuit; a third is German."

### 주장 3. "40%가 역방향"

**판정: 산술·단위 모두 부정확 — 오독 여지가 크다.**

- **산술**: `evidence_map_v2.md` §2 = forward 110 · reverse 74 · both 22 (합 206). **74/206 = 36%**. 40%는 both를 분모에서 제외한 74/184다. §3.5는 세 수를 제시한 **직후** 40%를 쓰므로 독자가 그 자리에서 검산하면 어긋난다. 원자료(`evidence_map_v2.md`)도 같은 근사를 쓰지만, 세 수를 병기하지 않아 원고만큼 노출되지 않는다.
- **단위**: 이 40%는 **도메인 레코드**(1편이 여러 도메인에 기여) 단위다. 원고는 §1.3·§3.5에서 "domain-level records"를 명시하지만, **Abstract L24 "Forty percent of the evidence runs in reverse"**와 **§4.1 L299 "accounts for 40% of what has actually been measured"**, **§5 L437 "Nearly half the evidence runs in the reverse direction"**에는 단위가 없다.
- **오독 여지 — 높다.** ① Abstract만 읽는 독자(대다수)는 단위를 볼 수 없다. ② 같은 원고 §3.2가 "**38 of 96 studies (40%)** were conducted in China"로 **동일한 40%를 논문 단위로** 쓴다. 두 개의 40%가 한 원고에 있고 하나만 단위가 붙어 있다 → "96편 중 40%가 역방향 연구"라는 오독은 거의 유도된다.
- **추가 비일관**: 40%(§1.3·§4.1) ↔ "nearly half"(§3.5 L276·§5 L437). 후자는 both를 포함한 47% 정의라야 성립하는데 정의 전환이 선언되지 않는다. 결론이 본문보다 강하게 말하는 전형적 패턴.

> 교정 골자: 분모를 206으로 통일하고 "36% of domain-level records (74/206) run in reverse, and 47% (96/206) involve the reverse pathway when bidirectional records are counted"로 한 번에 정의. Abstract·결론에 "domain-level records, not studies" 삽입. "nearly half"는 47% 정의를 쓴 자리에서만.

---

## 자기모순 (위에 포함되지 않은 것)

| # | 위치 | 모순 |
|---|---|---|
| M1 | Abstract L20 ↔ §4.3 L353 ↔ §5 L442 | "dwell time (*g* = +0.31) did **not** [exclude zero]" ↔ "Music **extends** the stay" ↔ "**Use** added sound to extend stay" (=C7) |
| M2 | §3.7 L289–291 | "**Every** behavioural domain × sound source cell in the evidence map is populated (Fig. 3), **with one exception**: aircraft noise appears in four records only." — `evidence_map_v2.md` §1에서 항공기소음 열은 이동 **0**·체류 **0**·사회 **0**·공간이용 2·활동 2다. **5칸 중 3칸이 빈칸**이므로 "every cell populated"도 "one exception"도 성립하지 않는다. Fig. 3과 본문이 어긋나면 심사자는 표를 먼저 믿는다. → "All cells are populated except the aircraft-noise column, which is empty for movement, staying and social behaviour and holds only four records in total." |
| M3 | §3.1 L191–195 ↔ `prisma_flow.md` | "converted at 46% (**17** of 37)" — PRISMA 흐름은 전문평가 54 → 배제 38 → 민감도 1 → 포함 15, 즉 생존 **16**(15+1)이다. A군 17 + B군 0 = 17이 되어 **1건이 맞지 않는다**. 흐름도 검산은 심사자·편집자가 반드시 한다. 원자료(`prisma_flow.md` A/B군 표)부터 재대조 필요. |
| M4 | §1.2 L42 ↔ §3.2 L199 | "**Two-thirds** of the eligible literature appeared **after 2020** (68 of 96)" ↔ "Sixty-eight of 96 studies (**71%**) appeared **in 2020 or later**". 같은 68/96에 서로 다른 분수 표현과 서로 다른 기준연도 정의. → 둘 다 "71% in 2020 or later"로 통일. |
| M5 | §3.1 L177 ↔ L183 | 데이터베이스 검색 총계와 인용추적 신규 후보가 **둘 다 2,073**이다. `prisma_flow.md`는 이를 "**숫자 우연 주의** — 복사 오류가 아니라 우연이며 각각 독립 산출"이라고 **명시적으로 경고**하는데, **원고에는 그 각주가 없다**. 편집자는 복붙 오류로 읽고 되돌려 보낼 확률이 높다. → Fig. 1 캡션 또는 각주에 "coincidentally identical; derived independently" 1줄. |
| M6 | 원고 ↔ 보충자료 S1 | `deviation_log.md` **D1-4**가 "현재 원고의 포함 81편은 DB 검색 갈래만 반영한 수치다 · PRISMA 흐름도에 인용추적 갈래를 'reports sought for retrieval'에서 멈춘 상태로 점선 표기했다 · 79건 중 자동 회수 8건"으로 남아 있다. 최종 원고(96편·CT 54편 확보·15편 포함)와 **정면 충돌**한다. §2.1은 "All deviations ... are listed in Supplementary S1"이라 하므로 이 파일이 그대로 제출된다. → S1 제출 전 D1-4 갱신 필수(D2-3의 "분모 84"·"ID 7 undetermined"도 현 `quality_v2.csv`(ID 7 = low, 96편 22/40/34)와 불일치). |
| M7 | §4.2 L318 ↔ §4.3 순서 | "No lever reaches high confidence" + L2·L4가 **둘 다 moderate**(design_matrix)인데, §4.3은 "**Act on operational noise first**"로 L4를 L2 위에 서열화한다. Table 3은 그 서열을 지지하지 않는다(C2). → 서열 근거를 Table 3에서 도출하거나, "operational noise is the cheapest and most controllable of the moderate-confidence levers"처럼 **비용·통제가능성** 근거로 이동(증거강도 주장 회피). |
| M8 | §4.3 L358 | "(lever L10, **new in this revision**)" — 투고 원고에 남아서는 안 되는 내부 개정 이력 표현. 심사자에게는 존재하지 않는 이전 판본을 가리킨다. → 삭제. |

---

## 누락 — 반드시 추가해야 할 것

**Limitations(§4.6)에 반드시 들어가야 하나 현재 없는 것** (단일 평가자 스크리닝은 이미 기술됨 ✅):

1. **출판편향 미검정** (=C8). "모든 클러스터가 k ≤ 7이라 funnel plot·Egger 검정을 시행하지 않았고, 따라서 소규모연구효과를 배제할 수 없다." 사전규약에 있던 항목이므로 이탈기록에도 추가.
2. **MA1의 비독립성**(=C4): k=4가 3편에서, 그중 3효과가 한 연구팀에서.
3. **MA3의 LOO 취약성 전모**(=C3): 4편 중 3편 어느 하나 제거 시 비유의 · high-only(k=2) p=.074.
4. **효과크기 변환 경로의 가정** — 현재 §2.6은 변환식만 나열하고 **가정을 밝히지 않는다**. 최소 세 가지: ① **Chinn 변환**(*d* = ln(OR)×√3/π)은 잠재 연속변수의 로지스틱 분포를 가정하며, MA3의 4효과 중 2효과(14·CT0025)가 이 경로로 만들어져 **연속형 SMD와 같은 척도에서 풀링**된다. ② **복원 3건**(`deviation_log.md` D3-3): 532는 **군당 n 미보고 → 총 N 균등분할 가정**(MA1 4효과 중 2효과가 여기 해당), 665는 사후 p에서 t 역산, 323은 SE→SD 환산. 민감도 ⑤는 k=2로 떨어져 **HK 구간 해석 자체가 불가**(t(1)=12.7)하므로 이 가정은 실질적으로 검증되지 않았다. ③ Fisher-z 역변환된 r을 **Spearman rho 포함** 상태로 보고(민감도 ④는 결과 동일하나 규약 §1의 "rho는 r 근사" 가정 자체는 남는다).
5. **I² = 90.5%의 해석과 풀링 적절성** — 현재 §3.4는 "which is expected given that it mixes forward and reverse pathways and several behavioural outcomes"로 **설명만 하고 넘어간다**. 그러나 (i)이질성의 원인이 **구성개념 이질성**이라면 그 풀링값이 무엇의 추정치인지가 문제이고, (ii)**예측구간(prediction interval)이 없다** — I²=90.5%·k=7이면 예측구간은 거의 확실히 0을 포함한다. Abstract는 이 r=+0.41을 "0을 제외한 둘 중 하나"로 내세우므로 **예측구간 보고가 사실상 의무**다.
   - 구성 자체도 문제다: MA4 7효과에 **CT0126(LAeq ↔ 자기보고 대화방해 = 노출–지각)**, **CT0184(동반상태 → 말소리 인지 = 행태→지각)**, **980(체류시간 ↔ 회복지각)**이 섞여 있다. 클러스터 명칭은 "Soundscape perception ↔ behaviour"인데 **최소 1효과는 지각–행태 상관이 아니다**(CT0126은 물리 노출–자기보고). 명칭과 내용을 일치시키거나, 클러스터를 분해할 것.
6. **역방향 74건의 성격** — 대부분 관찰·상관인데 §4.1은 이를 인과 arc("sound production")로 그린다. "역방향 증거는 거의 전부 관찰이며 예측규칙을 제시한 연구는 2편"이라는 사실(§4.5 어젠다 5에는 있음)이 Limitations에는 없다.
7. **사전등록 클러스터 5→4 축소**(`deviation_log.md` D3-1, 방문/공간이용 클러스터 소실)이 본문에 없다. §2.6은 "Four clusters met the pooling threshold"라고만 쓴다. OSF를 대조하는 심사자가 반드시 묻는다 — Methods에 1문장.
8. **하위군/조절변수 미실행**(=C9).
9. **PRISMA 2020 item 15(확실성 평가)** — 자체 confidence grade는 GRADE도 CERQual도 아니다. §4.2는 산출 기준(study count·MMAT·CI·setting diversity·sensitivity)만 언급하고 **가중·역치 규칙이 없다**. 각 등급의 결정규칙을 Supplementary에 명시하지 않으면 Table 3 전체가 "저자 의견"으로 읽힌다.

**LUP 적합성 관점의 누락(so what for planning?)**

- 처방 4개(operational noise / added sound / speech intelligibility / not-walking-speed-KPI)는 LUP 기준으로 **존재는 하나 공간적 구체성이 부족**하다. 원고에는 **거리·시간대·역치·배치 규칙이 하나도 없다** — 반면 `design_implications_v2.md`에는 쓸 수 있는 재료가 남아 있다(예: L1 "공공공간 감속을 만든 자극은 전부 100~110 bpm대", L6 밀도–SPL 예측식 3종 `LAeq = 33.74d + 67.12`·`60.75 + 7.41d`·음향쾌적 최적밀도 0.10~0.25인/m², L1 "323(high)은 멈추는 인원수에 무효과", L7 "광장무 1 m LAeq 87.5~104.4 dBA"). **밀도–SPL 예측식은 이 리뷰의 양방향 프레임을 실제 설계 도구로 바꾸는 유일한 정량 자산인데 본문에 전혀 등장하지 않는다.**
- 특히 §4.5 어젠다 5("only two studies give rules for predicting the sound a design will generate")는 **그 두 연구의 규칙을 제시하지 않는다**. LUP 심사자는 "역방향이 40%라면서 설계자가 쓸 예측식은 왜 안 주나"를 묻는다. §4.3에 밀도–SPL 관계를 **명시적 한계와 함께** 1문단으로 넣는 것이 이 원고의 planning 기여를 가장 크게 올리는 단일 수정이다.
- 반대로 **"더 연구가 필요하다"로 끝나지는 않는다** — §4.5가 6항 어젠다를 구체적으로 주고, "unsupported prescriptions"를 명시한다. 이 축에서는 통과.

---

## 강점 (유지)

1. **불리한 결과를 은폐하지 않는 구조** — MA1이 저품질 제외 시 k=0이 된다는 사실, MA3 유의성이 인용추적 1편에 걸린다는 사실, CT0414를 "넣으면 결과가 좋아지는데도" 뺀 판단(`deviation_log.md` D5-4)까지 공개한다. 이 정직성이 위 결함들에 대한 REVISE 판정을 가능하게 하는 유일한 이유다. C3처럼 **일부만 보고한 곳을 전부 보고로 바꾸면** 이 강점이 원고 전체로 확장된다.
2. **절단 인공물의 발견과 감사(§2.5·§4.4)** — 자동 전문 추출이 MMAT의 "can't tell" 판정을 통해 품질등급을 단방향으로 깎는다는 지적, 13편 재평가로 4편 등급 이동(3편은 2등급)을 실측한 것은 **이 논문 고유의 방법론 기여**이며 증거종합 커뮤니티에 실제 가치가 있다. 리뷰가 거절되더라도 이 부분은 살아남는다.
3. **양방향 프레임의 경험적 정당화** — 역방향이 공간이용·활동·사회적 행태에서 순방향과 대등하다는 도메인별 분해(25 vs 23 / 18 vs 21 / 19 vs 14)는 실제 자료에서 나왔고, "이동·체류는 순방향 지배"라는 비대칭까지 짚는다. 백분율 표기(C6)만 고치면 이것이 원고의 가장 독창적인 논지다.

---

## 교정 우선순위 (투고 전)

| 순위 | 항목 | 성격 |
|---|---|---|
| 1 | C1 "fifty years" → "thirteen years" (원고 4개소 + `deviation_log.md` D5-6) | 사실오류 |
| 2 | C2/C2-b/C2-c MA3 귀속 주장 철회·재서술 (Abstract·§4.2·§4.3·§5) | 반증된 주장 |
| 3 | C3 MA3 LOO 전모 + high-only 보고 | 선택적 보고 |
| 4 | C4·C5 MA1 = 4효과/3편 공개 + "all Czech"·"0.06 m/s" 교정 | 사실오류+규칙위반 |
| 5 | C7 MA2 귀무 결과로부터의 처방 철회 | 자기모순 |
| 6 | C6 40% 분모·단위 통일 (Abstract·§1.3·§3.5·§4.1·§5) | 산술+오독 |
| 7 | C8·C9 출판편향·하위군 미실행 명시 (Methods·Limitations·이탈기록) | PRISMA 미이행 |
| 8 | M2·M3·M5·M6 Fig 3 정합·PRISMA 17 vs 16·2,073 각주·S1 D1-4 갱신 | 검산 위험 |
| 9 | 누락 4·5(변환 가정·예측구간·MA4 구성) | 심사자 필문 |
| 10 | 밀도–SPL 예측식을 §4.3에 도입 | LUP 기여 상향 |
