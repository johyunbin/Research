# Paper32 — 독립 게이트 지적 사항 수정 목록 (2026-08-06)

3축 독립 게이트 판정: **수치 FAIL · 논증 REVISE · 보고 REVISE**.
아래는 내가 직접 재검산해 **사실로 확인한 것만** 옮긴 것이다. 확인 방법도 함께 적는다.

---

## A. 사실오류 — 즉시 수정 (내가 만든 오류)

### A1. ★ "fifty years apart" → **13년** [critical]
Mathews & Canon **1975** ↔ Moser **1988**. 간격은 13년이다.
- 확인: `table1_v2.csv`의 `uid=CT0025` year=1975, `uid=14` year=1988.
- 원 오류는 내가 `ma_v2.py`에 쓴 "50년 간격의 독립 반복" 문구이고, 그것이
  `ma_v2_summary.md` → `design_implications_v2.md` → 원고로 전파됐다.
- **원고 4개소**(Abstract · §3.4 · §4.3 · §5)와 `ma_v2_summary.md` · `deviation_log.md` D5-6 ·
  `writing_pack.md` · `design_implications_v2.md` 전부 수정.
- 함께 붙은 **"two countries"도 지지되지 않는다** — CT0025의 country가 `NR`(원문에 수집국 미명시).
  Moser는 France. "two countries"는 삭제하거나 "France and (probably) the United States"처럼
  불확실을 명시해야 한다.

### A2. ★ MA3 귀속 주장이 자기 입력값에 반증됨 [critical]
원고 §4.2의 *"the significant social-interaction result belongs to noise removal, not natural
sound addition"* 은 **성립하지 않는다.**
- 확인: 기여 4효과를 노출로 가르면 자연음 {931 +0.977, 1069 +0.547} 평균 **+0.762**,
  기계음 {14 +0.432, CT0025 +1.073} 평균 **+0.753**. 차이 **0.009**.
- 게다가 LOO에서 **기계음 14를 빼면 추정치가 +0.646 → +0.767로 오른다**(`ma_sensitivity_v2.csv`).
  기계음이 효과를 끌어올린다는 서술과 반대 방향이다.
- 결정적으로 **931·1069의 대비 자체가 "자연음 추가 + 소음 제거" 혼재**다(bird/water vs
  traffic/construction). 순수한 자연음-대-정온 대비는 코퍼스에 **0건**이라 애초에 식별 불가능하다.
- **지탱되는 것은 이것뿐**: 유의 문턱을 넘긴 *증분*이 기계음 연구 1편(CT0025)에서 왔다.
  총량 귀속이 아니라 증분 귀속으로 낮춰 써야 한다.
- 파급: `design_implications_v2.md`의 L4 "MA3 유의화는 이 레버로 귀속된다"도 같은 수정 필요.

### A3. MMAT 문항 백분율 오기 [major]
원고 §3.3: "non-response bias in 41%, confounding in 37%".
- 재산출(`quality_detail_v2.csv`, FINAL_INCLUDE 96편 한정):
  **4.2 표본대표성 10/51 = 19.6%** · **4.4 무응답편의 16/51 = 31.4%** ·
  **3.4 교란보정 9/22 = 40.9%** · 3.1 참가자대표성 4/22 = 18.2%.
- 두 값이 뒤바뀐 뒤 한쪽이 더 틀어졌다.

### A4. "the two randomised studies score lowest" 거짓 [major]
- RCT 2편(461·532)의 `n_yes`는 2와 1인데, **`n_yes=0`인 연구가 4편**(1122·123·7·791) 있다.
- 맞는 서술: RCT 2편이 **무작위화·기저동등성·눈가림을 전건 보고하지 않아 설계 이점을 전혀
  인정받지 못했다**(2.1·2.2·2.4 모두 CT). "최하위"가 아니라 "설계 대비 최저"다.

### A5. "the pooled studies … are all Czech" 거짓 [major]
- MA1 풀 4효과 중 **617은 Algeria**이고 관찰연구이며 부호도 반대(+0.227)다.
- 맞는 서술: 4효과 중 **3효과가 체코 Franěk 연구 2편**에서 나온다. 여기에 신규 고품질
  CT0090·CT0166도 체코 동일 루트다. 즉 "풀이 전부 체코"가 아니라 "**풀의 3/4과 신규 고품질
  보행연구가 한 연구 프로그램**"이다. 이렇게 쓰면 오히려 더 강한 진술이다.

### A6. "0.06 m/s"의 출처 명시 누락 [major]
- 이 값은 **uid 941**의 것인데 941은 **MA1 풀 밖**이다. 원고가 풀 안 이야기처럼 읽히게 썼다.
- "the only study that altered the environment itself"라는 수식도 풀 밖 연구를 가리킨다는 것을
  명시해야 한다.

### A7. 역방향 40%의 분모 미표기 + 단위 혼동 [major]
- 재산출: forward 110 · reverse 74 · both 22, 합 **206**.
  - reverse/전체 = **35.9%**
  - reverse/(forward+reverse) = 40.2% ← 원고의 40%는 이것(both를 분모에서 제외)
  - (reverse+both)/전체 = **46.6%** ← "nearly half"의 근거
- 원고는 §3.5에서 세 수를 나열한 직후 "40%"라 써서 독자가 74/206으로 계산하면 안 맞는다.
- 더 나쁜 것: **§3.2에서 "38 of 96 studies (40%)"(China)를 쓴다.** 같은 40%가 한 번은 논문 단위,
  한 번은 도메인 레코드 단위다. 오독이 사실상 유도된다.
- 수정: 분모를 명시하고(`74 of 206 domain-level records, 36%`), "nearly half"는 46.6% 근거로
  분리하거나 삭제. Abstract·§1.3·§4.1·§5의 "40%"도 전부 단위 표기.

### A8. A군 전환율 46% → **43%** [major]
- REC 410(중국어 게재) 배제 보정 **이전** 값이다. 정본 `ct_verdicts_final.csv` 기준 **16/37 = 43.2%**.
- `prisma_flow.md` · `writing_pack.md` · 원고 §3.1 전부 수정.

### A9. Table 1 코딩 오류 — CT0025 setting [minor→데이터]
- `setting = lab(outdoor scene)`인데 `design = field experiment`다. 메타분석에 쓴 것은 **Exp2(가로에서
  잔디깎기 가동)** 이므로 setting은 **street**이 맞다. 내 `pick()` 휴리스틱이 본문의 "laboratory"
  (Exp1 언급)에 걸린 것.
- 수정 후 setting 집계·Fig 7·writing_pack 재생성 필요.

---

## B. 방법론 실체 문제 — 판단이 필요한 것

### B1. ★ 두 갈래가 랩 재현에 **다른 규칙**을 적용했다 [critical]
- 원고 §2.2·§2.4는 "laboratory reproduction of outdoor scenes → sensitivity analysis only"라고
  선언한다. 그런데 **실제로는 VR-lab 5편(661·666·999·1122·1272)이 본분석 96편 안에 있다.**
- 확인: `apply_boundary_rulings.py`가 실제 적용한 규칙은 R1(주거소음×비특정PA→배제) ·
  R2(행동의향→SENS_ONLY) · R3(역방향 지각아웃컴→포함) **세 가지뿐**이고, 랩 관련 규칙은 없었다.
  P1은 **인용추적 스크리닝 프롬프트에서 내가 나중에 도입**한 것이고, 그 결과 CT0356 한 편만
  SENS_ONLY가 됐다.
- 즉 **DB 갈래는 랩 재현을 본분석에 넣고, 인용추적 갈래는 민감도로 보냈다.** 갈래 간 불일치다.
- **처리 방침(권고)**: 등록 프로토콜은 옥외장면 재현을 적격 세팅으로 허용했으므로 **본분석 포함이
  등록에 부합**한다. 따라서 ①Methods를 실제 수행대로 고치고 ②**랩 재현 전건 제외 민감도**를
  추가하고 ③갈래 간 불일치를 이탈 로그에 기록한다. 재분석은 불필요하다.

### B2. 출판편향 평가 전무 [critical · 보고]
- 원고 전체에 `publication bias / funnel / Egger` **0회**.
- `analysis_rules.md §6`이 "k≥10만 funnel+Egger, 그 외 서술"이라 사전 규정했는데 **서술조차 없다.**
- 최대 k가 7이라 검정은 불가능하다. **불가능하다는 사실과 그 함의를 명시**해야 한다
  (PRISMA 2020 item 13e·14·22).

### B3. MA4 예측구간 부재 [major]
- I² = 90.5%인데 Abstract가 "0을 제외한 둘 중 하나"로 내세운다. 이 정도 이질성에서는
  **신뢰구간이 아니라 예측구간**이 실질적 불확실성을 보여준다. 계산해 추가한다.

### B4. 등록 하위그룹 6축 미실행 + 미보고 [major]
- 등록서에 하위그룹 분석을 약속했는데 하나도 하지 않았고 **못 한 이유도 안 밝혔다**.
- k=3~7에서는 하위그룹이 무의미하다는 것이 정직한 답이다. 그렇게 쓴다.

---

## C. 보고 누락 — 추가 집필

| # | 항목 | 조치 |
|---|---|---|
| C1 | 이탈 보고 10건이 원고에 없음 | 특히 **D5-4**(CT0414를 넣으면 p .021→.005로 좋아지는데도 제외) — 최고의 방어 자산인데 빠졌다. Methods 또는 Supplementary에 명시 |
| C2 | References 절 없음 | 본문 인용 논문의 정식 서지 생성 필요 |
| C3 | 저자·CRediT·funding·COI·data availability 없음 | LUP 데스크 반려 사유. 전부 추가 |
| C4 | §2.2("intention excluded") ↔ §2.4("sensitivity only") 모순 | §2.2를 SENS_ONLY로 통일 |
| C5 | Fig 7이 본문에서 미인용 + 캡션이 81편 시절 "34 of 84" | 인용 추가 + 캡션 갱신 |
| C6 | MA3 LOO를 일부만 보고 | 4개 전부 보고(불리한 것 포함) — 이 원고의 강점을 확장 |
| C7 | Supplementary 14종 필요한데 영문 실물 0개 | 최소 S1~S4 영문화 |

---

## D. 진행 중 — 코퍼스가 다시 바뀔 수 있음

**OpenAlex 보조검색 329건 스크리닝**이 돌고 있다(등록한 보조 식별원인데 미이행이었고 이탈
로그에도 없었다). 결과에 따라 코퍼스·PRISMA·모든 수치가 바뀔 수 있으므로,
**A9(데이터 코딩)·B(분석 추가)·C(보고 추가)는 그 결과가 나온 뒤 일괄 처리**한다.
A1~A8의 순수 텍스트 오류는 지금 고쳐도 무방하다.

---

## 자평

게이트를 돌리길 잘했다. **A1(13년을 50년으로 씀)과 A2(MA3 귀속)는 둘 다 원고의 가장 눈에 띄는
주장에 박혀 있었고, 둘 다 내가 만든 오류다.** A2는 특히 나쁜데, 내 자신의 민감도 결과가 반증하고
있었는데도 그럴듯한 서사에 맞춰 썼다. 데이터를 보고 쓴 것이 아니라 쓰고 싶은 문장에 데이터를
맞춘 것이다.
