# Paper32 — 투고 준비도 감사 (최초 2026-08-03 · 갱신 2026-08-03 후반)

OSF 등록(osf.io/7ew8q) 항목별로 실제 산출물을 대조한 결과. **✅=해소 · 🟡=부분 · 🔴=미해소**

---

## A. 투고 차단 항목

### ✅ A1. 품질평가 MMAT 2018 — 해소
- 84편(FINAL_INCLUDE 81 + SENS_ONLY 3) 전수 평가. 전문을 읽은 평가로 진행.
- 산출: `quality_all.csv` · `quality_detail_all.csv`(420 문항판정) · `quality_summary.md` · **Fig 8**
- 결과: **high 20 · moderate 34 · low 30**. 범주 = 정량기술 44 · 비무작위 19 · 혼합 12 · 질적 6 · RCT 2.
- 파급 효과 2건:
  - 민감도 ②"저품질 제외"가 실행 가능해졌고, **MA1(보행속도)에서 k=0**이 나왔다.
  - MMAT 범주 사전배정(design 문자열 휴리스틱)이 다수 틀렸음이 드러나 이탈 로그에 기록.
- RoB 2는 대상이 2편뿐이라 병행하지 않고 MMAT 2.x로 대체(이탈 로그 D2-2).

### 🟡 A2. 인용 추적 — 수집·스크리닝 완료, 전문 확보 미완
- 시드 84편 → 신규 후보 2,073건(backward 413 · forward 1,660)
- 제목 스크리닝 428건 → 초록 스크리닝 146건 → **전문 확보 대상 79건**
- **자동 회수 8건(오픈액세스)·미확보 71건** — 출판사 봇 차단(HTTP 403). 우회하지 않음.
- 산출: `citation_tracking*.csv` · `ct_screen_final.csv` · `ct_retrieval_status.csv` ·
  **`ct_download_worklist.md`(직접 다운로드 목록, 우선순위 A 32 · B 39)**
- ⚠️ **남은 사용자 작업**: 목록의 논문을 기관 구독 브라우저로 내려받아 `ct_pdf/`에 넣으면
  본검색과 동일 파이프라인(전문심사 → 추출 → 효과크기 → MMAT)이 이어진다.
- PRISMA(Fig 1)에는 이 갈래를 점선·"awaiting retrieval"로 정직하게 표기.

---

## B. 논문의 핵심 기여

### ✅ B3. 개념 프레임워크 — 해소 (**Fig 6**)
"The soundscape–behaviour loop in public open space" — 등록서가 약속한
*acoustic environment × behaviour × measurement generations* 3자 연결을 한 장에 담았다.
- 양방향 루프(forward: 음환경→평가→행태 / reverse: 활동→음 생성→음환경) + 조절요인
- 행태를 **관여 경사**(Avoid → Pass → Linger → Interact → Appropriate)로 전개하고
  각 단에 k·효과크기·MMAT 품질 구성을 얹었다 — 프레임워크가 곧 증거 지도로 기능한다.
- 측정 3세대(G1 자기보고 31 · G2 체계적 관찰 27 · G3 센싱·궤적 21)를 하단에 배치.

### ✅ B4. 계획·설계 함의 매트릭스 — 해소
`design_matrix.csv`(9 레버) + `design_implications.md`. 레버는 미리 정하지 않고 **추출표 81편의
`exposure` 필드에서 실제로 조작·측정된 것만** 계획 언어로 번역해 귀납했다.

| 레버 | confidence | 핵심 |
|---|---|---|
| L1 음악·오디오 프로그래밍 → 체류 | moderate | 실험 반복 최다(6편·high 3), 다만 MA2 CI가 0 포함 |
| L2 자연음 도입 → 사회적 상호작용 | moderate | **저품질 제외에도 값이 유지되는 유일한 조건대비 클러스터** |
| L3 정온 경로 확보 → 경로·수단 선택 | moderate | MMAT high 4편 최다, GPS·Strava 등 행태의 물리적 흔적 |
| L4 기계소음 제거·차폐·시간대 조정 | moderate | — |
| L5 정온면·정온구역 → 걷기·운동 | low | direction=mixed(연구 간 부호 상충) |
| L6 활동 기능의 음향 구역화 | low | mixed |
| L7 소리를 만드는 활동 프로그래밍 | low | — |
| L8 음향 유도·경고 신호 설계 | low | mixed |
| L9 보행 감속 유도 | **very low** | MA1 붕괴(k=0) · high 0편 |

- **high 등급 레버는 하나도 없다.** 이것 자체가 리뷰의 결론이다.
- **「권고하지 않는다」 7건**을 명시(배경음악=인원 유치 / 수경 마스킹 / 소음저감=신체활동 /
  식재→조류음 사슬 / 전기차 횡단안전 / 77 dB 임계 / SVI 예측지표 처방).
- 독립 검증: 인용된 논문 46편 전부 추출표에 실재·전부 FINAL_INCLUDE, `quality_mix` 9행 모두 실집계와 일치.

---

## C. 보완 항목

| # | 항목 | 상태 |
|---|---|---|
| ✅ C5 | Table 1 study characteristics | `table1_study_characteristics.csv|md` (84행, sid 연속·품질 분포 원본 일치) |
| ✅ C6 | 민감도 분석 6축 실행 | `ma/ma_sensitivity.md` · `ma_sensitivity.csv`(32행) |
| ✅ C7 | 지리·시기 분포 그림 | **Fig 7** · `geo_time_counts.csv` |
| ✅ C8 | 프로토콜 이탈 기록 | `deviation_log.md`(D1-1~D4-2, 12건) |

---

## 발견된 정합 오류와 처리 (자체 감사)

투고 전 심사자가 잡을 만한 내부 불일치 2건을 이번에 찾아 고쳤다.

1. **MA1 주분석 정의 불일치** — `ma_summary.md`는 k=4(481 제외)인데 민감도 스크립트 초판이 k=5로
   계산했다. 정본에 맞춰 수정하고, 481 포함형은 민감도 ⑥으로 분리해 **제외가 보수적 방향**임을
   수치로 밝혔다(포함 시 g −0.742·p 0.082 → 제외 시 −0.500·p 0.179).
2. **측정세대 카운트 하드코딩** — Fig 5의 31/27/21이 데이터에서 산출되지 않고 코드에 박혀 있었고,
   Table 1의 코딩으로 다시 세면 47/35/19였다. 원인은 단일코딩 vs 다중코딩 규칙 차이.
   **다중코딩(한 연구가 두 세대를 쓰면 둘 다 계상)으로 통일**하고 Fig 5·Fig 6 모두
   `table1_study_characteristics.csv`에서 산출하도록 바꿨다(하드코딩 제거).
   시기별: ≤2009 G1 1·G2 2 / 2010–2019 G1 13·G2 9·G3 3 / 2020– G1 33·G2 24·G3 16.

## 현재 상태 요약

**분석·근거 산출은 완료됐고, 남은 병목은 두 가지다.**

1. **인용추적 71편의 전문 확보** — 자동화가 막힌 지점. 사용자 브라우저가 필요하다.
   이것이 끝나야 PRISMA의 포함 편수가 확정된다.
2. **집필** — Figure 8종·Table 2종·메타분석 4클러스터·민감도 6축·설계 레버 9개·이탈 로그가 준비됐다.

투고 차단 수준의 방법론 공백은 남아 있지 않다. A2는 "미완"이지만 **PRISMA가 허용하는 방식으로
투명하게 표기**되어 있어, 최악의 경우 현 상태로도 투고 자체는 가능하다(다만 심사자가
"인용추적을 끝내라"고 요구할 가능성이 높으므로 권하지 않는다).

## 산출물 지도

- 논문 표 = `table1_study_characteristics.csv|md`(Table 1) · `design_matrix.csv`(Table 2) ·
  `design_implications.md`(Discussion 재료)
- 코퍼스·판정 = `ft_verdicts_v2.csv` · 추출 = `ft_extraction_v2.csv` · 효과크기 = `effect_sizes_all.csv`
- 정량 종합 = `ma/ma_summary.md` · 민감도 = `ma/ma_sensitivity.md`
- 품질 = `quality_all.csv` · `quality_summary.md`
- 인용추적 = `ct_screen_final.csv` · `ct_retrieval_status.csv` · `ct_download_worklist.md`
- 절차 투명성 = `prisma_flow.md` · `deviation_log.md` · `analysis_rules.md` · `source_flags.md`
- 그림 = `../figures/` (Fig 1~8, PNG+PDF, `FIGURES_README.md`에 캡션 초안)
