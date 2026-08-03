# Paper32 — PRISMA 2020 흐름도 (확정 수치, 2026-08-03)

검색 실행일 2026-08-02 · 사전등록 osf.io/7ew8q · 산출 근거 = `prisma_numbers_*.json`,
`screening_results_*.csv`, `fulltext_status_20260803.csv`, `ft_verdicts_v2.csv`

## 흐름

```
IDENTIFICATION
  데이터베이스 검색 (2026-08-02)
    Web of Science Core Collection ............ 1,010
    Scopus .................................... 850
    PubMed .................................... 213
    ─────────────────────────────────────────────────
    합계 .................................... n = 2,073
  중복 제거 ................................. −757
                                              ↓
SCREENING
  제목·초록 스크리닝 ....................... n = 1,316
  배제 ...................................... −1,127
    EX1 동물·야생(soundscape ecology) ....... 479
    EX4 지각·건강·생리 아웃컴만 ............. 325
    EX2 세팅 부적합(실내 등) ................ 133
    EX5 비실증(리뷰·시뮬레이션) ............. 109
    EX3 음환경 변수 없음 .................... 81
                                              ↓
ELIGIBILITY
  전문 확보 대상 ............................ n = 189
  전문 미확보 ............................... −89
    구독 접근 불가·기관 미보유 .............. 84 (전부 BORDERLINE 등급)
    유료 개별구매 필요 ...................... 3
    초록만 존재(전문 미발행) ................ 2
                                              ↓
  전문 평가 ................................. n = 100
  전문 단계 배제 ............................ −16
    세팅 부적합(주거·공간 비특정) ........... 7
    관찰가능 행태 아웃컴 없음 ............... 6
    지각·평가 아웃컴만 ...................... 1
    음환경 노출 변수 없음 ................... 1
    비실증(서술적 리뷰·교육 성찰) ........... 1
  민감도 분석 전용(행동의향 아웃컴) ......... −3
                                              ↓
INCLUDED
  질적 종합(서술·증거지도) 포함 ............. n = 81
  정량 종합(메타분석) 기여 .................. n = 15 (4개 클러스터, 중복 계상 없음)
    보행속도 4 · 체류 3 · 사회적 상호작용 3 · 지각-행태 상관 5
```

## 인용 추적 갈래 (Identification of studies via other methods) — 2026-08-03 실행

등록 프로토콜 *"Backward and forward citation tracking of all included studies"* 이행.
포함 84편을 시드로 OpenAlex API에서 참고문헌(backward)·피인용(forward)을 수집하고,
DB 검색 풀(1,316)·최종 코퍼스와 DOI·제목 정규화로 대조해 신규 후보만 남겼다.

```
IDENTIFICATION (other methods)
  인용 추적 시드 ............................ 84편(FINAL_INCLUDE 81 + SENS_ONLY 3)
  backward 참고문헌 고유 3,456 · 2회 이상 인용분만 채택
  forward 피인용 고유 2,407
  기존 풀·코퍼스 중복 제거 + 영어·학술지 논문 필터
                                              ↓
  신규 후보 ................................. n = 2,073  (backward 413 · forward 1,660)
                                              ↓
  우선순위 분류 — 등록된 3블록 검색 논리(음환경 × 행태 × 옥외세팅)를 제목에 적용
    T1 3블록 모두 충족 ...................... 11
    T2 2블록 충족 ........................... 417
    T3 동물·생태음향 ........................ 26      ← 자동 배제
    T4 1블록 이하 ........................... 1,619   ← 자동 배제
                                              ↓
SCREENING (other methods)
  제목 스크리닝 ............................. n = 428
  배제 ...................................... −282
    E4 행태 아웃컴 없음(지각·건강·생리만) ... 211
    E3 음환경 변수 없음 ..................... 29
    E5 비실증(리뷰·방법론) .................. 26
    E2 세팅 부적합(실내 등) ................. 10
    E1 동물·생태음향 ........................ 6
                                              ↓
  초록 스크리닝 ............................. n = 146
  배제 ...................................... −67
    E3 음환경 변수 없음 ..................... 37
    E4 행태 아웃컴 없음 ..................... 26
    E5 비실증 / E2 세팅 부적합 .............. 4
                                              ↓
  전문 확보 대상 ............................ n = 79
    초록에서 행태 아웃컴 확인 ............... 40
    초록 기계 확보 불가(제목만 판단) ........ 39
  전문 확보 완료 ............................ 8   (오픈액세스 자동 회수)
  전문 확보 대기 ............................ 71  ← ⚠️현재 상태
```

**⚠️ 현 시점 상태**: 이 갈래는 **전문 심사 이전 단계에서 멈춰 있다.** 71편은 출판사 봇 차단(HTTP 403)으로
자동 회수가 불가능해 기관 구독 브라우저에서 직접 내려받아야 한다(목록 = `ct_download_worklist.md`).
따라서 **본 흐름도의 INCLUDED 수치(81편)는 DB 검색 갈래만 반영한 값**이며, 인용추적 갈래가 완료되면
갱신된다. PRISMA 2020은 이 상태를 "reports sought for retrieval"에서 멈춘 것으로 표기하도록 허용한다.

**숫자 우연 주의**: DB 검색 합계와 인용추적 신규 후보가 **둘 다 2,073**이다. 복사 오류가 아니라 우연이며,
각각 `prisma_numbers_*.json`과 `citation_tracking_log.json`에서 독립적으로 산출됐다.

## 주석 (논문 각주·Methods 기재용)

1. **미확보 89건의 성격**: 84건은 제목·초록 단계에서 BORDERLINE(불확실) 등급이었고 기관 구독 범위 밖.
   INCLUDE 등급 5건만 확보 실패(유료 3·초록만 2). 즉 **확실 적격군의 회수율은 78/83 = 94.0%**로,
   미확보가 결론을 좌우할 위험은 낮다(민감도 논의에 명기).
2. **BORDERLINE 등급의 낮은 생존율**: 확보된 BORDERLINE 22건 중 전문심사 통과는 소수 —
   미확보 84건이 모두 적격이었을 가능성은 낮다는 경험적 근거.
3. **보조 식별원**: OpenAlex 검색으로 3개 DB 미포착 352건을 별도 확인(`openalex_supplement_*.csv`).
   인용 추적 단계에서 재검토 예정 — 현 흐름도에는 미포함.
4. **SENS_ONLY 3건**은 등록 프로토콜대로 본분석 제외·민감도 분석 전용(행동의향 아웃컴).
5. **스크리닝 방식**: AI 보조 사전분류 + 인간 검증(등록 프로토콜 §Screening). 전문 심사는 원문 전수 정독.

## 보고 문장 초안 (Results 첫 단락)

> Database searches identified 2,073 records (Web of Science 1,010; Scopus 850; PubMed 213).
> After removing 757 duplicates, 1,316 records were screened by title and abstract, of which 1,127
> were excluded — most frequently because they concerned animal rather than human behaviour
> (soundscape ecology; n = 479) or reported only perceptual, health, or physiological outcomes
> (n = 325). Of the 189 records sought for retrieval, 89 could not be obtained (84 of which had been
> classified as uncertain at the screening stage), leaving 100 full texts assessed for eligibility.
> Sixteen were excluded at this stage and three were reserved for sensitivity analysis, yielding
> **81 studies in the qualitative synthesis and 15 contributing to the four meta-analytic clusters**.
