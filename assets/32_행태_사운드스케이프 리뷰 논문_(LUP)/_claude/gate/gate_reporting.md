# Gate — 보고 표준 준수 (PRISMA 2020 + 사전등록 이행)

독립 검증 게이트 · 축: reporting standards · 검증일 2026-08-06
검증 대상: `01_논문작업/Manuscript_EN_20260806_draft.md` (448행, 96편 코퍼스)
대조 근거: `_claude/review_protocol_draft_20260802_203917.md` · `_claude/osf_registration_draft_EN_20260802_211715.md` ·
`_claude/search_strings_20260802_205110.md` · `_claude/fulltext/{prisma_flow,deviation_log,analysis_rules,ct_retrieval_summary,quality_truncation_effect,source_flags}.md` ·
`_claude/fulltext/ma/ma_sensitivity_v2.md`

---

## 판정

**REVISE** — 방법론 실체와 수치는 근거 파일과 정합하나, 보고층에서 PRISMA 27항목 중 **5항목이 완전 누락(item 14·21·25·26·27)**이고
등록서에 약속한 보조 검색원(OpenAlex 352건)·인접 리뷰 참고문헌 스크리닝·6축 하위그룹 분석이 **이탈 로그에도 원고에도 없이 소멸**했다.
현 상태로는 투고 불가이나, 결함이 전부 보고·문서 층위여서 원 분석을 다시 돌리지 않고 교정 가능하다.

**○ 10 / △ 12 / ✕ 5** (27항목)

---

## PRISMA 2020 대조표

| item | 항목명 | 충족 | 원고 위치 | 비고 |
|---|---|---|---|---|
| 1 | Title | ○ | L1 제목 | "a systematic review and meta-analysis" 명시 |
| 2 | Abstract | △ | L9–30 | PRISMA-A 중 funding 진술 없음 · 참가자/표본 규모 없음 · registration(osf.io/7ew8q)은 L15에 있음 |
| 3 | Rationale | ○ | §1.1–1.4 (L36–56) | 인접 리뷰 대비 포지셔닝 포함 |
| 4 | Objectives | ○ | §1.5 (L58–65) | RQ1–RQ4. ⚠️등록서는 RQ1/RQ1b/RQ2/RQ3/RQ4 5문항 — 재번호 사실 미보고 |
| 5 | Eligibility criteria | ○ | §2.2 (L82–101) | 5기준 명시. ⚠️§2.2는 intention을 "excluded", §2.4는 "sensitivity analysis only" — **원고 내부 모순** |
| 6 | Information sources | △ | §2.3 (L105–111) | 3개 DB + 인용추적 ○. **등록한 OpenAlex 보조검색(873건·고유 352건)과 "인접 리뷰 참고문헌 스크리닝" 미보고·미실행** · 인용추적 실행일 없음 |
| 7 | Search strategy | △ | §2.3 (L107) "Supplementary S3" | 전문(全文) 검색식은 `search_strings_*.md`에 존재하나 **영문 Supplementary 파일 미작성**. 인용추적 3블록 제목필터 규칙도 S3에 편입 필요 |
| 8 | Selection process | △ | §2.4 (L115–131) | 단계·자동분류는 기술. **몇 명이 했는지 Methods에 없음** — §4.6(L418)에만 "one reviewer with AI assistance". 등록한 Rayyan/ASReview 미사용 사실도 미보고 |
| 9 | Data collection process | ○ | §2.5 (L135–138) | "One reviewer extracted data … 16-field schema" 명시 |
| 10 | Data items | △ | §2.2-4 (L93–97), §2.5 (L135–137), §2.6 (L162–164) | 아웃컴 도메인·변수 목록·다중효과 선택규칙 있음. **결측 정보 가정(deviation D3-3의 복원 3건: ID 532 균등분할·665 p→t 역산·323 SE→SD)은 어디에도 기술 없음** |
| 11 | Study risk of bias assessment | △ | §2.5 (L140–150) | MMAT 2018 · 절단 재평가 감사 ○. **평가자 수·독립 이중평가 여부 미기재** · MMAT 범주 재배정(deviation D2-1) 미보고 |
| 12 | Effect measures | ○ | §2.6 (L155–159) | Hedges' g · Chinn 변환 · Fisher z 명시 |
| 13 | Synthesis methods | △ | §2.6 (L154–169) | 13a·13b·13c·13d·13f ○ (REML+HK). **13e 이질성 탐색(하위그룹·메타회귀) 전무** — 등록서 6축·`analysis_rules §6` 3축 모두 미실행·미해명. 13f는 7축을 보고하나 post hoc을 1축으로만 표기(실제 ⑤⑥⑦ 3축이 `analysis_rules §6` 밖) |
| 14 | Reporting bias assessment | ✕ | — | **원고 전체에 publication/reporting bias·funnel·Egger·small-study 언급 0회.** `analysis_rules §6`은 "k≥10만 funnel+Egger, 그 외 서술"을 규정했는데 서술조차 없음 |
| 15 | Certainty assessment | △ | §4.2 (L314–316) | confidence 등급 산출 근거는 "study count, MMAT quality composition, CI 0 포함, setting diversity, sensitivity" 로 기술. **Methods가 아닌 Discussion 소재 · GRADE 등 표준체계 미준거·미인용 · 등급 정의(high/moderate/low/very low)의 조작적 기준 미제시** |
| 16 | Study selection | △ | §3.1 (L177–195) + Fig 1 | 16a 흐름·수치 ○(prisma_flow.md와 전건 일치, 100−16−3=81 · 54−38−1=15 검산 통과). **16b 전문단계 배제 연구의 서지 인용 목록 없음**(CSV만 존재) |
| 17 | Study characteristics | ○ | §3.2 (L199–211) + Table 1 | 세팅·설계·지역·연도·표본 |
| 18 | Risk of bias in studies | ○ | §3.3 (L215–226) + Fig 8 | 22/40/34 tier + 문항 수준 결손 패턴 |
| 19 | Results of individual studies | △ | Fig 2 (forest) | 개별 효과추정치·정밀도가 **본문 표로 제시되지 않음** · 각 연구 요약통계(집단별 n·평균) 미제시 |
| 20 | Results of syntheses | ○ | §3.4 (L230–263) | 20a 연구특성·RoB 서술 ○ · 20b 추정치·CI·I² 표 ○ · 20c 이질성 원인 서술(L255–256) ○ · 20d 민감도 ○. ⚠️③⑤⑥축 결과는 본문 미보고 |
| 21 | Reporting biases | ✕ | — | item 14와 동일 — 결과절에도 전무 |
| 22 | Certainty of evidence | △ | §4.2 (L318–321) + Table 3 | 등급은 **planning lever 10개**에 부착. **메타분석 4클러스터(주 아웃컴) 자체에 대한 certainty 등급 없음** — 심사자가 요구하는 형식과 불일치 |
| 23 | Discussion | ○ | §4.1–4.3 (L295–370), §4.6 (L412–421), §4.5 (L393–408) | 23a 해석 · 23b 증거 한계 · 23c 리뷰 과정 한계(단일 평가자·미확보 114건) · 23d 함의 모두 존재 |
| 24 | Registration and protocol | △ | §2.1 (L73–78) | 24a 등록번호 ○. **24b 프로토콜 접근처 문장 없음**(등록 ID만) · 24c는 "Supplementary S1"에 위임했으나 **S1 미작성 + 로그 자체에 누락 이탈 다수**(아래 §이탈 보고 누락). 등록 제목과 원고 제목 상이(meta-analysis 추가) 미보고 |
| 25 | Support (funding) | ✕ | — | 원고에 funding/support 문장 없음 |
| 26 | Competing interests | ✕ | — | 원고에 competing interests/COI 문장 없음 |
| 27 | Availability of data, code, other materials | ✕ | — | data/code availability 문장 없음 · 공개 저장소·DOI 없음 · S1–S3 실물 없음 |

---

## 치명적 누락

심사에서 즉시 지적될 순서.

1. **보고편향 평가 완전 부재 (item 14 + 21).** 4개 클러스터 중 어디에도 funnel·Egger·서술적 small-study 논의가 없다.
   `analysis_rules.md §6`이 "k≥10만 funnel+Egger, **그 외 서술**"을 사전 규정했으므로 k=3~7 클러스터에 대한 **서술적 평가는 등록된 의무**인데
   실행 흔적이 없다. MA3는 k=4에 두 편이 1975·1988년 고전이고 MA1은 전건 low quality — 심사자가 가장 먼저 묻는 지점이다.
   → 최소한 "k<10이라 정량 검정 불가, 대신 ①검색 범위 ②회색문헌 배제의 방향성 ③유의 결과의 단일연구 의존"을 명시한 단락 필수.

2. **투고 필수 front/back matter 전무 (item 25·26·27 + 참고문헌·저자).** 원고에 **References 절 자체가 없다**(§5 Conclusions로 끝남).
   Mathews & Canon 1975, Moser 1988, Zhang et al. 2025, ISO 12913 등을 본문에서 인용하는데 서지가 없다.
   저자·소속·CRediT·funding·COI·data availability 모두 미작성. LUP는 이 중 어느 하나라도 없으면 데스크 단계에서 반려된다.

3. **등록 하위그룹 분석 6축의 침묵 (item 13e).** OSF 등록서 *Analysis of subgroups or subsets*는
   행태 도메인·음원 범주·세팅 유형·연구 설계·측정세대·방향 6축을 약속했고 `analysis_rules.md §6`도 세팅/설계/측정세대를 재확인했다.
   원고는 하위그룹 메타분석을 **하나도 보고하지 않으면서 못 한 이유도 말하지 않는다**. deviation_log에도 항목이 없다.
   등록서를 열어본 심사자에게는 "결과가 안 나와서 뺐나"로 읽힌다 — k 부족이라는 정직한 사유를 명시해야 방어된다.

(차순위) **Supplementary S1·S2·S3가 실물로 존재하지 않는다.** 원고가 3회 참조하나 산출된 파일은 한국어 작업문서뿐이다.
(차순위) **§2.2와 §2.4의 intention 처리 모순** — §2.2는 "excluded", §2.4는 "sensitivity analysis only".
  deviation_log D5-2가 이미 "R2(=SENS_ONLY)를 정본으로 삼는다"고 해소한 사안인데, 원고가 폐기된 쪽(P2)을 §2.2에 박제했다.
(차순위) **Fig 7이 본문에서 한 번도 인용되지 않는다**(Fig 1·2·3·4·5·6·8만 참조). 게다가 `FIGURES_README.md`의 Fig 7 캡션은
  "34 of 84 (40%)"로 **81편 시절 수치**가 남아 있다(정본 38 of 96).

---

## 이탈 보고 누락

`deviation_log.md`에 있는데 원고에 없는 것 — 그리고 **로그에도 없는데 등록서에는 있었던 것**(더 심각).

### A. 로그에 있으나 원고 미보고

| 로그 | 내용 | 원고 상태 | 심각도 |
|---|---|---|---|
| **D5-4** | MA 클러스터 편입 판단 4건. **CT0414는 넣으면 MA3가 p .021→.005로 좋아지는데 독립성 위배로 제외**, CT0175는 넣으면 MA2가 p .074→.327로 나빠지는데 아웃컴 불일치로 제외, CT0137은 규약 밖 d→r 변환이라 변형분석 분리 | **전무.** 원고 어디에도 "포함 후보였으나 제외한 효과"가 없다 | ★최고 — 데이터를 보고 내린 편입/제외 판단이 보고되지 않으면 selective-inclusion 의심을 부른다. 역설적으로 **불리한 쪽으로 판단한 사례(CT0414)라 보고하면 오히려 강점** |
| **D3-1** | 등록한 5클러스터 중 **방문/공간이용 클러스터를 k<2로 서술종합 이관** | 전무. §3.4는 "Four clusters met the pooling threshold"라고만 함 | ★높음 — item 13e/24c 직결 |
| **D3-6** | MA1 주분석에서 **음악 자극 연구(ID 481) 제외**. 포함 시 g −0.500→−0.742, p .179→**.082** | §2.6이 "an alternative definition of the walking-speed pool"을 민감도 축으로 언급만 하고 **결과 수치를 보고하지 않음**. 무엇을 뺐는지도 없음 | ★높음 — "제외가 보수적 방향"이라는 방어 논리를 원고가 쓰지 않고 있다 |
| **D2-1** | 사전배정 MMAT 범주를 **전문 판독 결과로 전면 교체**(범주가 바뀌면 적용 문항·등급이 바뀜) | 전무 | 중 — item 11 |
| **D3-3** | 원문 미보고 값 **복원 3건**(532 균등분할 가정 · 665 p→t 역산 · 323 SE→SD) | §2.6이 민감도 축 이름으로만 암시. 어느 연구에 무슨 가정을 했는지 없음 | 중 — item 10/13b |
| **D3-5** | 민감도 ③(관측 n) **부분 실행** — MA3의 931·1069는 참가자 n 복원 불가로 서술만 | 전무. 원고는 7축을 모두 실행한 것처럼 읽힘 | 중 |
| **D1-2** | 경계 3원칙이 **사후 정식화**(등록서는 PICO 수준만) | §2.4가 "rules fixed before individual adjudication"이라 쓰나 **등록 이탈이라는 사실은 미표기** | 낮 |
| **D3-4** | 이질성 추적 중 발견한 **자체 코딩 오류 정정**(Moser 1988 OR 부호 반전, I² 93.6→43.5%) | 전무 | 낮 — 생략 가능하나 로그가 "투명성 기록"으로 남기라 함 |
| **D4-1** | 개념 프레임워크를 Results→**Discussion**으로 이동(등록서는 산출물로 명시) | 전무 | 낮 |
| **D5-1** | 배제코드 **X7(언어) 신설** — 중국어 게재지 1건 | §3.1이 CT 갈래 배제 38건을 사유 없이 총계로만 보고 | 낮 (Fig 1에 있으면 해소) |

### B. ★등록서에 있으나 로그에도 원고에도 없는 것 (미기록 이탈)

이 네 건이 이번 검증의 최대 발견이다. 이탈 로그의 존재 이유가 "숨기지 않고 기록한다"인데, **누락된 이탈은 로그가 잡지 못했다.**

1. **OpenAlex 보조 검색의 소멸.** 등록서 *Searches*: "Databases: WoS; Scopus; PubMed. **Supplementary: OpenAlex**".
   실제로 873건을 수집했고 3개 DB에 없는 **고유 352건**을 `openalex_supplement_20260802_211715.csv`로 분리했다.
   `prisma_flow.md` 주석 3은 "인용 추적 단계에서 재검토 예정 — 현 흐름도에는 미포함"이라 적었으나,
   **검증 결과 `fetch_citation_tracking.py`는 `merged_pool`(3개 DB 1,316건)만 입력으로 쓴다 — 352건은 끝내 스크리닝되지 않았다.**
   원고 §2.3은 "Three databases were searched"라고만 하여 등록한 네 번째 소스를 언급조차 하지 않는다.
   → 352건을 스크리닝하거나, **못 한 사실과 사유를 이탈로 기록**해야 한다. 현 상태는 등록 정보원의 무언(無言) 삭제다.

2. **"인접 리뷰 참고문헌 목록 스크리닝"의 소멸.** 등록서 *Searches* 보조방법에 명시("screening of reference lists of adjacent reviews").
   실행 흔적·로그 항목 모두 없음. 원고 §2.3은 인용추적만 보고한다.

3. **등록 하위그룹 6축의 소멸.** 위 §치명적 누락 3 참조. 로그에 항목 없음.

4. **SWiM 준거의 소멸.** 등록서 *Strategy for data synthesis*: "Structured narrative synthesis **following SWiM guidance**".
   원고에 SWiM 언급 0회. 서술 종합을 실제로 SWiM 9항목으로 보고하지 않았다면 이탈로 기록해야 한다.

(추가·중) **사전규약 §2 대비프레임 위반 가능성.** `analysis_rules.md §2`는 (N) 부정 음환경 대비와 (P) 긍정 음환경 대비를
**두 하위그룹으로 분리 풀링**하도록 규정했다. MA3는 "natural/quiet vs noise"로 두 프레임을 **합산 풀링**한다.
원고 §4.2가 그 혼재를 서술로 인정하면서도(기계소음 2건 vs 자연음 2건) **사전규약 이탈로는 표기하지 않는다.** 로그에도 없다.

(추가·소) **등록 스크리닝 도구(Rayyan/ASReview) 미사용**, **등록 시 미확정으로 남긴 평가자 수 괄호**
(`[Single reviewer with AI-assisted second check / two independent reviewers — 확정 필요]`)가 단일 평가자로 확정된 경위 미보고.

---

## Supplementary 목록

원고가 참조하는 것은 S1·S2·S3 세 개뿐이나, PRISMA 16b·11·19·27을 충족하려면 아래 범위가 필요하다.
**현재 실물로 존재하는 영문 Supplementary 파일은 0개다.**

| 번호 | 내용 | 현재 파일 | 상태 |
|---|---|---|---|
| **S1** | Protocol deviation log (등록 대비 전 이탈) | `_claude/fulltext/deviation_log.md` | ⚠️한국어 · **미기록 이탈 4+1건 추가 후 영문화 필요**(§이탈 보고 누락 B) |
| **S2** | Analysis decision protocol (효과크기·대비프레임·다중효과·유사반복) | `_claude/fulltext/analysis_rules.md` | ⚠️한국어 · 영문화 필요 · "계산 착수 전 확정" 날짜 명시 필요 |
| **S3** | Full search strategies (WoS/Scopus/PubMed 전문 + 실행일·히트수) | `_claude/search_strings_20260802_205110.md` | ⚠️쿼리는 영문이나 문서 골격이 한국어 · **인용추적 3블록 제목필터 규칙(T1~T4)과 `citation_tracking_triaged.csv` 편입 필요** |
| **S4** | PRISMA 2020 checklist (27항목 + Abstract 12항목, 페이지 매핑) | **없음** | ❌미작성 — LUP 투고 필수 첨부 |
| **S5** | 포함 96편 전체 목록 + 서지 + 특성표 (Table 1 확장) | `fulltext/table1_v2.md` / `.csv` | 데이터 존재 · 서지 인용 형식 미부착 |
| **S6** | 전문단계 배제 목록 + 사유 (DB 16건 + CT 38건 + SENS_ONLY 4건) | `fulltext/ft_verdicts_v2.csv`, `ct_verdicts_final.csv` | 데이터 존재 · **item 16b 충족하려면 서지 인용 필수** |
| **S7** | MMAT 2018 문항 수준 판정표 (100편 × S1·S2·Q1–Q5) | `fulltext/quality_v2.csv`, `quality_detail_v2.csv` | 존재 · 영문 헤더화 필요 |
| **S8** | 메타분석 입력 데이터 (효과크기·변환경로·출처 위치 verbatim) | `fulltext/ma/ma_*_input.csv`, `ct_effect_sizes.csv`, `effect_sizes_all.csv` | 존재 · **복원 3건(D3-3)의 note 필드 노출 필요** |
| **S9** | 민감도 분석 전체 표 (7축 × 4클러스터) | `fulltext/ma/ma_sensitivity_v2.md` / `.csv` | ⚠️한국어 · 영문화 · **사전지정/post hoc 구분 열 추가 필요** |
| **S10** | 전문 추출 절단 감사 (13편 재평가 대조표) | `fulltext/quality_truncation_effect.md` | ⚠️한국어 · 원고 §2.5·§4.4의 핵심 주장 근거이므로 반드시 동반 |
| **S11** | 인용추적 트리아지 전건 (2,073 → T1~T4 배정) | `fulltext/citation_tracking_triaged.csv` | 존재 · D1-3 재현성 근거 |
| **S12** | Planning levers 매트릭스 (Table 3 근거·confidence 산출) | `fulltext/design_matrix_v2.csv`, `design_implications_v2.md` | ⚠️한국어 · **confidence 등급의 조작적 정의를 여기에 명시해야 item 15 방어 가능** |
| **S13** | 원문 검증 플래그 (오타·표 이상·중복표본) | `fulltext/source_flags.md` | ⚠️한국어 · 선택적이나 강력한 투명성 자산 |
| **S14** | 분석 코드 (검색·스크리닝·MA·figure 생성) | `_claude/*.py` 40여 개 | ❌공개 저장소 미생성 — item 27 |

---

## 투고 전 체크리스트

### 원고에 문장을 새로 써야 하는 것 (아직 존재하지 않음)

- [ ] **References 절** — 본문 인용 전건 + 포함 96편 서지 (현재 원고에 참고문헌 절 자체가 없음)
- [ ] **저자·소속·교신저자** 블록
- [ ] **CRediT authorship contribution statement**
- [ ] **Funding / Support 문장** (item 25 — 없으면 "This research received no specific grant…" 명시)
- [ ] **Declaration of competing interest** (item 26)
- [ ] **Data availability statement** (item 27) — OSF 저장소 DOI 발급 후 "추출표·MMAT 판정·MA 입력·분석 코드는 …에서 이용 가능" 형식
- [ ] **Reporting bias 단락** (item 14 Methods + item 21 Results) — k<10 사유 + 서술적 평가
- [ ] **Certainty 조작적 정의**를 Methods로 이동 (item 15) + 4개 MA 클러스터 자체의 certainty 등급 (item 22)
- [ ] **하위그룹 미실행 사유** 1–2문장 (item 13e) — "등록한 6축은 클러스터당 k≤7로 하위그룹 추정이 불가능"
- [ ] **평가자 수를 Methods §2.4에 명시** (현재 §4.6에만 존재)
- [ ] **인용추적 실행일** 명시 (§2.3)
- [ ] Ethics statement (2차 문헌 연구이므로 비해당 명시)

### 원고를 고쳐야 하는 것 (모순·누락)

- [ ] **§2.2 ↔ §2.4 intention 모순 해소** — deviation D5-2가 정한 정본(SENS_ONLY)으로 §2.2 수정
- [ ] **§2.3에 OpenAlex 보조검색 처리 명시** — 스크리닝하거나, 미실행 사실·사유를 이탈로 기록
- [ ] **§2.6 민감도 7축 중 post hoc 구분 정정** — 현재 1축만 post hoc 표기, 실제 ⑤⑥⑦ 3축이 `analysis_rules §6` 밖
- [ ] **§3.4에 D5-4 편입/제외 판단 보고** — 특히 CT0414(제외가 결과에 불리)를 밝히는 것이 최대 방어 자산
- [ ] **§2.6/§3.4에 MA1 음악연구 제외(D3-6)와 그 민감도 수치(g −0.742, p .082) 보고**
- [ ] **§2.6/§3.4에 등록 5클러스터 → 4클러스터 축소(D3-1) 보고**
- [ ] **§2.5에 MMAT 범주 재배정(D2-1) 보고**
- [ ] **§2.6에 복원 3건(D3-3)의 대상·가정 명시**
- [ ] **§2.6에 사전규약 (N)/(P) 분리 풀링 대비 MA3 합산 사유 명시** 또는 분리 재보고
- [ ] **Fig 7을 본문에서 인용하거나 삭제** + `FIGURES_README.md` Fig 7 캡션의 "34 of 84" → "38 of 96" 정정
- [ ] 등록 제목과 원고 제목 상이(meta-analysis 추가)를 §2.1 또는 S1에 1줄 기록

### 재현성 — 지금 공개 가능한 수준

| 요건 | 현 상태 | 공개 가능성 |
|---|---|---|
| 검색식 전문 | ✅ 3개 DB 쿼리 + 실행일·히트수·벤치마크 5/5 리콜 확보 | **즉시 가능**(S3 영문화만) |
| 인용추적 필터 규칙·배정 | ✅ `citation_tracking_triaged.csv` 전건 보존 | **즉시 가능** |
| 배제 사유별 목록 | ✅ 스크리닝·전문 판정 CSV 전건 존재 | **가능** — 단 서지 인용 부착 필요(item 16b) |
| 추출 데이터 | ✅ `corpus_v3_extraction.csv` 16필드 + verbatim 효과 통계·출처 위치 | **가능** — 저작권 안전(수치·위치만) |
| 품질평가 문항 수준 | ✅ `quality_v2.csv` / `quality_detail_v2.csv` 100편 | **가능** |
| MA 입력·민감도 | ✅ `ma_*_input.csv`, `ma_sensitivity_v2.csv` | **가능** |
| 분석 코드 | ✅ 40여 스크립트 존재(검색→스크리닝→MA→figure 전 구간) | **가능** — ⚠️저장소 미생성·라이선스 미결정 |
| 전문 PDF | ❌ 저작권상 비공개 | 불가(정상) |
| **결정적 공백** | **OSF 저장소·DOI가 없어 위 전부가 "공개 가능"일 뿐 "공개됨"이 아니다** | 등록(osf.io/7ew8q) 프로젝트에 컴포넌트 추가 필요 |

---

## 검증 부기 — 정합이 확인된 것

누락만 보고하지 않기 위해, 대조 결과 **일치**한 항목을 남긴다.

- PRISMA 흐름 수치 전건 정합: DB 2,073−757=1,316 · 189−89=100 · 100−16−3=**81** · CT 54−38−1=**15** · 합 **96**(+민감도 4=100)
- A군/B군 적중률 46%(17/37)·0%(0/17)와 회수율 92%(37/40)가 `prisma_flow.md`·`ct_retrieval_summary.md`와 일치
- 미확보 총계 114건 = DB 89 + CT 25 ✅
- MA 4클러스터 추정치·CI·p·I² 전건이 `ma_sensitivity_v2.md` 주분석 행과 일치
- MA3 LOO −CT0025 → p .053, MA4 인용추적 제외 → r .426 — 원고 §3.4 서술과 일치
- MMAT 22/40/34(96편)가 `quality_v2.csv`(100편: high 24·mod 41·low 35)에서 민감도 4편 차감분과 정합
- 절단 재평가(13편 중 7편 판정 변경·4편 등급 이동·3편 2등급)가 `quality_truncation_effect.md`와 일치
- D1-3(인용추적 자동 필터)·D2-2(RoB 2 생략)·D5-3(절단 재평가)·D5-5(동일표본)·D4-2(언어·문헌유형)는
  **원고가 이탈 또는 한계로 명시적으로 보고하고 있다** — 이 다섯은 모범 사례
