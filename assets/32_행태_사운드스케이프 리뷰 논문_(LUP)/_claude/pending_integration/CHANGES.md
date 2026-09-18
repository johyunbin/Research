# paper32 — 추가 전문평가 77건 + 1226 재판정 코퍼스 반영 (사본 드라이런 결과, 2026-09-18 02:0x KST)

작업 위치: `scratchpad/integ/assets/32_행태_사운드스케이프 리뷰 논문_(LUP)/`(git archive 사본).
비교 기준: `scratchpad/integ_base/`(손대지 않은 같은 사본) · 원본 `C:\Users\wh850\Research\...` 는 **읽기만** 했다.

## 0. 원본에 그대로 옮기는 절차

```bash
cd "C:/Users/wh850/Research/assets/32_행태_사운드스케이프 리뷰 논문_(LUP)/_claude"
# ① 스크립트 8개 복사(신규 1 + 수정 7) — 또는 scripts_20260918.patch 적용
#    integrate_newft_into_corpus.py · finalize_oa_supp_branch.py · fig_data.py · make_fig1_prisma_v2.py
#    rebuild_table1.py · build_appendix.py · verify_manuscript.py · verify_consistency_v2.py
# ② 정본 반영(멱등 — 이미 반영됐으면 건너뛴다)
python integrate_newft_into_corpus.py fulltext/newft_final                     # 검사·계획만
python integrate_newft_into_corpus.py fulltext/newft_final --apply --fetch-authors
# ③ 다운스트림 재생성(지도 §3 순서, ma_*.py 제외)
python integrate_newft_into_corpus.py --downstream --log-dir ../_logs
# ④ 내일 추가 평가분이 생기면 폴더를 인자로 덧붙이기만 하면 된다
python integrate_newft_into_corpus.py fulltext/newft_final fulltext/newft_final_20260919 --apply --fetch-authors --downstream
```

데이터 파일은 사본에서 옮겨 붙이지 말고 **원본에서 ②를 다시 실행**하는 편이 안전하다(사본·원본의
`수집논문_PDF`·`txt` 유무가 달라도 통합 스크립트는 영향을 받지 않는다). 사본의 데이터 산출물을 그대로
복사해도 결과는 같다(같은 입력·같은 스크립트).

## 1. 신규 스크립트

### `_claude/integrate_newft_into_corpus.py` (신규 · 800줄)
- 입력: 판정 세트 폴더 여러 개(`fulltext/newft_final[_<날짜>]/{verdict,extract,mmat,detail,es}.csv`) + 서지
  `fulltext/newft_manifest_*.csv` + 8월 판정 변경 상수 `CORPUS_CHANGES`(1226) + 선택 `--changes CSV`.
- 검사: branch↔id 형식, verdict/reason_code(X1–X7)/boundary_rule 값, 포함·SENS 의 extract·mmat·detail 유무,
  MMAT `n_yes`·`quality_tier` 재계산 일치, detail 5문항 item_no 가 MMAT 범주와 일치, 서지 대조.
- 멱등성: `fulltext/newft_integration_ledger.csv`(uid·내용 해시). 같은 내용 재실행 = 건너뜀,
  내용이 달라졌거나 원장에 없는 id 가 이미 정본에 있으면 **한 줄도 쓰지 않고 중단**(음성 테스트 3종 통과).
- 반영: DB `ft_verdicts_v2`·`ft_extraction_v2`·`quality_all`·`quality_detail_all` /
  CT `ct_verdicts_final`·`ct_extraction_final`·`qa_results/qa_ctfull_<세트>_{mmat,detail}.csv`(merge_quality_v2 의 glob) /
  OAS `oas_results_ft/oasft_<세트>_{verdict,extract}.csv`·`qa_results/qa_oas_{mmat,detail}.csv` /
  확보 상태 `ct_retrieval_final.state`·`oa_supp_retrieval.result` / `ruling_audit.csv`.
- 파생 표(매번 재생성): `ft_exclusion_reasons.csv`(전문 배제 120건의 X 코드·PRISMA 범주),
  `retrieval_status_all.csv`(전문 확보 대상 283건의 확보 상태·미확보 사유).
- `--fetch-authors`: 새 포함 연구의 저자 표기를 Crossref 로 조회해 `appendix_b_author_lookup.json`·
  `appendix_b_crossref_all.json` 에 추가(추측 금지 — Crossref 에 저자가 없는 1191 은 원문 첫 쪽 저자 줄을
  `AUTHOR_OVERRIDES` 에 근거와 함께 명시).
- `--downstream`: 지도 §3 순서 13개 + 검증 2개 실행, 실패 시 중단·로그 저장. `ma_*.py` 는 실행하지 않는다.
- X 코드 → PRISMA 배제 범주 매핑표(스크립트 상수 `X_CATEGORY`·`X_CATEGORY_BRANCH`):
  X1 Animal · X2 Setting not eligible · X3 No acoustic exposure(CT·OAS 표기는 No acoustic variable) ·
  X4 No observed behaviour(OAS 표기 No behavioural outcome) · X5 Not empirical · X6 Not a journal article(신설) ·
  X7 Not in English. 8월 DB 16건은 코드 기록이 없어 `AUG_DB_EXCLUSION` 으로 재배정(집계 7·6·1·1·1 그대로 재현).

## 2. 수정한 기존 스크립트 (7개, 전부 주석에 이유 표기 · 전문 diff = `scripts_20260918.patch`)

| 파일 | 변경 요지 |
|---|---|
| `finalize_oa_supp_branch.py` | ① `oasft_*_{verdict,extract}.csv` glob(세트별 파일) ② 확보 판정에 `ok-<날짜>` 포함 ③ 추출표 대조를 포함∪민감도로(OAS 에 SENS 가 처음 생김) ④ 흐름 블록에 배제 사유(X 코드)·민감도 줄 추가 ⑤ "전문 확보 5편" → "8월 전문 확보 5편" |
| `fig_data.py` | PRISMA 전문 단계 하드코딩(`sought·not_retrieved·nr·assessed·ft_excluded·ftx·sens·included`) 제거 → `corpus_v4_verdicts.csv` + `ft_exclusion_reasons.csv` + `retrieval_status_all.csv` 에서 계산. 식별·스크리닝 칸(2,073·1,316·428 등)은 이번 작업과 무관해 유지. 정본↔파생 표 불일치 시 중단하는 검사 3종 추가 |
| `make_fig1_prisma_v2.py` | 기타방법 열의 민감도 = 인용추적 + 보조검색(`SP.get("sens", 0)`) — OAS0005 반영 |
| `rebuild_table1.py` | ① 통합 원장의 새 연구는 **구 Table 1(84행) 매칭 대상에서 제외**(남의 국가·설계·n 을 빼앗는 사고 차단) ② 새 연구의 setting·design 은 추출 통제어휘 머리말로 코딩(자유서술 전체 정규식은 15편 중 12편 오코딩) ③ 민감도 구성 문구를 경계 규칙 기록에서 계산 ④ 매칭 분모 출력 |
| `build_appendix.py` | 포함 편수 기대값 리터럴 98 → `corpus_v4_verdicts.csv` 의 FINAL_INCLUDE 수와 대조 |
| `verify_manuscript.py` | 필수 수치(98·81·15·2편·19%·40%)를 정본 산출값으로 대체 |
| `verify_consistency_v2.py` | 기대값 98·4 → 108·8, 81·15·2 → 90·16·2, `prisma_flow.md` 기대 문자열 98·15 → 108·16 |

## 3. 데이터 파일 변경 (사본 기준 행 수)

| 파일 | 행 | 내용 |
|---|---|---|
| `fulltext/ft_verdicts_v2.csv` | 100 → 167 (+67) | DB 새 판정 67행 추가 · 1226 `FINAL_INCLUDE→FINAL_EXCLUDE`, reason 갱신 |
| `fulltext/ft_extraction_v2.csv` | 84 → 95 (+11) | 새 포함·SENS 12행 추가, 1226 1행 제거 |
| `fulltext/quality_all.csv` | 84 → 95 (+11) | 〃 (mmat_category 는 기존 Title case 관례) |
| `fulltext/quality_detail_all.csv` | 420 → 475 (+55) | 새 12편 × 5문항 추가(item 열 = `"<item_no> <영문 문항>"` 기존 관례), 1226 5행 제거 |
| `fulltext/ct_verdicts_final.csv` | 54 → 60 (+6) | CT 6건(batch = nft_XX, screen1 = manifest) |
| `fulltext/ct_extraction_final.csv` | 16 → 18 (+2) | CT0223(포함)·CT0142(SENS) |
| `fulltext/qa_results/qa_ctfull_newft_{mmat,detail}.csv` | 신규 2·10행 | CT 품질(기존 `qa_ctfull_*` glob 에 잡힌다) |
| `fulltext/oas_results_ft/oasft_newft_{verdict,extract}.csv` | 신규 4·1행 | OAS 판정 4건·추출 1건(OAS0005) |
| `fulltext/qa_results/qa_oas_{mmat,detail}.csv` | 2→3 · 10→15 | OAS0005 품질 |
| `fulltext/ct_retrieval_final.csv` | 79 (행 수 동일) | CT0017·0034·0108·0142·0190·0223 `state = retrieved-20260918`, reason 에 8월 미확보 사유 보존 |
| `fulltext/oa_supp_retrieval.csv` | 15 (행 수 동일) | OAS0001·0005·0040·0222 `result = ok-20260918`(8월 사유는 reason 에 보존) |
| `fulltext/ruling_audit.csv` | 9 → 10 | 1226 R1 재판정 기록(old FINAL_INCLUDE → new FINAL_EXCLUDE) |
| `fulltext/newft_integration_ledger.csv` | 신규 78행 | 멱등성 원장(77건 + 1226) |
| `fulltext/ft_exclusion_reasons.csv` | 신규 120행 | 전문 배제 DB 72 · CT 42 · OAS 6 의 X 코드·PRISMA 범주·근거 |
| `fulltext/retrieval_status_all.csv` | 신규 283행 | DB 189 · CT 79 · OAS 15 의 확보 상태·확보 회차·미확보 사유 |
| `fulltext/appendix_b_author_lookup.json`·`appendix_b_crossref_all.json` | +11 | 새 포함 연구 11편 저자 표기(Crossref, 1191 만 원문 확인) |

다운스트림이 다시 쓴 산출물: `corpus_v3_*`·`corpus_v4_*`·`quality_v2`·`quality_detail_v2`·`table1_v2.{csv,md}`·
`evidence_counts_v2`·`evidence_map_v2.md`·`geo_time_counts`·`table1_en.md`·`table3_en.md`·`appendix_ab.md`·
`manuscript_facts.md`·`oa_supp_summary.md`·`corpus_v3_summary.md`·`figures/fig_data.json`·
`figures/Fig1·3·4·5·7·8.{png,pdf}`(Fig2 는 MA 불변이라 내용 동일, Fig6 는 수치를 쓰지 않아 미재생성).

## 4. 반영 전후 핵심 수치

| 항목 | 전 | 후 |
|---|---|---|
| 포함(FINAL_INCLUDE) | 98 (DB 81 · CT 15 · OAS 2) | **108** (DB 90 · CT 16 · OAS 2) |
| 민감도 전용 | 4 (DB 3 · CT 1) | **8** (DB 5 · CT 2 · OAS 1) |
| 분석 대상 | 102 | **116** |
| 전문 평가 | DB 100 · CT 54 · OAS 5 | DB 167 · CT 60 · OAS 9 |
| 전문 배제 | DB 16 · CT 38 · OAS 3 | DB 72 · CT 42 · OAS 6 |
| 미확보 | DB 89 · CT 25 · OAS 8 | DB 22 · CT 19 · OAS 4 |
| 방향(도메인 레코드) | fwd 113 · rev 74 · both 25 (역방향 40%) | fwd 119 · rev 87 · both 31 (역방향 **42%**) |
| 측정 세대 | G1 54 · G2 42 · G3 24 · 병용 22 | G1 62 · G2 45 · G3 28 · 병용 26 |
| MMAT(포함분) | high 22 · moderate 40 · low 36 | high 24 · moderate 43 · low 41 |
| 최다국 | China 38/98 (39%) | China 45/108 (**42%**) |
| 인용추적 기여율 | 15/81 = 19% | 16/90 = **18%** |

## 5. 아직 하지 않은 것(원고·MA·문서)

- **메타분석**: `ma/*_input.csv`·`ma_v2.py`·민감도 스크립트 미변경(범위 밖). 후보 목록 = `integ/ma_candidates_newft.csv`.
- **원고·수기 문서**: `01_논문작업/Manuscript_KO.md`(98·81·15·102·89 등), `fulltext/prisma_flow.md`,
  `01_논문작업/Figure.pptx`(사용자 작도 PRISMA) 는 갱신하지 않았다 → `verify_*` 실패는 이 때문이다.
- `build_references.py`·`build_numbered_refs.py`(포함 연구 참고문헌 목록, 1226 은 본문 인용 중), `build_docx.py`,
  `verify_design_matrix_v2.py`, `design_matrix_v2.csv`(계획 레버 근거 편수)는 실행하지 않았다.
