# 새 포함 연구 추가용 파이프라인 지도 (2026-09-17, Explore 에이전트 조사 · 읽기 전용)

목적: DB 경로 전문 미확보 89건 중 뒤늦게 확보한 원문을 전문 평가해 코퍼스에 넣을 때의 절차·위험.

## 1. 정본
- `fulltext/corpus_v4_verdicts.csv`(uid, source, final_verdict, orig_id; 159행: 포함 98 = DB 81·CT 15·OAS 2, SENS 4, 배제 57)
- `fulltext/corpus_v4_extraction.csv`(uid, source + 16열; 102행)
- uid: DB = 스크리닝 풀 `no` 정수, CT = `CT####`, OA 보조 = `OAS####`
- 체인: DB `merge_fulltext_results.py` → `apply_boundary_rulings.py` → `merge_final_corpus.py` → `ft_verdicts_v2.csv`(100)·`ft_extraction_v2.csv`(84)·`effect_sizes_all.csv`
  / CT `merge_ct_fulltext.py` → `apply_ct_corrections.py` → `ct_*_final` / 통합 `merge_corpus_v3.py` → `finalize_oa_supp_branch.py` → `corpus_v4_*`
- **권장 추가 경로**: 새 DB 행을 `ft_verdicts_v2.csv`·`ft_extraction_v2.csv`·`quality_all.csv`·`quality_detail_all.csv` 에 직접 추가 → `merge_corpus_v3.py` → `finalize_oa_supp_branch.py`(단 `oa_supp_summary.md` 덮어씀·수치 문자열 박힘 :93-111)
- ⛔ 재실행 금지: `build_fulltext_packets.py`(fulltext/pdf 비어 있음 → packet_index 0행 덮어씀), `merge_fulltext_results.py`(배치명 고정, ft_*_all 덮어씀), `merge_quality.py`(OAS uid 에서 int 변환 실패·item_no 없는 4열 출력)

## 2. 새 연구 필드
| 파일 | 열 | 허용값 |
|---|---|---|
| ft_verdicts_v2.csv | no, final_verdict(FINAL_INCLUDE/SENS_ONLY/FINAL_EXCLUDE), confidence(high/medium/low), reason(한글 자유서술), screening_verdict, year, journal, title | |
| ft_extraction_v2.csv | country(영문 `;`, NR), setting(street/park/square/campus/residential/recreation/waterfront/VR-lab/mixed + 괄호), design(survey/mixed/field-experiment/natural-experiment/quasi-experiment/field-observation/observational/sensor-bigdata/lab-VR-experiment/qualitative/NR), sample_n, exposure(음원 범주는 여기+제목 정규식), behaviour_measure, key_finding, behaviour_domain(movement/staying/space-use/social/activity `; `), measurement_method(self-report/observation/sensor-GPS-video + 괄호 → G1/G2/G3 정규식), direction(forward/reverse/both), effect_stats(원문 인용 또는 NR) | |
| quality_all.csv | no, mmat_category(Qualitative/Quantitative RCT/Quantitative non-randomised/Quantitative descriptive/Mixed methods), S1, S2, Q1–Q5(Y/N/CT), n_yes, quality_tier(≥4 high·3 moderate·≤2 low), note, year, journal, title | |
| quality_detail_all.csv | no, item, item_no(예 4.1, 필수), verdict, rationale — 연구당 5행 | |
| appendix_b_author_lookup.json | {uid: {label, year, doi, title}} — 없으면 build_appendix.py:96 중단 | |
| 선택 | effect_sizes_all.csv(no, cluster, outcome_measure, comparison, statistic_type, values_verbatim, n_info, location, quote, computable), pdf_map.csv, txt/TXT_<no>.txt, packet_index.csv | |

## 3. 다운스트림 순서
1 merge_corpus_v3 → finalize_oa_supp_branch · 2 merge_quality_v2 · 3 rebuild_table1(table1_v2) · 4 rebuild_evidence_map(evidence_counts_v2) · 5 make_fig7_geo_time ·
6 ma_v2 → ma_sensitivity_v2, ma_supplementary_analyses · 7 fig_data.py(**PRISMA 하드코딩 :171-211**) · 8 make_figures, make_fig1_prisma_v2, make_fig8_quality ·
9 build_ma_char_table / build_ma_sensitivity_table · 10 build_tables_en(Table 1·4, design_matrix_v2 수기) · 11 build_appendix · 12 manuscript_facts ·
13 Manuscript_KO.md 수기 반영 + Figure.pptx(Fig.1) · 14 verify_manuscript, verify_consistency_v2 → build_docx
- 그림 파일명 ≠ 원고 번호: Fig7_GeoTime = Fig.2, Fig8_Quality = Fig.3, Fig6_Framework = Fig.8

## 4. 메타분석 입력
- walking: ma_walking_speed.py STUDIES 수기 / staying·social: ma_all_clusters.py 하드코딩 / correlation: 08-06 수기 편집(재생성 금지)
- CT 효과(CT0025·CT0126·CT0184)는 ma_v2.py:131-161 하드코딩
- 새 효과 추가 = ma/ma_*_input.csv 행 직접 추가 + FOREST_LABEL(make_figures.py:39-59), N_ANALYTIC·CONTRAST(build_ma_char_table.py:35-53), MA dict(rebuild_table1.py:99-105), INTEXT(build_references.py:17-36), 민감도 uid 하드코딩(ma_sensitivity_v2.py:56,99-106,177,185,187; ma_supplementary_analyses.py:93-96,150), build_ma_sensitivity_table.py:141-145
- 규약: fulltext/analysis_rules.md §2 부호 §3 1효과 §4 참가자 n §5 변환 불가

## 5. 판정 방식
- DB 갈래 지시문 파일 없음(결과 CSV 헤더로 역산). 배제 사유는 자유서술, PRISMA 5분류는 fig_data.py:182-184 수기
- R1–R3(DB, UNCERTAIN 에만 키워드 적용) / P1–P3(CT; P2는 R2가 정본)
- **새 연구 판정·추출·품질 양식으로는 OA 보조 갈래 스키마(oas_results_ft/oasft_01_*, qa_oas_* item_no 포함)가 가장 완전**
- CT 배제 코드 X1 동물 / X2 세팅 / X3 음환경 없음 / X4 행태 없음 / X5 비실증 / X6 중복·철회 / X7 언어

## 6. 하드코딩 수정 목록
fig_data.py:178-186 · build_appendix.py:87(98) · verify_consistency_v2.py:34,37-38,88-97(옛값, 지금도 실패 추정),102-103 · verify_manuscript.py:115-116(98·81·15·2편·19%·40%) ·
rebuild_table1.py:36,99-105,168,172 · make_figures.py:43-58 · build_ma_char_table.py:36-52 · build_references.py:18-36 · ma_v2.py:131-188 · ma_sensitivity_v2.py · ma_supplementary_analyses.py ·
prisma_flow.md · Manuscript_KO.md(98: 줄 5,29,53,131,137,143,190,196,289,380 / 81·84·100·189·102: 131 / 89: 372 / Table 1 145-188 / B1 554) · Figure.pptx
- ⚠️ rebuild_table1.py 는 DB 행을 구 Table 1(84행)에 (year, direction, domain) 키로 매칭, 후보 복수면 첫 후보 → 새 행이 기존 연구 값을 빼앗을 수 있음. 실행 후 매칭 수·table1_v2 diff 확인 필수.
