# paper32 — 미확보 전문 추가 확보·평가 작업 기록 (2026-09-18 시작)

> 상태: **진행 중** (00:10 KST 기준 전문 평가 1차 6배치 실행)

## 배경·사용자 결정
- 2026-09-17 최종 검토에서 DB 경로 전문 미확보 89건(대부분 BORDERLINE) 중 49건이 무료 원문이 있음을 확인.
- 8월에 BORDERLINE 을 "연관성 낮음"으로 보고 확보 요청을 안 한 것은 Claude 판단이었고, 근거(생존율 낮음)는 자료와 달랐음(22건 중 12건 포함).
- 사용자: "이러면 추가해야지. 논문 자체가 설득력이 떨어지는건데" → 확보·평가·코퍼스 반영 결정. 표·그림 재생성 수용.
- OSF 등록본 갱신은 이 작업이 끝난 뒤 한 번에(문안 초안 `review/osf_update_draft_20260917_234900.md`, 수치 갱신 필요).
- 10% 재스크리닝은 사용자 시간 부족으로 미수행 → 편차로 보고(시트 `screening/rescreen10/` 는 보존).

## 확보 현황 (01:00 KST)
- 스크립트: `fetch_db_oa_fulltext.py`(자동 OA) → `ingest_manual_pdfs.py`(사용자 다운로드 대조·명명·이동 + 미확보 목록 + 경로별 대조표)
- 전문 확보 대상 283건 = DB 189 · CT 79 · OAS 15 / 누락 0
- 추가 확보: DB 56 PDF + 11 XML · CT 6 · OAS 4 = **77건** → 평가 중
- 미확보 45건(DB 22 · CT 19 · OAS 4) → 사용자 9/18 회사에서 시도. 목록 `01_논문작업/미확보_원문_목록_20260918_010058.xlsx`
- 라이브러리 `수집논문_PDF` 224 파일 = DB 156 · CT 61(60건, CT0018 두 파일) · OAS 4 · 참고 논문 3

## 전문 평가
- 텍스트: `build_newft_texts.py` → `fulltext/txt/TXT_<id>.txt`(자르지 않음) · manifest `fulltext/newft_manifest_20260918.csv` · 배치 `newft_batches_20260918.json`(18배치)
- 지시문: `fulltext/newft_instructions_20260918.md` (8월 기준 + R1–R3·P1·P3, X1–X7, OAS 양식, MMAT item_no, 효과 후보)
- 결과: `fulltext/newft_results/nft_XX_{verdict,extract,mmat,detail,es}.csv`
- 검사·통합: `validate_newft_results.py [--partial]` → `fulltext/newft_combined/`
- 실행: nft_01–06 시작(00:1x), 07–12, 13–18 순차(동시 6)

## 다음 단계
1. 18배치 완료 → validate → 오류 배치 재요청
2. 독립 점검: 포함·경계(low confidence) 판정을 fresh 에이전트로 재확인
3. 사용자 확인용 판정표(한 장) 작성 → 사용자 읽고 동의
4. 9/18 사용자 추가 다운로드분 ingest → 같은 절차로 평가
5. 코퍼스 반영(pipeline_map_add_studies_20260917.md 순서: ft_*_v2·quality_all·detail 행 추가 → merge_corpus_v3 → finalize_oa_supp_branch → merge_quality_v2 → rebuild_table1 … ) + 하드코딩 수치 갱신 + MA 클러스터 후보 규칙 적용
6. 원고 수치·표·그림·PRISMA(Figure.pptx) 갱신 → verify → build → 독립 게이트
7. 그다음 원래 검토 반영(주장 강도·2.9 편차·한계) → OSF 한 번 갱신 → 커밋

## 평가 중 발견한 기존 코퍼스 문제 (나중에 일괄 점검)
- [nft_02] 8월 코퍼스 **1226**(불가리아, 거주지 Lden × 설문 신체활동)은 FINAL_INCLUDE 인데, 같은 구조인 779·880·1141 은 R1 로 배제(`ruling_audit.csv`). 1226 판정 재검토 필요(R1 적용 시 배제 → 편수·표 영향).
- [nft_03] 1215(Data in Brief 데이터 논문, X5)가 같은 자료를 분석한 **Korpilo et al. 2025, Urban Forestry & Urban Greening 113:129088** 을 언급 — 체계적 검색 밖에서 알게 된 후보. 포함하려면 PRISMA "기타 출처"로 기록해야 함(결정 필요).
- [nft_06·07 불일치] 같은 주제(운전 행태 → 차량 근거리 엔진 소음) 두 편이 반대로 판정됨: **186**(Calvo 2012) FINAL_INCLUDE(low, 경적 선례 799 근거) vs **194**(Ibarra 2012) FINAL_EXCLUDE X3. 등록본은 driving behaviour 배제·공공공간 음환경 경로 필요 → 재확인 단계에서 한쪽으로 통일.
- [nft_06] 562(MTurk 컨조인트, 가상 경로 의도) SENS_ONLY(R2) — es 후보(참가자 내 비교) 기록. 713(홍콩 공원, 방문빈도→선호, R3) 포함 low 품질.
- 재확인 대상(low/medium 경계): 1190·268(nft_04), 624(nft_07), 186·623·562·713(nft_06).
- 재확인 대상 추가: OAS0005(SENS P1 low, nft_08) · 1305(포함 low, 음 변수 = 이상적 녹지 속성 중요도, nft_10) · CT0142(SENS R2 low, 지불의사) · 1143(X4, 국립공원 사운드워크) · CT0223(포함 R3 low, 계획 체류시간→자연 소리 만족, nft_11) · 972(X4 medium)
- 재확인 대상 추가: 1191(포함 low — 행태는 방법 절 기술·논의 서술뿐, 결과 절에 행태 자료 없음, nft_15) · 888(포함 R3 medium, 소수 코드 의존, CT0184 저자·현장 동일) · 639(X4 medium)
- [nft_14] **표본 공유 의심**: 104(Yu & Kang 2009 JASA, 포함 R3)가 코퍼스 **123**(Yu & Kang 2010, 같은 19개 사이트)과 같은 설문 자료일 가능성. nft_16 의 80(Yu & Kang 2008)도 같은 자료일 수 있음 → 셋을 대조해 표본 공유면 "서술 종합에는 각각 유지, 같은 MA 클러스터엔 한 편"(8월 D5-5 규칙). 499(지난 Forest Park, 방문빈도→음원 선호 R3) 포함 medium.
- [nft_17] 709(Muir Woods "quiet" 안내판 교대 현장실험, 포함 R3 high, both) — 보행속도 조건 비교 es(± 가 SE 인지 불명) → 대비가 "음환경 조건"이 아니라 "행태 개입 조건"이라 walking_speed 클러스터 적합성 검토 필요. 미확보 DB 89(Park Science 2009, 같은 Muir Woods 프로젝트)와 표본 공유 가능성.
- 재확인 대상 추가: 55(X2 low, 주택 정원 체류 — 코퍼스 71 과 같은 연구진) · 1013(X4 medium)

## 1차 평가 완료 (01:24 KST)
- 77건: 포함 13(DB 12·CT 1) · SENS 4 · 배제 60 — `validate_newft_results.py` 전체 통과, 통합 `fulltext/newft_combined/`
- 독립 재판정(블라인드) 36건 = 포함·SENS 17 + 경계 배제 19 → `newft_recheck_batches_20260918.json`(rc_01–07), 결과 `fulltext/newft_recheck/rc_XX_verdict.csv`
  rc_01–06 실행(01:25), rc_07 대기(444·823·562·499·559)
- 불일치 → 판정 조정(원문 근거) → 확정표 → 사용자 확인

## 판정 확정 (02:0x KST)
- `finalize_newft_verdicts.py` → `fulltext/newft_final/{verdict,extract,mmat,detail,es}.csv` + `adjudication_log.md`
- 재판정 일치 33/36 (91.7%), κ=0.83. 조정: 186→배제 X3(194와 일관), 624→배제 X4, 1191→포함(행태 측정했으나 결과 수치 미보고 = 보고 결함)
- **최종: 포함 11**(DB 80·104·499·709·713·888·1190·1191·1218·1305 + CT0223) · **SENS 4**(562·1172·CT0142·OAS0005) · 배제 62
- 사용자 확인표: `01_논문작업/추가전문_판정확인_20260918.xlsx` (판정 확인 시트 + 요약·결정 시트)
- 사용자 결정 대기(권장안 제시): ① 8월 포함 1226 → R1 배제 ② 80·104·123 표본 공유 → 서술 각각·MA 한 편 ③ Korpilo 2025 추가 안 함
- ⏸ 대기: 사용자 판정 확인·동의 + 9/18 회사에서 받을 미확보 45건 → 받은 분 같은 절차 평가 → 코퍼스 일괄 반영(다음 단계 5)
- ✅ **사용자 확인(2026-09-18)**: "판정 근거들 읽어봤는데 적절하게 선정된거같아" — 77건 판정 동의. 결정 3건은 사전 위임("너 나름대로 알아서")에 따라 권장안 채택: 1226 R1 배제 · 80/104/123 서술 각각·MA 한 편 · Korpilo 2025 미추가.
- 체크포인트 커밋 `be21ef6` (확보·판정 스크립트·결과·확인표·검토 보고서)
- 통합 스크립트 드라이런 진행 중: 사본 `scratchpad/integ/`(git archive) 에서 `integrate_newft_into_corpus.py` 작성·전체 다운스트림 재생성 → `scratchpad/integ/CHANGES.md` 로 원본 반영 패치 요약. 기대값: 포함 108(DB 90·CT 16·OAS 2) · SENS 8 · 분석 116. MA 입력은 범위 밖(후보 목록만).

## 통합 드라이런 완료 (사본, 03:1x KST) — 산출물 `_claude/pending_integration/`
- `integrate_newft_into_corpus.py`(신규·멱등·원장 `newft_integration_ledger.csv`) + 수정 7개 스크립트 패치 `scripts_20260918.patch` + `CHANGES.md`(원본 반영 절차) + `ma_candidates_newft.csv`
- 다운스트림 13단계 완주. 반영 후: 포함 108(DB 90·CT 16·OAS 2) · SENS 8 · 분석 116 · 역방향 레코드 42% · MMAT high 24·mod 43·low 41 · China 45(42%)
- PRISMA 계산값 = 기대값 일치. 미확보 사유 재작성(구독 전용 11·DOI 없음 8·OA 실패 3 / CT 17·2 / OAS 3·1) — "No institution access 84" 폐기
- Table 1 매칭 위험: 새 행을 구 Table 1 매칭에서 제외하도록 수정(실측 피해 0, 예방적)

### 결정(내가 정함, 필요 시 사용자 재검토)
1. 8월 DB 배제 16건 사유 코드 = 사유 서술대로 **정확히 재배정**(909 → 비실증). 원 집계(7·6·1·1·1) 보존보다 정확성 우선 — 편차 기록에 남긴다.
2. 부록 B 연도 = **코퍼스 연도로 통일**(Crossref issued 로 인한 CT0223 2019→2020 등 표기 불일치 제거).
3. 기존 CT·OAS 행의 setting·design 정규식 오코딩(CT0090·CT0137·CT0356·CT0414) = **소급 통일**(추출 통제어휘 머리말 코딩).
4. 1226 배제 → 본문 인용·참고문헌(`build_references.py` INTEXT, `build_numbered_refs.py`) 손질은 원고 갱신 단계에서.

### ⚠️ 확인된 도구 결함(별도 수정 필요)
- `verify_manuscript.py` 의 수치 검사는 "그 숫자가 원고 어딘가에 있는가"만 본다 → 참고문헌 쪽수·인용번호와 우연 일치로 통과. 원고 갱신 판정에 쓸 수 없다. 절·문장 맥락을 요구하는 검사로 교체 필요.
