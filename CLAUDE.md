# Research — 연구·발표·제출 자료 (다기기 실시간 동기 repo)

논문·발표·제출 자료의 단일 논리 repo. **Syncthing `research` 폴더(`.git` 포함)가 정본 채널** — 맥미니·맥북·그램 3기기 실시간 동기(실측 ~4초 전파), desktop(새PC)은 macmini 쪽 공유 추가 완료(2026-07-12) 상태로 전원 on 시 수락만 하면 편입. git origin push = 이력·모바일 열람·오프사이트용(세션 단위 수동). 기계 상세 정본 = `~/.claude/knowledge/sync_mechanics.md §research`.

## 세션 규칙 (모든 기기 공통 · 필수)

- **동시 편집 금지** — 단일 사용자 순차 작업 전제. 다른 기기에서 방금 저장/커밋했다면 전파(수 초) 완료 후 git write. 두 기기 동시 저장 = `.sync-conflict` 사본·index 경합.
- git `MM`/`needs update` 표시(기기 간 `.git/index` stat 캐시 불일치) = 콘텐츠 무결 → **`git reset -q`(무손실)로 해소**.
- **push는 세션 단위 수동**(아무 기기서나 — Mac이면 ssh키로 인증 불요). ⛔ `cross_machine_sync`·gram `auto_pull` 대상 아님(`.git` 경합).
- `*.sync-conflict*` 발견 즉시 내용 대조 후 해소(gitignore 차단됨) · versioning=trashcan 7일(실수 삭제 보호).
- Windows(그램·desktop) 편집기는 **LF 유지**(CRLF 저장 시 Mac 쪽 dirty 노이즈) · 공유 `.git/config`의 `core.filemode=false` 유지.

## 멀티모델 (기기 게이팅 — 정본 `~/.claude/rules/multi-model.md`)

- macmini = 직접 실행 허브. **그램·맥북·desktop 세션은 `multimodel_request` MCP / `mm_request.py` broker 왕복**으로 Codex·ChatGPT·Gemini 이용(로컬 토큰 복제 금지). ~50초 미완이면 `job_id` 반환(에러 아님) → `multimodel_result(job_id)` 반복 회수.
- ★ 이 repo는 macmini에도 실시간 동기 — 요청 패킷에 macmini 경로(`/Users/hyunbin/Research/...`)로 파일을 지목할 수 있다(Windows 경로 그대로 던지지 말고 변환). 단 방금 저장한 파일은 전파 수 초 대기 후. 핵심 발췌 동봉이 더 안전.
- 논문·학술 작업 = `/paper-write` 워크플로우(Opus 작성 → Codex 방법론/수식/재현성 검증 → Gemini ∥ ChatGPT Deep Research 병렬 보강).

## 핸드오프·산출물

- 이 repo 작업의 핸드오프 = `_internal/handoff/{active,archive}/`(repo 내 → 전 기기 자동 전파·git 추적). 구 그램 `T:\claude_research` 기록 = sub/백업.
- 산출물 네이밍 = 타임코드 `<문서명>_YYYYMMDD_HHMMSS.ext`(`v넘버` 분기자 금지) · 사용자 열람용 md는 md2pdf PDF 쌍 원칙(`~/.claude/rules/hygiene.md`).
