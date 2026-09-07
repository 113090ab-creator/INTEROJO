# 오늘 작업 로그 (2026-09-05)

## 1) Validated Snapshot Set 적용
- 작업 시각: `2026-09-05 10:49:56 +09:00`
- 대상 저장소: `113090ab-creator/INTEROJO`
- 기준 브랜치: `origin/main`
- 반영 커밋:
  - `fd0cd2e7f8809e732d91b8e00892fb5035a3ef8c Implement validated APS snapshot sets`

## 2) 작업 전 local main 상태 처리
- 작업 전 local `main`은 `origin/main`보다 2개 커밋 ahead 상태였음.
- 확인한 ahead 커밋:
  - `9ca00f4 Merge remote APS snapshot status`
  - `d253d81 Refresh APS snapshots`
- `git diff origin/main..main --stat` 확인 결과 `cloud_snapshots/*` 진단용 snapshot 변경만 포함되어 있었음.
- 해당 진단용 snapshot 커밋은 push하지 않았음.
- `origin/main` 기준으로 `codex/validated-snapshot-set` branch를 만들어 코드 변경만 별도 반영함.

## 3) Validated Snapshot Set 핵심 변경
- PLAN / WIP / 생산부족 결과를 독립 publish하지 않고 하나의 검증 완료 set 단위로 관리하도록 변경함.
- set 구조:
  - `cloud_snapshots/sets/{set_id}/manifest.json`
  - `cloud_snapshots/sets/{set_id}/wip_inventory_snapshot.csv.gz`
  - `cloud_snapshots/sets/{set_id}/aps_plan_operations*.csv.gz`
  - `cloud_snapshots/sets/{set_id}/shortage_snapshot*.csv.gz`
  - `cloud_snapshots/sets/{set_id}/shortage_file_info*.csv.gz`
  - `cloud_snapshots/sets/{set_id}/process_map*.csv.gz`
  - `cloud_snapshots/current_snapshot_set.json`
- manifest에는 `set_id`, `slot_key`, `plan_updated_at`, `wip_updated_at`, `created_at`, source label, checksum, row count, validation result를 기록함.

## 4) stale WIP 재조회 버그 수정
- 기존 문제:
  - snapshot refresh 중 새 PLAN을 가져온 뒤 생산부족 계산 과정에서 저장소의 기존 published WIP snapshot을 다시 읽을 수 있었음.
  - 이 경우 `PLAN=오늘 오전`, `WIP=어제 오후` 조합으로 계산될 위험이 있었음.
- 수정:
  - `build_api_shortage_data_from_frames()`에 `inventory_df`, `inventory_source_label`, `plan_updated_at` 인자를 추가함.
  - refresh build 중에는 이번 실행에서 fetch/validate한 staged WIP를 직접 계산 함수에 전달함.
  - 계산 중 기존 published WIP를 다시 읽지 않도록 테스트로 고정함.

## 5) Slot 판정 및 상태 전이
- slot은 날짜 + AM/PM 기준으로 판정함.
- 예:
  - `2026-09-05 08:01:23` -> `2026-09-05 AM`
  - `2026-09-05 16:10:53` -> `2026-09-05 PM`
- PLAN slot과 WIP slot이 같을 때만 `READY`.
- PLAN이 더 최신 회차이면 `WAITING_FOR_WIP`.
- WIP가 더 최신 회차이면 `WAITING_FOR_PLAN`.
- 정상적인 PLAN/WIP 전달 시차는 `FAILED`가 아니라 `WAITING_FOR_PLAN` 또는 `WAITING_FOR_WIP`로 관리함.
- 허용 대기 시간을 넘으면 `DELAYED`.
- 실제 API/schema/검증/저장 실패는 `FAILED`.

## 6) Atomic Publish
- 새 set publish 순서:
  1. staged PLAN 확보
  2. staged WIP 확보
  3. 같은 slot 검증
  4. 생산부족 계산
  5. 모든 site 결과 생성
  6. 필수 컬럼 검증
  7. row count guard 검증
  8. checksum/manifest 생성
  9. set 파일 저장
  10. `current_snapshot_set.json` 포인터를 마지막에 갱신
- 어느 단계에서 실패해도 기존 `current_snapshot_set.json`은 변경하지 않음.
- GitHub Actions refresh 실패 시 partial snapshot 파일을 commit하지 않고 status 파일만 commit하도록 보호함.

## 7) GitHub Actions 변경
- 기존 장시간 sleep loop를 제거함.
- 각 run은 짧게 실행:
  - meta 확인
  - 준비 안 됨: WAITING/DELAYED 기록 후 종료
  - 준비됨: build/validate/publish
- KST 오전/오후 갱신 구간에 5분 단위 cron으로 재확인하도록 변경함.
- 무한 재귀 `workflow_dispatch` 구조는 사용하지 않음.

## 8) Streamlit 표시 변경
- refresh 상태와 현재 표시 데이터인 published set을 분리함.
- WAITING/DELAYED/FAILED 중에도 기존 validated set 전체를 유지해 표시함.
- 현장 표시 상태:
  - 정상: `최신`
  - 정상 갱신 중: `갱신 중`
  - 허용시간 초과: `갱신 지연`
  - 실제 현재 장애: `갱신 실패`
- 내부 상태값 `PUBLISHED` 또는 `published`를 현장 화면에 그대로 노출하지 않고 `최신`으로 표시함.

## 9) 실제 publish 결과
- GitHub Actions에서 Validated Snapshot Set publish 정상 확인:
  - `set=aps_20260905_am_plan080123_wip081516_5b7800e6`
  - `slot=2026-09-05 AM`
  - PLAN: `2026-09-05 08:01:23`
  - WIP: `2026-09-05 08:15:16`
- Streamlit 실제 표시 기준:
  - 수요: `2026-09-05 08:01:23`
  - WIP: `2026-09-05 08:15:16`
  - 상태: `published`로 확인됨

## 10) stale FAILED 상태 표시 수정
- 문제:
  - Validated Snapshot Set publish는 성공했지만 Streamlit 상단에 과거 `FAILED` 메시지가 남음.
  - 표시된 과거 실패:
    - 마지막 확인: `2026-09-05 09:59:15`
    - 이후 정상 publish: `2026-09-05 10:39:13`
- 수정 원칙:
  - `current_snapshot_set.published_at` 또는 manifest 생성시각과 `refresh_status.checked_at/published_at`을 비교함.
  - `refresh_status`가 현재 published set보다 과거 이벤트이면 현재 장애로 보지 않고 무시함.
  - 현재 set이 정상 publish된 상태이면 `최신`으로 표시함.
- 보강:
  - publish 성공/skip 시 `refresh_status`에 `PUBLISHED`, `set_id`, `slot_key`, `published_at`을 남기도록 확인 및 보강함.
- 반영 커밋:
  - `a8f9255c3f7b397202d784f1de401d2546b40ea0 Fix snapshot status display priority`

## 11) 검증 결과
- Validated Snapshot Set 적용 검증:
  - `python -m py_compile app.py scripts\refresh_snapshot.py tests\test_validated_snapshot_set.py`
  - `python -m unittest tests.test_validated_snapshot_set`
  - `python scripts\refresh_snapshot.py --validate-existing --sites "C관,A관,S관,전체"`
  - `git diff --check`
- 실제 API dry-run:
  - PLAN `2026-09-05 08:01:23`
  - WIP `2026-09-05 08:15:16`
  - 같은 AM slot
  - publish 없이 staged 계산 검증 완료
- stale status 표시 수정 검증:
  - A. 과거 `FAILED` 이후 `PUBLISHED` -> `최신` 표시 통과
  - B. 현재 `PUBLISHED` 이후 `WAITING_FOR_WIP` -> `갱신 중`, current set 유지 통과
  - C. 현재 `PUBLISHED` 이후 실제 `FAILED` -> `갱신 실패`, current set 유지 통과
  - D. `PUBLISHED` 화면 표시 -> `published`가 아니라 `최신` 표시 통과
  - `python -m unittest tests.test_validated_snapshot_set` 결과: 10 tests OK
  - 실제 현재 파일 기준 확인:
    - `status=PUBLISHED`
    - `display=최신`
    - `sidebar=최신`
    - failure message 빈 값
