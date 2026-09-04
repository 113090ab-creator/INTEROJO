# 오늘 작업 로그 (2026-09-04)

## 1) 확인한 문제
- Streamlit 생산현황 화면에 오류 표시:
  - 스냅샷 기준시각: `2026-09-03 16:03:26`
  - APS API 최신 기준시각 표시값: `2026-09-03 16:03:26`
  - 원인 표시: 전일 스냅샷
- `2026-09-04 09:01 KST` 기준 GitHub Actions `Refresh APS snapshots`의 오늘 아침 예약 실행 기록이 없었음
- 원격 GitHub `cloud_snapshots`에는 `2026-09-03 16:26:22` 상태만 반영되어 있었음
- 로컬 PC에는 `2026-09-04 08:17:27` 기준 정상 스냅샷 변경분이 생성되어 있었지만, GitHub 원격에는 아직 올라가지 않은 상태였음

## 2) 원인 판단
- 자동 스냅샷 구조가 안정적으로 한 경로로 정리되지 않았음
- 기존 로컬 예약/실행은 오늘 아침 스냅샷을 정상 생성했지만, GitHub 원격으로 자동 반영되지 않았음
- 새 `Refresh APS snapshots` GitHub Actions workflow는 active 상태였지만, 오늘 아침 schedule run 자체가 생성되지 않았음
- 이전 구조는 5분마다 여러 GitHub Actions run을 만드는 방식이라 실패 상태와 스냅샷 충돌이 발생하기 쉬웠음
- APS PLAN은 `08:00:15`에 갱신됐고, APS WIP는 `08:14:45`에 갱신되어 WIP가 늦게 준비되는 패턴이 확인됨

## 3) 즉시 조치
- 로컬에 생성되어 있던 정상 아침 스냅샷을 검증함:
  - APS PLAN 기준시각: `2026-09-04 08:00:15`
  - APS WIP 기준시각: `2026-09-04 08:14:45`
  - WIP: `30,122`건
  - C관: `4,381`건
  - A관: `1,361`건
  - S관: `7,822`건
  - 전체: `13,564`건
- 검증 명령:
  - `python scripts\refresh_snapshot.py --validate-existing --sites C관,A관,S관,전체`
  - `python -m json.tool cloud_snapshots\aps_snapshot_refresh_status.json`
  - `python -m json.tool cloud_snapshots\aps_snapshot_refresh_state.json`
- 정상 스냅샷을 `cloud_snapshots`에 커밋하고 GitHub `main`에 푸시함
- 관련 커밋:
  - `e82884c Refresh APS snapshots for Sep 4 morning`

## 4) 예약 안정성 보완
- 기존 `Refresh APS snapshots` workflow는 유지:
  - 평일 `08:20 KST`
  - 평일 `16:30 KST`
  - 실패 시 workflow 내부에서 5분 간격으로 최대 7회 재시도
- 추가로 기존에 schedule 실행 이력이 안정적으로 있던 `Keep Streamlit Alive` workflow에 백업 갱신 job을 추가함
- 백업 갱신 기준:
  - 평일 `09:00 KST`
  - 평일 `17:00 KST`
  - `01:00 KST` keepalive 실행에서는 APS 스냅샷 갱신을 건너뜀
- 백업 갱신도 실행 내부에서 5분 간격으로 최대 7회 재시도함
- `Refresh APS snapshots`와 `Keep Streamlit Alive` 백업 갱신은 같은 concurrency group을 사용해 동시에 스냅샷 파일을 쓰지 않도록 함

## 5) 검증
- 문법 검증:
  - `python -m py_compile app.py scripts\refresh_snapshot.py snapshot_storage.py`
- 공백 검사:
  - `git diff --check`

## 6) 운영 메모
- 앞으로 기대 동작:
  - 08:20에 전용 workflow가 먼저 시도
  - 전용 workflow가 누락되거나 실패하면 09:00 keepalive 백업 job이 다시 시도
  - 16:30에 전용 workflow가 먼저 시도
  - 전용 workflow가 누락되거나 실패하면 17:00 keepalive 백업 job이 다시 시도
- GitHub schedule 이벤트 자체가 생성되지 않으면 해당 workflow 로그도 남지 않음
- 그래도 오늘처럼 로컬에만 스냅샷이 생기고 GitHub에 올라가지 않으면 Streamlit은 원격의 전일 스냅샷을 보게 됨
- 정확한 당일 반영이 급하면 GitHub `Actions > Refresh APS snapshots > Run workflow` 수동 실행으로 원격 스냅샷을 갱신함

## 7) WIP 지연 패턴 반영
- APS PLAN보다 APS WIP가 늦게 들어오는 패턴을 전제로 갱신 로직을 보강함
- 새 PLAN 기준시각이 확인돼도 WIP 기준시각이 PLAN보다 오래되면 스냅샷 생성을 중단하고 대기 상태로 기록하도록 변경함
- 이 상태에서는 기존 WIP와 새 PLAN이 섞인 스냅샷을 절대 저장하지 않음
- Streamlit 화면의 최신 스냅샷 대기 허용 시간을 15분에서 60분으로 늘림
- GitHub Actions 전용 갱신과 keepalive 백업 갱신 모두 5분 간격 최대 13회, 약 60분까지 재시도하도록 변경함
- 기존 상태 JSON에 남아 있던 로컬 `raw_dir` 경로 표시는 제거함
- 검증:
  - `python -m py_compile app.py scripts\refresh_snapshot.py scripts\refresh_aps_shortage_snapshots.py snapshot_storage.py`
  - `python scripts\refresh_snapshot.py --validate-existing --sites C관,A관,S관,전체`
  - `python -m json.tool cloud_snapshots\aps_snapshot_refresh_status.json`
  - `python -m json.tool cloud_snapshots\aps_snapshot_refresh_state.json`
  - `git diff --check`

---
저장 위치: `C:\Users\유현아\Documents\GitHub\INTEROJO\today_work_log_2026-09-04.md`
