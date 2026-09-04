# 오늘 작업 로그 (2026-09-03)

## 1) 작업 범위
- 기준 브랜치: `main`
- 대상 저장소: `113090ab-creator/INTEROJO`
- 대상 앱: `https://interojo-2.streamlit.app/`
- 최종 기준 커밋: `5034604 Retry APS refresh inside scheduled run`
- 작업 목적:
  - GitHub Actions APS 스냅샷 예약 실행이 기대대로 동작하지 않는 원인 확인
  - 아침/오후 APS PLAN, APS WIP 기준시각에 맞춘 자동 스냅샷 갱신 구조 보완
  - 스냅샷 상태 파일 충돌 해결
  - 현장 화면에 오래된 오류 상태가 남지 않도록 최신 정상 스냅샷 상태 정리

## 2) 확인한 주요 현상
- Streamlit 화면에 아래 오류가 표시됨:
  - 마지막 갱신 실패: `2026-09-02 21:38:03`
  - 원인: APS WIP 정리 스냅샷이 최신 기준보다 오래됨
  - 스냅샷 기준: `2026-09-01 16:25:27`
  - APS WIP 최신 기준: `2026-09-02 16:14:43`
- GitHub Actions `Refresh APS snapshots`의 최근 예약 실행 기록은 `2026-09-02` 실패 2건까지만 확인됨
- `2026-09-03`의 `Refresh APS snapshots` 예약 실행은 GitHub Actions 기록에 생성되지 않았음
- 반면 `Keep Streamlit Alive` workflow는 `2026-09-03 13:02:20 KST`에 정상 실행되어, 저장소 전체 Actions가 꺼진 상태는 아니었음

## 3) 스냅샷 상태 확인
- 원격 `cloud_snapshots/aps_snapshot_refresh_status.json` 기준 최신 상태:
  - `checked_at`: `2026-09-03 16:26:22`
  - `slot_key`: `2026-09-03 16:00`
  - APS PLAN 기준시각: `2026-09-03 16:03:26`
  - APS WIP 기준시각: `2026-09-03 16:19:08`
  - 상태: `completed`
- 갱신 행 수:
  - WIP: `30,157`
  - C관: `4,563`
  - A관: `1,463`
  - S관: `7,421`
  - 전체: `13,447`
- 단, 해당 스냅샷 커밋 `c5fcb66`은 `github-actions[bot]`이 아니라 사용자 계정 커밋으로 확인됨
- 따라서 오늘 오후 스냅샷 데이터는 최신으로 올라왔지만, GitHub Actions 예약 실행이 성공해서 올라온 상태로 보기는 어려움

## 4) 충돌 해결
- GitHub Desktop에서 `cloud_snapshots/aps_snapshot_refresh_status.json` 충돌 발생
- 충돌 원인:
  - 로컬에는 `2026-09-03 08:20:09` 성공 상태가 있었음
  - 원격에는 `2026-09-02 21:38:03` 실패 상태가 있었음
  - 자동 스냅샷 상태 파일이 양쪽에서 동시에 바뀌어 병합 충돌 발생
- 해결 기준:
  - 더 최신이고 실제 스냅샷 파일과 맞는 `2026-09-03 08:20:09` 성공 상태를 남김
- 처리 결과:
  - 충돌 제거
  - merge 커밋 생성
  - 원격 `main`에 푸시 완료
- 관련 커밋:
  - `c299e0d Merge latest snapshot refresh changes`

## 5) 아침 예약 복구
- 기존 `refresh_snapshot.yml`에는 오후 `16:30~17:00 KST` 예약만 남아 있었음
- 그래서 오늘 아침 GitHub Actions 자동 실행은 구조상 발생하지 않았음
- 아침 예약을 다시 추가함
- 적용한 예약:
  - 아침: 평일 `08:20~08:50 KST`
  - 오후: 평일 `16:30~17:00 KST`
- 관련 커밋:
  - `8ff138b Add morning APS snapshot refresh window`

## 6) 예약 실행 구조 개선
- 기존 방식:
  - 5분마다 별도 GitHub Actions run을 여러 개 생성
  - GitHub schedule 지연/누락 시 실패 상태가 흩어짐
  - 스냅샷 파일을 여러 run이 건드릴 수 있어 충돌 가능성이 커짐
- 변경 방식:
  - 아침 `08:20 KST`, 오후 `16:30 KST`에 각각 한 번만 실행
  - 실행 내부에서 5분 간격으로 최대 7번 재시도
  - 최대 약 30분 동안 APS PLAN/WIP 최신 기준시각 반영을 기다림
  - workflow 전체 제한 시간은 45분으로 설정
  - 실행 시작 시 `git pull --ff-only origin main`으로 최신 스냅샷을 먼저 반영
  - push는 `git push origin HEAD:main`으로 명시
- 현재 UTC cron:
  - 아침: `20 23 * * 0-4`
  - 오후: `30 7 * * 1-5`
- 관련 커밋:
  - `5034604 Retry APS refresh inside scheduled run`

## 7) 현재 판단
- 오늘 `2026-09-03` 오후 데이터 자체는 스냅샷에 반영되어 있음
- 다만 오늘 반영은 GitHub Actions 예약 성공이 아니라 사용자 계정 커밋으로 올라온 것으로 확인됨
- GitHub Actions schedule은 active 상태지만, `Refresh APS snapshots`의 오늘 예약 run은 생성되지 않았음
- GitHub Actions schedule은 GitHub 내부 부하나 예약 처리 지연으로 늦거나 누락될 수 있음
- 누락 가능성을 줄이기 위해 여러 개 예약 run을 만들기보다, 한 번 실행 안에서 재시도하는 구조로 변경함

## 8) 검증 결과
- 문법 검증:
  - `python -m py_compile app.py scripts\refresh_snapshot.py snapshot_storage.py`
- 공백 검사:
  - `git diff --check`
- 원격 상태 확인:
  - `main`과 `origin/main` 동기화 완료
  - `refresh_snapshot.yml`에 아침/오후 예약 반영 확인
  - `aps_snapshot_refresh_status.json` 최신 completed 상태 확인

## 9) 운영 메모
- 다음 자동 확인 시각:
  - 평일 아침 `08:20 KST`
  - 평일 오후 `16:30 KST`
- APS WIP가 아직 최신화되지 않았으면 같은 run 안에서 5분 뒤 재시도함
- 최대 7회까지 실패하면 마지막 실패 상태가 화면에 경고로 표시됨
- GitHub Actions 예약 자체가 생성되지 않는 경우에는 workflow 로그가 남지 않음
- 정확한 당일 반영을 즉시 확인해야 할 때는 GitHub `Actions > Refresh APS snapshots > Run workflow`를 수동 실행
- 수동으로 스냅샷을 커밋하면 Actions가 만든 스냅샷 커밋과 충돌할 수 있으므로, 가능하면 수동 실행 workflow를 사용하는 것이 안전함

## 10) 2026-09-04 추가 확인 및 보완
- `2026-09-04 09:01 KST` 기준 `Refresh APS snapshots`의 아침 예약 실행 기록이 GitHub Actions에 생성되지 않았음
- 같은 기간 `Keep Streamlit Alive` workflow는 최근까지 정상 schedule 실행 기록이 있어 저장소 전체 Actions 중지는 아닌 것으로 판단
- 로컬 예약/실행 로그에서는 `2026-09-04 08:17:27`에 아침 APS 스냅샷 생성이 성공했음
- 생성된 아침 스냅샷:
  - APS PLAN 기준시각: `2026-09-04 08:00:15`
  - APS WIP 기준시각: `2026-09-04 08:14:45`
  - WIP: `30,122`건
  - C관: `4,381`건
  - A관: `1,361`건
  - S관: `7,822`건
  - 전체: `13,564`건
- 해당 스냅샷을 `cloud_snapshots`에 커밋하고 원격 `main`에 푸시함
- `Refresh APS snapshots` 예약 누락에 대비해, 이미 작동 이력이 있는 `Keep Streamlit Alive` workflow에도 백업 APS 갱신 job을 추가함
- 백업 갱신은 평일 `09:00`, `17:00 KST`에만 실행되며, `01:00 KST` keepalive에서는 스냅샷 갱신을 건너뜀
- 백업 갱신도 실행 내부에서 최대 7회, 5분 간격으로 재시도함

---
저장 위치: `C:\Users\유현아\Documents\GitHub\INTEROJO\today_work_log_2026-09-03.md`
