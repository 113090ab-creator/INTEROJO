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

## 8) 추가 안정화 작업
- Streamlit Cloud에서 앱이 배포 시점의 `cloud_snapshots`만 읽는 문제가 있어, 운영 환경에서는 GitHub `main`의 최신 `cloud_snapshots`를 직접 읽도록 보강함
- 로컬 PC가 꺼져 있어도 GitHub Actions가 스냅샷을 만들고 `main`에 push하면, 사용자는 Streamlit 브라우저 새로고침만으로 최신 스냅샷을 읽을 수 있는 구조로 변경함
- GitHub 스냅샷 읽기는 read-only로만 동작하게 했고, 스냅샷 생성/쓰기 경로는 기존 GitHub Actions commit/push 방식을 유지함
- GitHub Actions 전용 갱신과 keepalive 백업 갱신을 토요일/일요일에도 실행하도록 변경함
- 예약 갱신 로그에 변경 파일, commit SHA, push 성공 여부, 스냅샷 행 수와 파일 크기 리포트를 출력하도록 추가함
- WIP 스냅샷이 이미 최신이면 WIP API 재조회 없이 기존 정상 WIP 스냅샷을 재사용하도록 변경함
- 각 관별 생산부족 스냅샷이 이미 최신이면 해당 관 재계산을 건너뛰도록 변경함
- 신규 스냅샷 행 수가 기존 정상 스냅샷 대비 50% 미만으로 급감하면 실패 처리하고 기존 정상 스냅샷을 유지하도록 보호장치를 추가함

## 9) 성능 측정 메모
- CSV gzip 로딩 측정:
  - WIP `30,130`건: `0.0303초`
  - C관 생산부족 `5,093`건: `0.0255초`
  - A관 생산부족 `1,306`건: `0.0086초`
  - S관 생산부족 `7,112`건: `0.0400초`
  - 전체 생산부족 `13,511`건: `0.0690초`
  - 전체 품목 현황 `110,882`건: `0.3995초`
- Parquet은 읽기는 빠르지만 현재 파일 기준 CSV gzip보다 파일 크기가 커졌음
- GitHub에서 파일을 내려받는 운영 구조에서는 즉시 Parquet 전환보다 불필요한 원격 확인 요청 제거와 반복 계산 건너뛰기가 우선이라고 판단함
- `DEBUG_PERFORMANCE=1`일 때만 `[PERF]` 로그가 출력되도록 추가함

## 10) 추가 검증
- 문법 검증:
  - `python -m py_compile app.py snapshot_storage.py scripts\refresh_snapshot.py scripts\refresh_aps_shortage_snapshots.py scripts\report_snapshot_status.py`
- 기존 스냅샷 검증:
  - `python scripts\refresh_snapshot.py --validate-existing --sites C관,A관,S관,전체`
- GitHub 원격 스냅샷 읽기 검증:
  - `SNAPSHOT_STORAGE_BACKEND=github`
  - `SNAPSHOT_GITHUB_REPOSITORY=113090ab-creator/INTEROJO`
  - `SNAPSHOT_GITHUB_BRANCH=main`
  - `SNAPSHOT_GITHUB_PREFIX=cloud_snapshots`
  - `snapshot_meta.csv` 원격 읽기 성공

## 11) Streamlit 최초 접속 성능 개선
- 목표:
  - Cloud Snapshot 모드에서 최초 접속 시 로컬 Excel 파일을 열지 않도록 개선
  - 기존 화면, 필터, 계산식, 데이터 결과는 유지
  - GitHub Actions, API 갱신 로직, snapshot 생성 로직은 수정하지 않음
- 확인된 주요 병목:
  - `select_data_source()`에서 `get_data_updated_at()` fallback이 즉시 실행되어 로컬 Excel 기준일자를 읽음
  - `render_sidebar_reference_dates()`가 생산 부족 현황 + Cloud Snapshot 화면에서도 로컬 WIP/수요 파일 기준일자를 읽음
  - `snapshot_meta.csv`, `aps_snapshot_refresh_status.json`, `aps_snapshot_refresh_state.json`을 여러 경로에서 반복 조회함
  - WIP 최신 여부 판단 시 전체 WIP snapshot을 먼저 읽는 경로가 있었음
  - `openpyxl`이 앱 import 시점에 전역 로드됨
- 조치:
  - `get_cloud_snapshot_meta_value()` fallback을 lazy 구조로 변경
  - `snapshot_context`를 추가해 Cloud metadata/status/state를 한 번 읽고 재사용
  - Cloud Snapshot 생산 부족 화면에서는 사이드바에 간단한 현장용 기준 정보만 표시:
    - 수요 기준시각
    - WIP 기준시각
    - 상태
  - 상세 기준 정보와 저장소 정보는 `DEBUG_PERFORMANCE=1`일 때만 표시
  - WIP 최신 여부는 metadata/status 기준으로 먼저 판단하고, 전체 WIP snapshot은 실제 계산에 필요할 때만 읽도록 변경
  - `openpyxl` 전역 import를 제거하고 Excel fallback/다운로드 기능 사용 시점에만 lazy import하도록 변경
- 수정 전/후 측정:

```text
                     변경 전     변경 후
Warm First Load       74.33초     1.97초
Snapshot load          2.39초     0.27초
Sidebar               40.47초     0.54초
Page calculation      30.06초     0.10초
Render                30.80초     0.69초
TOTAL                 74.33초     1.97초
```

- 추가 측정:
  - 첫 실행 GitHub HTTP 요청 횟수: `17회 -> 7회`
  - 로컬 Excel open 횟수: `5회 -> 0회`
  - 49MB `1Day_2026Y,해외수주,포장,출고관리_0829_정재훈.xlsx` 최초 접속 open 여부: `YES -> NO`
- 최종 검증:
  - `python -m py_compile app.py snapshot_storage.py`
  - `git diff --check`
  - 로컬 Streamlit 실행:
    - `STREAMLIT_CLOUD=1`
    - `DEBUG_PERFORMANCE=1`
    - `streamlit run app.py --server.port 8502 --server.headless true --server.fileWatcherType none --browser.gatherUsageStats false`
  - 실제 Streamlit 로그:
    - `snapshot_context: 0.120 sec`
    - `sidebar_setup: 0.542 sec`
    - `data_load_total: 0.267 sec`
    - `shortage_apply_filters: 0.099 sec`
    - `render_total: 0.689 sec`
    - `main_total: 1.968 sec`
- 반영 커밋:
  - `f10da9a86715805938c8fe52ed51f470ad232b84 Improve Streamlit cloud snapshot first load`
