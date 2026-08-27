# 오늘 작업 로그 (2026-08-27)

## 1) 작업 범위
- 기준 브랜치: `main`
- 대상 앱: `https://interojo-2.streamlit.app/`
- 최종 기준 커밋: `7726428 Schedule APS shortage snapshot refresh`
- 작업 목적:
  - 생산 부족 현황 화면의 조회 속도 개선
  - APS 수요/WIP API 기준의 데이터 정확성 강화
  - 오래된 스냅샷 또는 API 실패 데이터가 현장에 표시되지 않도록 오류 안내 강화
  - APS API 하루 2회 갱신 시점에 맞춘 자동 스냅샷 갱신 구조 구성

## 2) 확인한 주요 문제
- 왼쪽 메뉴와 사이드바 필터를 열 때 항목 수가 많은 필터 때문에 화면 반응이 느림
- 생산 부족 현황에서 APS API 또는 WIP API가 실패해도 기존 파일/스냅샷이 섞여 표시될 수 있는 위험 확인
- WIP API는 창고별 데이터량이 많고 응답 시간이 길어 화면에서 직접 조회하면 체감 속도가 크게 떨어짐
- PLAN API 기준시각과 화면 스냅샷 기준시각이 다르면 현장에 오래된 생산 부족 데이터가 표시될 수 있음
- 기존 GitHub Actions 스냅샷 workflow가 WIP 관련 메타데이터를 지울 수 있는 구조 확인

## 3) 사이드바와 화면 속도 개선
- 메뉴/사이드바 진입 시 불필요하게 live API나 무거운 계산을 반복하지 않도록 흐름 점검
- 기본 생산 부족 현황은 저장된 cloud snapshot을 우선 표시하도록 유지
- 사이드바 반영 기준일자 표시에서 불필요한 live API 호출을 줄임
- 항목이 많은 사이드바 필터 중 `시트 분류`, `분류별 요약`은 요청대로 `pills` 버튼 나열 방식으로 원복
- 관련 커밋:
  - `5778f11 Improve Streamlit sidebar performance`
  - `3068fcd Avoid live API calls in sidebar reference dates`
  - `9d973ed Improve sidebar filter responsiveness`
  - `1e0739c Revert sidebar category filters to pills`

## 4) APS/WIP 오류 표시 정책 변경
- APS API가 정상 조회되지 않거나 필요한 API 일부가 실패하면 큰 빨간 오류 문구를 표시하도록 정리
- WIP API 실패 시 WIP 엑셀 파일로 대체하지 않도록 변경
- 화면에서는 잘못된 데이터가 보이는 것보다 오류로 막는 쪽을 우선하도록 처리
- 전일 스냅샷이거나 APS 최신 기준시각보다 오래된 스냅샷이면 생산 부족 현황 표시를 중지하고 오류 안내
- 관련 커밋:
  - `cd44076 Require APS WIP API for inventory data`

## 5) WIP 정리 스냅샷 구조 추가
- WIP API 원본을 화면에서 직접 매번 조회하지 않고, 먼저 별도 위치에 받은 뒤 정리본을 사용하도록 변경
- 원본 WIP 저장 위치:
  - `outputs/aps_wip_raw/<기준시각>/aps_wip_raw.csv.gz`
  - `outputs/aps_wip_raw/<기준시각>/metadata.json`
- 앱이 실제로 읽는 배포용 정리 스냅샷:
  - `cloud_snapshots/wip_inventory_snapshot.csv.gz`
- WIP 정리 스냅샷 메타데이터:
  - `wip_inventory_updated_at`
  - `wip_inventory_source_label`
  - `wip_inventory_refreshed_at`
- 현재 생성된 WIP 정리 스냅샷:
  - 기준시각: `2026-08-27 16:11:43`
  - 행 수: `29,727`
  - 대상 창고: `사출창고`, `분리창고`, `검사접착`, `누수규격검사`
- 관련 커밋:
  - `807cf31 Use prebuilt APS WIP inventory snapshot`

## 6) 생산 부족 스냅샷 재생성
- WIP 정리 스냅샷을 기준으로 생산 부족 스냅샷을 다시 생성
- 현재 반영된 APS PLAN 기준시각:
  - `2026-08-27 15:57:23`
- 현재 반영된 APS WIP 기준시각:
  - `2026-08-27 16:11:43`
- 갱신된 범위:
  - `C관`
  - `A관`
  - `S관`
  - `전체`
- 최신 상태에서 재실행하면 `WIP/C관/A관/S관/전체` 모두 `skip-current`로 빠르게 종료되는 것을 확인

## 7) 자동 스냅샷 예약 구조 추가
- APS API가 하루 2회 갱신된다는 운영 기준을 반영
- 스냅샷 확인 시간:
  - 오전: `07:25 ~ 08:00`, 5분 간격
  - 오후: `15:55 ~ 16:30`, 5분 간격
- 처리 순서:
  - APS PLAN 기준시각 확인
  - APS WIP 정리 스냅샷 먼저 확인/생성
  - C관/A관/S관/전체 생산 부족 스냅샷 생성
  - 성공한 슬롯은 상태 파일에 기록하고 이후 실행은 `skip-current`
- 새 GitHub Actions workflow:
  - `.github/workflows/refresh-aps-shortage-snapshots.yml`
- 상태 로그 파일:
  - `cloud_snapshots/aps_snapshot_refresh_status.json`
  - `cloud_snapshots/aps_snapshot_refresh_state.json`
- 관련 커밋:
  - `7726428 Schedule APS shortage snapshot refresh`

## 8) GitHub Actions 관련 주의사항
- 자동 예약 실행에는 GitHub repository secret `PLAN_API_KEY`가 필요
- Streamlit Cloud의 secret과 GitHub Actions secret은 별도임
- GitHub 저장소의 Settings > Secrets and variables > Actions에 `PLAN_API_KEY`가 없으면 예약 작업은 API 키 없음으로 실패함
- 기존 `refresh-cloud-snapshots.yml`은 엑셀 파일 변경 시에만 실행되도록 제한
- 기존 cloud snapshot refresh가 `snapshot_meta.csv`의 WIP 메타데이터를 지우지 않도록 보존형으로 수정

## 9) 검증 결과
- 문법 검증:
  - `python -m py_compile app.py scripts\refresh_aps_shortage_snapshots.py`
  - `python -m py_compile scripts\refresh_cloud_snapshots.py`
- 공백 검사:
  - `git diff --check`
  - `git diff --cached --check`
- WIP 정리 스냅샷 수동 생성 확인:
  - `python scripts\refresh_aps_shortage_snapshots.py --wip-only --only-if-stale`
- 전체 생산 부족 스냅샷 생성 확인:
  - `python scripts\refresh_aps_shortage_snapshots.py --only-if-stale`
- 예약 모드 실행 확인:
  - `python scripts\refresh_aps_shortage_snapshots.py --scheduled --only-if-stale`
- 현재 로컬/원격 상태:
  - `main`과 `origin/main` 동기화 완료

## 10) 운영 메모
- 화면은 가능한 빠르게 `cloud_snapshots`를 읽도록 구성
- 최신 WIP 정리 스냅샷이 없거나 오래되면 live WIP API로 대체하지 않고 오류로 막음
- 생산 부족 현황은 현장 사용 데이터이므로, 오래된 스냅샷을 표시하는 것보다 오류를 띄우는 정책을 우선 적용
- 원본 WIP raw 파일은 `outputs/` 하위에 저장되며 `.gitignore` 대상이라 GitHub에는 올라가지 않음
- 배포 앱에 필요한 정리본만 `cloud_snapshots/wip_inventory_snapshot.csv.gz`로 관리
- Streamlit 배포 화면 반영은 GitHub push 이후 Streamlit Cloud가 재배포하는 시간에 따라 약간 지연될 수 있음

---
저장 위치: `C:\Users\유현아\Documents\GitHub\INTEROJO\today_work_log_2026-08-27.md`
