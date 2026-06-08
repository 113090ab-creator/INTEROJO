# 오늘 작업 로그 (2026-06-08)

## 1) 작업 파일
- `app.py`
- `.gitignore`

## 2) Streamlit 대시보드 엑셀 로딩 최적화
- `app.py`에 `pickle` 기반 로컬 캐시 사용을 추가
- 앱 캐시 버전을 `20260608-bom-streaming-v5`로 갱신
- 대용량 엑셀을 `pandas.read_excel()`로 한 번에 읽는 구간을 줄이고, `openpyxl`의 `read_only=True` 스트리밍 방식으로 필요한 행/컬럼만 읽도록 변경
- 수요 파일 헤더 추출 로직을 첫 2행만 읽는 방식으로 정리
- 재고 파일 부분 로딩을 스트리밍 방식으로 변경
- 재작업 리스트 매칭 정보도 전체 시트 파싱 대신 필요한 컬럼만 순회하도록 변경

## 3) 대시보드 디스크 캐시 추가
- `.dashboard_cache` 폴더에 pickle 캐시를 저장하는 공통 함수 추가
  - `build_dashboard_cache_path()`
  - `read_pickle_cache()`
  - `write_pickle_cache()`
- 원본 파일명, 파일 크기, 수정시각을 캐시 키에 포함해 원본 엑셀이 바뀌면 캐시가 자동으로 달라지도록 구성
- 재고 부분 데이터와 수요 부분 데이터에 캐시 적용
- 캐시 읽기/쓰기 실패 시 앱 실행이 중단되지 않도록 예외를 무시하고 원본 로딩으로 진행

## 4) BOM 매핑 로딩 개선
- BOM 시트 매핑 계산을 `load_bom_maps_streaming()` 함수로 분리
- BOM 전체를 DataFrame으로 만든 뒤 정렬/중복 제거하던 흐름을 openpyxl 행 순회 방식으로 변경
- `SALES_ITEM_CD`, `TO_ITEM_ID`, `FROM_ITEM_ID`, `SEQ` 컬럼만 사용
- `SEQ` 기준 우선순위 계산을 위해 `parse_sequence_priority()` 추가
- 행 값 안전 추출을 위해 `get_row_value()` 추가
- R/Q 기본 매핑과 정확 매핑을 스트리밍 중 누적해 메모리 사용량을 줄이는 구조로 변경
- BOM 매핑 결과도 `.dashboard_cache`에 pickle로 저장해 재실행 시 반복 계산을 줄임

## 5) Git 제외 설정
- `.gitignore`에 `.dashboard_cache/`를 추가
- 대시보드 실행 중 생성되는 로컬 캐시 파일이 Git 변경사항으로 잡히지 않도록 처리

## 6) 운영 표준서 반영
- `운영_표준서.md` 문서 버전을 `v1.4`로 갱신
- 대용량 엑셀 스트리밍 로딩 기준 추가
  - `openpyxl.load_workbook(read_only=True, data_only=True)` 우선 검토
- `.dashboard_cache` 기반 로컬 디스크 캐시 허용 기준 추가
- 캐시 무효화 기준을 원본 파일명/파일 크기/수정시각/처리 버전 기반으로 명시
- 캐시 읽기/쓰기 실패 시 원본 로딩으로 계속 진행하는 운영 원칙 추가
- `.dashboard_cache/` Git 제외 및 커밋 전 확인 항목 추가

## 7) 검증/상태
- 현재 Git 변경 파일:
  - `.gitignore`
  - `app.py`
  - `운영_표준서.md`
- 변경 규모:
  - `app.py`: 엑셀 로딩/캐시/BOM 매핑 최적화
  - `.gitignore`: `.dashboard_cache/` 제외
  - `운영_표준서.md`: v1.4 표준 반영
- 이번 로그 파일:
  - `today_work_log_2026-06-08.md`

## 8) 생산현황 UI 복구 및 배포 동기화
- 생산현황 화면 UI 정리 시도 후, 사용자 요청에 따라 기존 화면 구조로 복구
- 새 UI 관련 코드 제거 확인
  - `전체 생산현황`
  - `상세 필터 열기`
  - `render_shortage_dashboard_v2`
- 기존 공정별 현황 구조 유지 확인
  - `생산 현황`
  - `사출 현황`
  - `분리 현황`
  - `공용 품목 현황`
- `APP_CACHE_VERSION`을 로그 기준 시점인 `20260608-bom-streaming-v5`로 복구
- 불필요한 재배포용 `requirements.txt` 주석 제거
- `python -m py_compile app.py`로 문법 검사 통과 확인
- GitHub `main` 최신 커밋 확인
  - `461fc37 Restore app to work log baseline`
- Streamlit Cloud 앱이 과거 배포본을 계속 표시하던 문제 확인
- GitHub에는 정상 반영되어 있었고, Streamlit Cloud 앱 재부팅 후 화면 복구 확인

## 9) 현재 기준 상태
- GitHub 원격 저장소: `https://github.com/113090ab-creator/INTEROJO.git`
- 기준 브랜치: `main`
- 현재 기준 커밋: `461fc37 Restore app to work log baseline`
- 배포 앱: `https://interojo-2.streamlit.app/`
- 운영 기준: 이 커밋을 복구 기준점으로 유지

---
저장 위치: `C:\Users\유현아\Documents\GitHub\INTEROJO\today_work_log_2026-06-08.md`
