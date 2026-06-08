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

---
저장 위치: `C:\Users\유현아\Documents\GitHub\INTEROJO\today_work_log_2026-06-08.md`
