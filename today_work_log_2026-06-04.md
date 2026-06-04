# 오늘 작업 로그 (2026-06-04)

## 1) Streamlit 대시보드 성능 최적화
- `app.py` 전체 로딩/전처리/렌더링 흐름 점검
- 엑셀 원본 로딩과 전처리 함수를 분리
  - `load_raw_data()`
  - `preprocess_data()`
  - `filter_data()`
- 수요/재고 엑셀 로딩 시 화면 계산에 필요한 컬럼만 `usecols`로 읽도록 변경
- `pd.read_excel()` 반복 실행을 줄이기 위해 원본 로딩/전처리 함수에 `@st.cache_data` 적용
- 업데이트 시각 조회도 파일 크기/mtime 기반 캐시로 변경
- 필터 옵션 집계와 실제 필터링을 분리
  - `build_filter_option_maps()`
  - `filter_data()`
- 반복 계산이 큰 요약 함수에 캐시 적용
  - R코드 요약
  - Q코드 요약
  - RQ 그룹 요약
  - 이니셜별 사출 요약
  - 전체 수요 요약
  - 리드지 요약

## 2) 메뉴별 지연 로딩 적용
- 앱 시작 시 모든 리드지 데이터를 즉시 로딩하지 않도록 변경
- `리드지 현황`, `생산코드별 리드지` 메뉴에 진입할 때만 리드지 기준정보를 로딩
- `리드지 발주현황.xlsx`는 `리드지 현황` 메뉴에서만 로딩
- 초기 `생산 부족 현황` 진입 시 불필요한 리드지 파일 로딩을 제거

## 3) 대용량 표 렌더링 안정화
- 큰 DataFrame을 화면에 그대로 모두 출력하지 않도록 표시 행 제한 적용
- `DATAFRAME_DISPLAY_ROW_LIMIT = 1000`
- 화면에는 상위 1,000건만 표시
- 다운로드 버튼은 필터 결과 전체 데이터를 유지
- 표시용 숫자 포맷 변환과 엑셀 다운로드 바이트 생성도 캐시 적용

## 4) 리드지 기준파일 추가 최적화
- `제품명 기준 정보.xlsx`의 첫 번째 시트는 필요한 컬럼만 읽도록 변경
- `리드지정보` 시트는 전체 컬럼 대신 3개 컬럼만 로딩
  - `생산`
  - `B1코드`
  - `B1코드명`
- `리드지재고` 시트는 전체 컬럼 대신 3개 컬럼만 로딩
  - `품목코드`
  - `창고`
  - `재고`
- 리드지 매핑 계산 분리 및 캐시 적용
  - `build_leadji_code_mapping()`
- 리드지 재고 pivot/groupby 계산 분리 및 캐시 적용
  - `build_leadji_stock_pivot()`

## 5) 데이터 소스 UI 깨짐 수정
- 사이드바 `데이터 소스` 영역에서 Streamlit expander 화살표 아이콘이 `_arrow_right` 텍스트로 보이는 문제 확인
- 원인: Material icon 폰트 fallback 문제
- 조치:
  - 데이터 소스 업로드 UI를 expander 대신 toggle 기반 UI로 변경
  - expander 아이콘 fallback 텍스트가 보이지 않도록 CSS 방어 코드 추가

## 6) Streamlit Cloud 자동 유지 설정
- GitHub Actions 워크플로우 추가
  - `.github/workflows/keepalive.yml`
- 워크플로우명: `Keep Streamlit Alive`
- 실행 방식:
  - 수동 실행: `workflow_dispatch`
  - 자동 실행: 매일 3회, 8시간 간격
- GitHub repository secret 생성 확인
  - `STREAMLIT_URL`
  - 값: `https://interojo-2.streamlit.app/`
- Actions 수동 실행 결과 확인
  - `Keep Streamlit Alive #1`
  - `ping`
  - `Success`

## 7) 저장소 표준 설정 반영
- Excel 임시 잠금 파일이 Git 변경사항으로 잡히지 않도록 `.gitignore`에 제외 기준 추가
- 추가 패턴:
  - `~$*.xls`
  - `~$*.xlsx`
  - `~$*.xlsm`
- 운영 표준서에도 임시 잠금 파일 제외 기준과 커밋 전 확인 항목을 추가

## 8) 사용자 안내 및 운영 확인
- GitHub Desktop에서 workflow 파일 커밋/푸시 상태 확인 지원
- GitHub secret 이름 입력 규칙 안내
  - `Name = STREAMLIT_URL`
  - `Secret = https://interojo-2.streamlit.app/`
- GitHub Actions 수동 실행 성공 화면 확인

## 9) 검증
- 문법 검증
  - `python -m py_compile app.py` 통과
- 실제 데이터 로딩 검증
  - 메인 데이터 로딩 정상
  - 리드지정보 로딩 정상
  - 리드지재고 로딩 정상
  - 리드지 요약 계산 정상
- Streamlit 로컬 응답 확인
  - `http://127.0.0.1:8501`
  - HTTP 200 확인
- GitHub Actions keepalive 실행 확인
  - Success 확인

## 10) 운영 메모
- 현재 최적화는 별도 parquet/pkl 파일을 생성하지 않는 안전한 방식
- 원본 엑셀 파일은 그대로 유지
- 캐시 무효화 기준은 파일명/크기/mtime 기반
- 첫 접속은 Streamlit Cloud cold start와 엑셀 로딩 때문에 여전히 느릴 수 있음
- 같은 세션의 필터 변경/메뉴 전환은 이전보다 빨라지는 구조

---
저장 위치: `C:\Users\유현아\Documents\GitHub\INTEROJO\today_work_log_2026-06-04.md`
