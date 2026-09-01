# GitHub APS 스냅샷 자동 갱신 설정

이 브랜치는 GitHub Actions가 APS API를 조회한 결과를 `cloud_snapshots` 폴더에 저장하고, 변경된 스냅샷을 GitHub 저장소에 커밋하도록 구성합니다.

이 방식은 저장소가 Public이면 스냅샷 데이터도 공개될 수 있습니다. 운영상 스냅샷 공개가 허용될 때만 사용합니다.

## 현재 적용 기준

- 대상 저장소: `113090ab-creator/INTEROJO`
- GitHub Actions 실행: 평일 한국시간 16:30부터 17:00까지 5분 간격
- UTC cron: `30,35,40,45,50,55 7 * * 1-5`, `0 8 * * 1-5`
- 이미 같은 APS 기준시각의 스냅샷이 있으면 다시 쓰지 않음
- API 실패, 0건, 필수 컬럼 누락, 저장 실패 시 기존 정상 스냅샷을 덮어쓰지 않음

## GitHub Actions Secrets

GitHub 저장소의 `Settings > Secrets and variables > Actions > Repository secrets`에 아래 이름으로 등록합니다.

- `PLAN_API_KEY`
- `PLAN_API_BASE_URL`

실제 값은 코드, PR 설명, 로그에 적지 않습니다.

`PLAN_API_BASE_URL`은 기본값 `https://plan.interojo.net`을 그대로 쓸 경우 생략할 수 있습니다.

## Streamlit Cloud 설정

Streamlit 앱은 GitHub 저장소에 커밋된 `cloud_snapshots` 파일을 읽습니다. 별도 S3/R2 설정은 필요 없습니다.

Streamlit 화면에서 APS API를 직접 호출하지 않고 저장된 스냅샷만 빠르게 보여주려면 `PLAN_API_KEY`는 Streamlit Cloud에 등록하지 않아도 됩니다. APS 조회와 갱신은 GitHub Actions에서 수행합니다.

## 최초 실행 확인

1. PR을 생성하고 코드 변경을 검토합니다.
2. GitHub Actions Secrets를 등록합니다.
3. PR merge 후 GitHub `Actions > Refresh APS snapshots > Run workflow`를 누릅니다.
4. 실행 로그에서 `snapshot refresh completed`와 각 사이트별 행 수를 확인합니다.
5. workflow가 `Refresh APS snapshots` 커밋을 자동 생성했는지 확인합니다.
6. Streamlit 앱에서 사이드바의 스냅샷 저장소와 기준시각을 확인합니다.
7. 실패 시 화면에는 기존 정상 스냅샷과 함께 `마지막 갱신 실패` 경고가 표시됩니다.

## 주의할 점

- APS API가 사내망 또는 VPN 전용이면 GitHub-hosted runner에서는 실패합니다.
- 이 경우 GitHub Actions를 사내 PC/서버의 self-hosted runner에서 실행해야 합니다.
- 저장소가 Public이면 스냅샷 파일도 공개됩니다.
