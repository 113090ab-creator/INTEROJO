# 비공개 APS 스냅샷 자동 갱신 설정

이 브랜치는 Public GitHub 저장소에 생산 스냅샷을 커밋하지 않고, GitHub Actions가 APS API를 조회한 결과를 비공개 S3 호환 저장소에 저장하도록 구성합니다.

## 현재 적용 기준

- 대상 저장소: `113090ab-creator/INTEROJO`
- GitHub Actions 실행: 평일 한국시간 16:00부터 16:30까지 5분 간격
- UTC cron: `0,5,10,15,20,25,30 7 * * 1-5`
- 이미 같은 APS 기준시각의 스냅샷이 있으면 다시 쓰지 않음
- API 실패, 0건, 필수 컬럼 누락, 저장 실패 시 기존 정상 스냅샷을 덮어쓰지 않음

## 권장 저장소

Public 저장소를 유지하려면 Cloudflare R2 또는 AWS S3 같은 비공개 S3 호환 저장소를 사용합니다.

Cloudflare R2를 사용할 경우 일반적인 값은 다음과 같습니다.

- `SNAPSHOT_STORAGE_BACKEND`: `s3`
- `SNAPSHOT_S3_REGION`: `auto`
- `SNAPSHOT_S3_ENDPOINT_URL`: `https://<cloudflare-account-id>.r2.cloudflarestorage.com`
- `SNAPSHOT_S3_PREFIX`: 예: `interojo/snapshots`

## GitHub Actions Secrets

GitHub 저장소의 `Settings > Secrets and variables > Actions > Repository secrets`에 아래 이름으로 등록합니다.

- `PLAN_API_KEY`
- `PLAN_API_BASE_URL`
- `SNAPSHOT_S3_BUCKET`
- `SNAPSHOT_S3_PREFIX`
- `SNAPSHOT_S3_REGION`
- `SNAPSHOT_S3_ENDPOINT_URL`
- `SNAPSHOT_S3_ACCESS_KEY_ID`
- `SNAPSHOT_S3_SECRET_ACCESS_KEY`

실제 값은 코드, PR 설명, 로그에 적지 않습니다.

## Streamlit Cloud Secrets

Streamlit 앱이 저장된 스냅샷을 읽으려면 Streamlit Cloud의 app settings/secrets에도 아래 값을 등록합니다.

```toml
SNAPSHOT_STORAGE_BACKEND = "s3"
SNAPSHOT_S3_BUCKET = "<bucket-name>"
SNAPSHOT_S3_PREFIX = "interojo/snapshots"
SNAPSHOT_S3_REGION = "auto"
SNAPSHOT_S3_ENDPOINT_URL = "https://<cloudflare-account-id>.r2.cloudflarestorage.com"
SNAPSHOT_S3_ACCESS_KEY_ID = "<access-key-id>"
SNAPSHOT_S3_SECRET_ACCESS_KEY = "<secret-access-key>"
```

Streamlit 화면에서 APS API를 직접 호출하지 않고 저장된 스냅샷만 빠르게 보여주려면 `PLAN_API_KEY`는 Streamlit Cloud에 등록하지 않아도 됩니다. APS 조회와 갱신은 GitHub Actions에서 수행합니다.

## 최초 실행 확인

1. PR을 생성하고 코드 변경을 검토합니다.
2. GitHub Actions Secrets를 등록합니다.
3. Streamlit Cloud Secrets를 등록합니다.
4. PR merge 후 GitHub `Actions > Refresh APS snapshots > Run workflow`를 누릅니다.
5. 실행 로그에서 `snapshot refresh completed`와 각 사이트별 행 수를 확인합니다.
6. Streamlit 앱에서 사이드바의 스냅샷 저장소와 기준시각을 확인합니다.
7. 실패 시 화면에는 기존 정상 스냅샷과 함께 `마지막 갱신 실패` 경고가 표시됩니다.

## 주의할 점

- APS API가 사내망 또는 VPN 전용이면 GitHub-hosted runner에서는 실패합니다.
- 이 경우 GitHub Actions를 사내 PC/서버의 self-hosted runner에서 실행해야 합니다.
- 이번 PR은 앞으로 생산 데이터 파일이 새로 커밋되는 것을 막지만, 이미 Public 저장소의 과거 커밋에 들어간 데이터는 별도 history purge 또는 저장소 Private 전환이 필요합니다.
