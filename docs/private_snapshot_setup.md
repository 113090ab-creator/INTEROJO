# GitHub APS 스냅샷 자동 갱신 설정

이 브랜치는 GitHub Actions가 APS API를 조회한 결과를 `cloud_snapshots` 폴더에 저장하고, 변경된 스냅샷을 GitHub 저장소에 커밋하도록 구성합니다.

이 방식은 저장소가 Public이면 스냅샷 데이터도 공개될 수 있습니다. 운영상 스냅샷 공개가 허용될 때만 사용합니다.

## 현재 적용 기준

- 대상 저장소: `113090ab-creator/INTEROJO`
- GitHub Actions 실행: 매일 한국시간 08:20~09:20, 16:30~17:30에 여러 차례 예약하며 실패 시 실행 내부에서 5분 간격으로 최대 13회 시도
- `Keep Streamlit Alive`도 8시간 간격 예약으로 빠진 목표 회차의 백업 갱신을 시도함
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

### 운영 화면 갱신 지연 방지

- 클라우드 앱은 실행 시 및 사용 중 5분 간격으로 GitHub의 게시 포인터를 별도로 확인합니다. Streamlit의 코드 배포/저장소 동기화에 지연이 생겨도 게시된 데이터는 독립적으로 받아옵니다.
- 전체 manifest와 데이터 파일의 SHA-256 체크섬을 검증한 뒤 읽기 대상을 전환합니다. 실패·일부 파일 누락 시 기존 정상 데이터를 유지하며, 과거 회차로 되돌아가지 않습니다.
- 내려받은 파일은 `.local_cache/published_snapshots`에 저장합니다. Git 관리 파일을 수정하지 않으므로 앱 내 데이터 동기화가 Git pull 충돌을 만들지 않습니다. 오래된 캐시는 최근 정상본과 배포본을 보호하며 정리하고, 메모리 캐시 수도 제한합니다.
- 새 게시본을 저장한 `Refresh APS snapshots`와 백업 작업은 `Verify live APS snapshot`을 이어서 실행합니다. 이 검사는 브라우저로 실제 운영 화면의 Set ID와 상태를 읽어 게시본과 대조합니다. HTTP 접속 성공만으로 통과하지 않습니다.
- 별도로 `Verify live APS snapshot`을 매시 13분·43분에 예약합니다. 오전 09:55/오후 17:55까지의 갱신 허용시간을 고려하고, 허용시간 이후 게시 누락이나 운영 화면 반영 실패를 작업 실패로 남깁니다. GitHub 예약 실행 자체는 지연될 수 있습니다.
- 검사 결과는 Actions 실행 요약과 `live-snapshot-check` artifact(JSON, 14일 보관)에 남습니다. 작업 실패 알림은 해당 저장소의 GitHub Actions 알림 설정을 따릅니다.
- API HTTP 429 재시도는 `Retry-After`를 우선 적용하며, 안내가 없으면 10초·20초 대기를 사용합니다.

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
