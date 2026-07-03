# calendar_sync — 프로젝트 규칙

## 규칙 관리

- 프로젝트 규칙의 **추가·변경·삭제**는 항상 이 파일(`claude.md`)을 수정한다.
- 규칙 변경을 채팅, `.cursor/rules/`, 기타 파일에만 반영하지 않는다.
- 사용자가 다른 저장 위치를 명시적으로 요청한 경우에만 예외로 한다.

## 언어·커뮤니케이션

- 모든 답변, 문서, 커밋 메시지는 **한국어**로 작성한다.
- 코드 주석도 한국어로 작성한다. (의도와 이유 중심)

## 기술 스택

- **언어**: Python 3.12+
- **주요 라이브러리**:
  - Google Calendar: `google-api-python-client`, `google-auth-oauthlib`
  - Microsoft Outlook: `pywin32` (로컬 Outlook COM — MS Graph/MSAL 미사용)
  - 환경 변수: `python-dotenv`, `tzdata`
- **실행 환경**: Windows PC + Outlook 데스크톱 앱, Windows Task Scheduler (30분 주기)
- 스택 변경이 필요하면 기획서와 이 파일을 함께 갱신한다.

## 개발 워크플로우

- 모든 개발은 **기획서**를 기반으로 진행한다. 기획서 없이 코드 작업을 시작하지 않는다.
- 기획서에는 반드시 **체크리스트**를 포함한다. (요구사항·작업 항목·완료 기준 등을 체크 가능한 목록으로 작성)
- 구현 전 기획서를 사용자와 확인하고, 승인 후 작업한다.

### 기획서 위치·파일명

- 기획서는 `Docs/` 하위에 저장한다.
- 파일명 형식: `NN_주제.md` (예: `01_캘린더_동기화_기획.md`)
- 개발이 완료된 기획서는 `Docs/Finished/`로 옮긴다.

### 코드 위치

- 모든 프로그램 코드는 `Programs/` 하위에 작성한다.
- 테스트용 코드는 `Programs/test_code/` 하위에서만 작성한다. 다른 경로에 테스트 코드를 두지 않는다.

### 의존성 관리

- Python 패키지 목록은 `Programs/requirements.txt`에 관리한다.
- 새 패키지 추가 시 `requirements.txt`도 함께 갱신한다.

### 완료 정의 (Definition of Done)

개발 완료 시 아래를 모두 충족한다.

- [ ] 기획서 체크리스트 전부 완료
- [ ] `Programs/test_code/`에서 테스트 통과
- [ ] git 커밋 및 push
- [ ] 기획서를 `Docs/Finished/`로 이동

## 환경 변수·보안

- 비밀 값은 `.env`에만 저장한다. `.env`는 커밋하지 않는다.
- 필요한 키 목록은 `.env.example`에 값 없이 문서화한다.
- OAuth 토큰, `credentials.json`, `token_google.json` 등 인증 파일은 커밋하지 않는다.
- 동기화 상태·이벤트 매핑 파일(`Programs/sync/*.json`)은 커밋하지 않는다.
- `.gitignore`에 위 파일들이 포함되어 있는지 확인한다.

## 동기화 정책

- **방향**: Outlook ↔ Google 양방향 동기화
- **대상**: Outlook 기본 캘린더 ↔ Google primary 캘린더
- **범위**: 오늘 날짜 이후 미래 일정 전체 (과거 일정 동기화 안 함)
- **충돌 처리**: 최신 수정 시각(last modified) 우선
- **삭제 처리**: 사용자가 **직접** 한쪽에서 삭제한 경우에만 반대쪽에도 삭제 전파
- **삭제 금지**: 날짜 변경, 스캔 범위, 동기화 버그 등으로 프로그램이 임의 삭제하지 않음
- **Outlook → Google 삭제 조건**: `GetItemFromID`로 **실제 삭제 확인** + **2회 연속** 확인 후에만 삭제
- **Google → Outlook 삭제 조건**: Google API로 **cancelled/missing 재확인** + **2회 연속** 확인 후에만 삭제
- **COM/API 오류 시**: 삭제 전파 **보류** (unknown 상태)
- **원칙**: 한 번 동기화된 일정은 사용자가 삭제하기 전까지 Google·Outlook에 유지
- **반복 일정**: 개별 인스턴스로 분해하여 처리 (오늘부터 최대 365일까지, `SYNC_RECURRING_HORIZON_DAYS`)
- **실행 주기**: 30분마다 자동 실행
- **색상 동기화**: Google `colorId` ↔ Outlook `GSync-*` 카테고리 (Outlook에서 범주 색상은 수동 지정)
- 상세 구현은 `Docs/Finished/01_Outlook_Google_캘린더_동기화_기획.md` 참조

## CLI

```bash
cd Programs
python main.py                   # 동기화 (콘솔 출력, 수동 실행용)
python main.py --quiet           # 로그 파일만 (pythonw / 스케줄러용)
run_sync.vbs                     # 창 없이 실행 (Task Scheduler 권장)
python main.py --list-calendars  # Google 캘린더 ID 목록
python main.py --rebuild-map     # event_map 재구성 (중복 복구)
```

Task Scheduler는 **`wscript.exe` + `run_sync.vbs`** 로 등록 (CMD 창 안 뜸). 로그는 `Programs/logs/sync.log`.

## 로그·에러 처리

- 로그 파일은 `Programs/logs/` 하위에 저장한다.
- API 호출 실패 시 재시도 횟수와 대기 시간을 코드에 명시한다. (Google API HTTP 에러)
- 동기화 실패 시 로그에 원인(캘린더 ID, 이벤트 ID, HTTP 상태 등)을 남긴다.
- 로컬 상태 파일: `Programs/sync/last_sync.json`, `Programs/sync/event_map.json`

## Git

- **Remote**: `https://github.com/indychung1031/calendar_sync.git`
- 개발 완료 후 변경 사항을 **항상 git에 커밋**하고 **push**한다.
- 커밋 메시지 형식: `타입: 설명` (예: `feat: Google Calendar 읽기 API 연동`)
  - 타입: `feat`, `fix`, `docs`, `refactor`, `test`, `chore`
- `.env` 등 비밀 정보가 포함된 파일은 커밋하지 않는다.
