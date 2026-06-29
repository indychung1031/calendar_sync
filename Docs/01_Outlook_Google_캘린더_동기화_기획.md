# 01. Outlook ↔ Google 캘린더 양방향 동기화 기획서

## 1. 프로젝트 개요

Microsoft Outlook 캘린더와 Google 캘린더 간의 양방향 동기화 프로그램.  
두 캘린더에 각각 등록된 일정을 자동으로 동기화하여 항상 동일한 상태를 유지한다.

---

## 2. 목표

- Outlook에 일정을 추가하면 Google에도 자동 반영
- Google에 일정을 추가하면 Outlook에도 자동 반영
- 수정·삭제도 양방향으로 동기화
- 사람이 직접 조작하지 않아도 스케줄러가 주기적으로 자동 실행

---

## 3. 기술 스택

| 항목 | 선택 |
|---|---|
| 언어 | Python 3.12+ |
| Google Calendar | `google-api-python-client`, `google-auth-oauthlib` |
| Microsoft Outlook | `msal` (Microsoft Authentication Library) + MS Graph API (REST) |
| 스케줄러 | Windows Task Scheduler |
| 환경 변수 | `python-dotenv` |
| 로깅 | Python 표준 `logging` 모듈 |

---

## 4. 동기화 정책

> ✅ **확정됨** (2026-06-29) — 아래 정책으로 구현한다.

### 4-1. 동기화 방향
- **양방향**: Outlook ↔ Google 양쪽 모두 읽고 쓴다.

### 4-2. 동기화 범위 (기간)
- **오늘 날짜 이후 미래 일정 전체** 동기화
- 과거 일정은 동기화하지 않는다.

### 4-3. 충돌 처리 (같은 일정이 양쪽에서 동시에 수정된 경우)
- **최신 수정 시각(last modified) 우선**

### 4-4. 삭제 처리
- 한쪽에서 삭제 시 **반대쪽도 삭제**
- ⚠️ 실수로 삭제한 일정이 반대편에서도 사라질 수 있음 — **위험 수용** 정책으로 진행
- 동기화 지연으로 인한 오삭제 방지: 마지막 동기화 이후 실제 API에서 `deleted`로 확인된 이벤트만 삭제 전파

### 4-5. 반복 일정 (Recurring Events)
- 반복 일정을 **개별 인스턴스**로 분해하여 처리
- 이유: 두 캘린더의 반복 규칙(RRULE) 형식이 달라 변환 복잡도가 높음
- **전개 상한**: 오늘부터 **365일** 이내 (`SYNC_RECURRING_HORIZON_DAYS`, `.env`로 조정 가능)

### 4-6. 동기화 주기
- **30분마다** 자동 실행

### 4-7. 동기화 대상 캘린더
- Outlook: 기본 캘린더 (Default Calendar)
- Google: 기본 캘린더 (primary)

---

## 5. 주요 기능

### 5-1. 인증 관리
- Google OAuth 2.0 인증 및 토큰 갱신 (`Installed App` flow, `token_google.json` 저장)
- Microsoft OAuth 2.0 (MSAL) 인증 및 토큰 갱신
  - **로컬 PC 단독 실행**: Azure 앱을 **Public client**로 등록, **device code flow** 사용 (client secret 불필요)
  - `MS_TENANT_ID=common` (개인 Microsoft 계정) 또는 조직 tenant ID
- 토큰 만료 시 자동 갱신

### 5-2. 이벤트 읽기
- Google Calendar API로 동기화 범위 내 이벤트 조회
- MS Graph API로 동기화 범위 내 이벤트 조회

### 5-3. 이벤트 변환
- Google 이벤트 형식 ↔ Outlook 이벤트 형식 간 변환
- 매핑 필드: 제목, 시작/종료 시각, 설명, 장소, 전일 일정 여부

### 5-4. 동기화 실행

#### 로컬 상태 파일 (`Programs/sync/`)

| 파일 | 용도 |
|---|---|
| `last_sync.json` | 마지막 동기화 시각, 처리 통계 |
| `event_map.json` | Google 이벤트 ID ↔ Outlook 이벤트 ID 매핑 테이블 |

#### 이벤트 매칭 (ID 연결)
- 한쪽에서 생성한 일정을 반대쪽에 복사할 때 `event_map.json`에 양쪽 ID를 저장
- 이후 수정·삭제는 매핑 테이블로 대응 이벤트를 찾아 처리
- 매핑이 없는 이벤트는 **신규**로 간주

#### 최초 동기화 (첫 실행)
1. 양쪽 캘린더에서 동기화 범위(오늘 이후) 이벤트를 모두 조회
2. **자동 짝짓기**: 제목 + 시작 시각 + 종료 시각이 동일한 이벤트를 같은 일정으로 판단하여 `event_map.json`에 등록
3. 짝이 없는 이벤트는 반대쪽에 생성 후 매핑 등록
4. 짝이 2개 이상 매칭되면 로그에 경고 남기고 수동 확인 대상으로 표시 (자동 병합 안 함)

#### 증분 동기화 (2회차 이후)
- 마지막 동기화 시각 이후 변경된 이벤트만 처리
- 충돌 감지: 매핑된 양쪽 이벤트가 모두 수정된 경우 → last modified 우선
- 삭제 전파: API에서 삭제 확인된 이벤트만 반대쪽 삭제 (§4-4)

#### 기타
- 충돌 감지 및 정책 적용

### 5-5. 로깅
- 동기화 시작/종료 시각 기록
- 처리된 이벤트 수 (추가/수정/삭제) 기록
- 에러 발생 시 원인(캘린더 ID, 이벤트 ID, HTTP 상태) 기록

---

## 6. 디렉토리 구조 (예상)

```
Programs/
├── main.py                  # 진입점
├── requirements.txt
├── config.py                # 설정값 로드
├── auth/
│   ├── google_auth.py       # Google OAuth 처리
│   └── ms_auth.py           # Microsoft MSAL 처리
├── calendar/
│   ├── google_calendar.py   # Google Calendar API 래퍼
│   ├── ms_calendar.py       # MS Graph Calendar API 래퍼
│   └── event_mapper.py      # 이벤트 형식 변환
├── sync/
│   ├── sync_engine.py       # 동기화 핵심 로직
│   ├── state.py             # 마지막 동기화 상태 관리
│   ├── last_sync.json       # (런타임) 마지막 동기화 시각
│   └── event_map.json       # (런타임) Google ↔ Outlook ID 매핑
├── logs/                    # 로그 파일 저장
└── test_code/               # 테스트 코드
    ├── test_google_auth.py
    ├── test_ms_auth.py
    └── test_sync_engine.py
```

---

## 7. 환경 변수 (.env.example)

프로젝트 루트 `.env.example`과 동일. 주요 키:

```
# Google Calendar
GOOGLE_CREDENTIALS_PATH=credentials.json
GOOGLE_TOKEN_PATH=token_google.json
GOOGLE_CALENDAR_ID=primary

# Microsoft Outlook
MS_CLIENT_ID=
MS_TENANT_ID=common
MS_TOKEN_PATH=token_ms.json
MS_CALENDAR_ID=

# 동기화 설정
SYNC_START_FROM_TODAY=true
SYNC_RECURRING_HORIZON_DAYS=365
```

---

## 8. API 제한 사항

| API | 제한 | 대응 |
|---|---|---|
| Google Calendar | 1,000,000 req/day | 변경된 이벤트만 조회 |
| MS Graph | 10,000 req/10min | 지수 백오프(exponential backoff) 재시도 |

---

## 9. 작업 체크리스트

### 인증
- [ ] Google OAuth 2.0 인증 구현 및 토큰 저장
- [ ] MS MSAL 인증 구현 및 토큰 저장
- [ ] 토큰 자동 갱신 처리

### 캘린더 읽기/쓰기
- [ ] Google Calendar 이벤트 조회 (범위 내)
- [ ] Google Calendar 이벤트 생성/수정/삭제
- [ ] MS Graph 이벤트 조회 (범위 내)
- [ ] MS Graph 이벤트 생성/수정/삭제

### 이벤트 변환
- [ ] Google → Outlook 필드 매핑 구현
- [ ] Outlook → Google 필드 매핑 구현
- [ ] 전일 일정(all-day) 처리

### 동기화 엔진
- [ ] `event_map.json` 저장/로드 (Google ↔ Outlook ID 매핑)
- [ ] `last_sync.json` 저장/로드
- [ ] 최초 동기화: 제목+시각 기반 자동 짝짓기
- [ ] 변경 감지 로직 구현
- [ ] 충돌 감지 및 처리 로직 구현
- [ ] 삭제 동기화 처리 (API 삭제 확인 후 전파)

### 로그·에러 처리
- [ ] 로그 파일 설정 (Programs/logs/)
- [ ] API 에러 재시도 로직 (지수 백오프)
- [ ] 동기화 결과 요약 로그 출력

### 테스트
- [ ] Google 인증 테스트
- [ ] MS 인증 테스트
- [ ] 이벤트 변환 단위 테스트
- [ ] 동기화 엔진 통합 테스트

### 배포
- [ ] Windows Task Scheduler 등록 방법 문서화
- [ ] .env.example 작성 완료
- [ ] requirements.txt 작성 완료

---

## 10. 완료 기준 (Definition of Done)

- [ ] 기획서 체크리스트 전부 완료
- [ ] `Programs/test_code/`에서 테스트 통과
- [ ] git 커밋 및 push
- [ ] 기획서를 `Docs/Finished/`로 이동
