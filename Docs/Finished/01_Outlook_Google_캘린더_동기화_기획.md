# 01. Outlook ↔ Google 캘린더 양방향 동기화 기획서

> **상태**: ✅ 구현·실행 완료 (2026-06-30)  
> **아키텍처 변경**: Outlook 연동을 MS Graph API → **로컬 Outlook COM(pywin32)** 으로 전환 (Azure 앱 등록 불필요)

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
| Microsoft Outlook | **`pywin32` COM** (로컬 Outlook 앱 직접 제어) |
| 스케줄러 | Windows Task Scheduler |
| 환경 변수 | `python-dotenv`, `tzdata` |
| 로깅 | Python 표준 `logging` 모듈 |

### 아키텍처 결정: Outlook COM 방식

| | MS Graph (기존 기획) | Outlook COM (최종 구현) |
|---|---|---|
| Azure 앱 등록 | 필요 | **불필요** |
| 인증 | MSAL device code flow | Outlook 앱이 처리 |
| 계정 유형 | M365 위주 | POP/IMAP/Exchange 등 **모든 Outlook 계정** |
| delta query | 지원 | 미지원 → 전체 스캔 + event_map 기반 삭제 감지 |
| 색상 동기화 | 카테고리 API | GSync-* 카테고리 (색상은 Outlook에서 수동 지정) |

---

## 4. 동기화 정책

> ✅ **확정됨** (2026-06-29) — 아래 정책으로 구현한다.

### 4-1. 동기화 방향
- **양방향**: Outlook ↔ Google 양쪽 모두 읽고 쓴다.

### 4-2. 동기화 범위 (기간)
- **오늘 날짜 00:00(로컬) 이후** ~ horizon 일까지 동기화
- **'지금(now)'이 아닌 '오늘 자정'** 기준 — 당일 일정은 하루 종일 유지
- 어제 이전 일정은 동기화 대상에서 제외 (과거 일정은 새로 가져오지 않음)

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
- **Google**: OAuth 2.0 Installed App flow (`credentials.json` → `token_google.json`)
- **Outlook**: 별도 인증 없음 — 로컬 Outlook 앱 COM 연결 (`Outlook.Application`)
  - Outlook이 실행 중이거나 백그라운드로 동작해야 함
  - `auth/ms_auth.py`는 COM 전환 이후 **미사용** (deprecated 주석만 유지)
- Google 토큰 만료 시 `google-auth`가 자동 refresh

### 5-2. 이벤트 읽기
- **Google** (`calendar_api/google_calendar.py`): bounded-window + `updatedMin` 증분
  - `timeMin=오늘`, `timeMax=오늘+HORIZON`, `singleEvents=True`
  - 증분: `updatedMin` + `showDeleted=True` (삭제는 `status: cancelled`)
- **Outlook** (`calendar_api/ms_calendar.py`): COM 전체 스캔
  - `IncludeRecurrences=True` + `Sort([Start])` + Restrict 날짜 필터
  - delta 미지원 → `ms_delta_link`는 `"outlook_com_v1"` 마커만 저장
  - 삭제 감지: 스캔 목록에 없을 때 `GetItemFromID`로 **실제 삭제 여부 확인** 후 Google 삭제

### 5-3. 이벤트 변환
- Google 이벤트 형식 ↔ Outlook 이벤트 형식 간 변환 (`event_mapper.py`)
- 매핑 필드: 제목, 시작/종료 시각, 설명, 장소, 전일 일정 여부, **색상**

#### 색상 동기화
- Google `colorId` (11색) → Outlook `GSync-*` 카테고리 이름 (예: `GSync-Tomato`)
- 역방향: Outlook 카테고리 중 `GSync-*` 첫 번째 매칭 → Google `colorId`
- 사용자 기존 카테고리는 유지, GSync 카테고리만 추가/갱신
- Outlook 범주 **색상**은 COM으로 자동 지정 불가 → Outlook UI에서 수동 설정

#### 타임존 처리 정책
- 모든 시각 비교·저장은 **UTC 기준으로 정규화**한다.
- Google: `dateTime` 필드가 RFC 3339 오프셋 포함 → UTC 변환 후 처리
- Outlook: `dateTime` 필드(로컬 시각) + `timeZone` 필드 조합 → UTC 변환 후 처리
- **전일 일정(all-day)**:
  - Google: `date` 필드만 존재 (예: `"2025-06-29"`)
  - Outlook: `isAllDay: true` + 시각을 UTC 00:00으로 저장
  - 변환 시 날짜 문자열(`YYYY-MM-DD`)만 사용하고 UTC 변환 적용 안 함 (타임존에 따라 날짜가 하루 어긋나는 문제 방지)

### 5-4. 동기화 실행

#### 로컬 상태 파일 (`Programs/sync/`)

| 파일 | 용도 |
|---|---|
| `last_sync.json` | 마지막 동기화 시각, `google_updated_min`, `ms_delta_link`(COM 마커) |
| `event_map.json` | Google ↔ Outlook ID 매핑 + 양쪽 마지막 수정 시각 |
| `sync.lock` | 동시 실행 방지용 잠금 파일 (실행 중에만 존재) |

**event_map.json 구조:**
```json
{
  "google_event_id_123": {
    "outlook_id": "outlook_event_id_456",
    "google_modified": "2025-06-29T10:00:00Z",
    "outlook_modified": "2025-06-29T10:00:00Z"
  }
}
```
- `google_modified` / `outlook_modified`: 우리가 마지막으로 동기화한 시점의 수정 시각
- 무한 루프 방지에 사용 (§ 증분 동기화 참조)

#### 이벤트 매칭 (ID 연결)
- 한쪽에서 생성한 일정을 반대쪽에 복사할 때 `event_map.json`에 양쪽 ID와 수정 시각을 저장
- 이후 수정·삭제는 매핑 테이블로 대응 이벤트를 찾아 처리
- 매핑이 없는 이벤트는 **신규**로 간주

#### 동시 실행 방지
- 실행 시작 시 `sync.lock` 파일 생성 (내용: PID, 시작 시각)
- 실행 종료 시 `sync.lock` 삭제
- 시작 시 `sync.lock`이 이미 존재하면 즉시 종료 (중복 실행 방지)
- 단, lock 파일 생성 후 비정상 종료된 경우를 대비해 파일 내 시작 시각이 2시간 이상 경과했으면 무시하고 덮어씀

#### event_map.json 손상·삭제 복구 전략
- 파일 손상·삭제 시 모든 이벤트를 신규로 인식하여 **양쪽 캘린더에 이벤트가 중복 생성**됨
- 복구 절차:
  1. 양쪽 캘린더에서 중복 이벤트를 수동 정리
  2. `python main.py --rebuild-map` 으로 제목+시각 기반 재짝짓기 실행
- 예방: `event_map.json` 변경 시마다 `.bak` 파일로 백업

#### 최초 동기화 (첫 실행)
1. 양쪽 캘린더에서 동기화 범위(오늘 이후) 이벤트를 모두 조회
2. **자동 짝짓기**: 제목 + **UTC 기준** 시작 시각 + 종료 시각이 동일한 이벤트를 같은 일정으로 판단하여 `event_map.json`에 등록
3. 짝이 없는 이벤트는 반대쪽에 생성 후 매핑 등록
4. 짝이 2개 이상 매칭되면 로그에 경고 남기고 수동 확인 대상으로 표시 (자동 병합 안 함)

#### 증분 동기화 (2회차 이후)
- Google: `updatedMin`으로 변경 이벤트만 조회
- Outlook: COM 전체 스캔 후 `event_map`과 비교
- **Outlook 삭제 감지**: `_detect_and_apply_outlook_deletions()` — map에 있으나 스캔 목록에 없는 ID
- **무한 루프 방지**: 서버 반환 수정 시각이 `event_map` 기록과 동일하면 건너뜀
- 충돌: 양쪽 모두 변경 시 last modified 우선
- **삭제 처리**:
  1. Google `cancelled` → Outlook 삭제
  2. Outlook 목록에서 사라짐 → Google 삭제
  3. `event_map`에서 해당 항목 즉시 제거

### 5-5. CLI (`main.py`)

| 명령 | 설명 |
|---|---|
| `python main.py` | 일반 동기화 (최초 또는 증분) |
| `python main.py --list-calendars` | Google 캘린더 목록·ID 출력 |
| `python main.py --rebuild-map` | event_map 재구성 (중복 복구) |

### 5-6. 로깅
- 동기화 시작/종료 시각 기록
- 처리된 이벤트 수 (추가/수정/삭제) 기록
- 에러 발생 시 원인(캘린더 ID, 이벤트 ID, HTTP 상태) 기록

---

## 6. 디렉토리 구조 (구현 완료)

```
Programs/
├── main.py
├── requirements.txt
├── config.py
├── auth/
│   ├── google_auth.py       # Google OAuth
│   └── ms_auth.py           # (deprecated) COM 전환 후 미사용
├── calendar_api/
│   ├── google_calendar.py
│   ├── ms_calendar.py       # OutlookComClient
│   └── event_mapper.py
├── sync/
│   ├── sync_engine.py
│   ├── state.py
│   ├── last_sync.json       # (런타임)
│   ├── event_map.json       # (런타임)
│   ├── event_map.json.bak   # (런타임)
│   └── sync.lock            # (런타임)
├── logs/
│   └── sync.log
└── test_code/
    ├── test_event_mapper.py
    ├── test_state.py
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

# Microsoft Outlook (COM — Azure 앱 등록 불필요)
MS_CALENDAR_ID=    # 비워두면 기본 캘린더

# 동기화 설정
SYNC_START_FROM_TODAY=true
SYNC_RECURRING_HORIZON_DAYS=365
```

---

## 8. API 제한 사항

| API | 제한 | 대응 |
|---|---|---|
| Google Calendar | 1,000,000 req/day | `updatedMin` 증분 조회 |
| Outlook COM | 로컬 제한 없음 | 전체 스캔 — 이벤트 수 많을 시 실행 시간 증가 |

---

## 9. 작업 체크리스트

### 인증
- [x] Google OAuth 2.0 인증 및 토큰 저장 (`auth/google_auth.py`)
- [x] Outlook COM 연결 (`calendar_api/ms_calendar.py`) — 별도 MS 인증 불필요
- [x] Google 토큰 자동 갱신

### 캘린더 읽기/쓰기
- [x] Google Calendar 이벤트 조회 (bounded-window, updatedMin 증분)
- [x] Google Calendar 이벤트 생성/수정/삭제
- [x] Outlook COM 이벤트 조회 (전체 스캔, 반복 일정 전개)
- [x] Outlook COM 이벤트 생성/수정/삭제

### 이벤트 변환
- [x] Google → Outlook 필드 매핑 구현
- [x] Outlook → Google 필드 매핑 구현
- [x] 일반 일정: UTC 정규화 후 변환
- [x] 전일 일정(all-day): 날짜 문자열만 사용
- [x] 색상 동기화: Google colorId ↔ GSync-* 카테고리 (11색 매핑)

### 동기화 엔진
- [x] `event_map.json` 저장/로드 + `.bak` 백업
- [x] `last_sync.json` 저장/로드
- [x] `sync.lock` 동시 실행 방지 (2시간 stale lock 해제)
- [x] 최초 동기화: 제목 + UTC 시각 기반 자동 짝짓기
- [x] Google `updatedMin` 증분 조회
- [x] Outlook COM 전체 스캔 + event_map 기반 삭제 감지
- [x] 무한 루프 방지 (수정 시각 비교)
- [x] 충돌 감지 및 last modified 우선 처리
- [x] 삭제 전파 후 event_map 항목 즉시 제거
- [x] `--rebuild-map` 플래그
- [x] `--list-calendars` 플래그

### 로그·에러 처리
- [x] 로그 파일 설정 (`Programs/logs/sync.log`)
- [x] 동기화 결과 요약 로그 출력

### 테스트
- [x] Google 인증 — 실제 credentials로 수동 확인 완료 (2026-06-30)
- [x] Outlook COM 연결 — 로컬 Outlook으로 수동 확인 완료 (2026-06-30)
- [x] 이벤트 변환 단위 테스트 (34개)
- [x] 상태·락 단위 테스트 (18개)
- [x] 동기화 엔진 통합 테스트 Mock 기반 (16개)
- [x] **합계 68/68 통과** (2026-06-30)

### 배포
- [x] Windows Task Scheduler 등록 방법 문서화 (§12)
- [x] `.env.example` 작성 완료
- [x] `requirements.txt` 작성 완료

---

## 10. 완료 기준 (Definition of Done)

- [x] 기획서 체크리스트 전부 완료
- [x] `Programs/test_code/`에서 테스트 통과 (68/68)
- [x] git 커밋 및 push
- [x] 기획서를 `Docs/Finished/`로 이동

---

## 11. 실행 검증 기록

| 일시 | 내용 | 결과 |
|---|---|---|
| 2026-06-30 15:05 | Google OAuth 최초 인증 | ✅ `token_google.json` 저장 |
| 2026-06-30 15:06 | 증분 동기화 1회차 | ✅ 추가 0건 |
| 2026-06-30 15:37 | 최초 동기화 (실데이터) | ✅ **추가 26건** / 오류 0건 |
| 2026-06-30 | pytest 전체 | ✅ **68 passed** |

로그 위치: `Programs/logs/sync.log`

---

## 12. Windows Task Scheduler 등록

1. **작업 만들기** → 이름: `CalendarSync`
2. **트리거**: 30분마다 반복
3. **동작**: 프로그램 시작
   - 프로그램: `C:\Users\<사용자>\AppData\Local\Programs\Python\Python313\python.exe` (본인 환경에 맞게)
   - 인수: `main.py`
   - 시작 위치: `C:\Users\indyc\Desktop\antigravity\project\calendar_sync\Programs`
4. **조건**: "컴퓨터의 AC 전원이 켜져 있을 때만" 해제 (노트북 배터리에서도 실행)
5. **설정**: "이미 실행 중인 작업 적용 규칙" → **새 인스턴스 시작 안 함** (`sync.lock`과 이중 방어)

> Outlook이 로그인된 상태여야 COM 연결이 성공합니다.

---

## 13. 버그 수정 이력

| 일시 | 증상 | 원인 | 수정 |
|---|---|---|---|
| 2026-07-02 | 7/2 일정이 7/2에 Google에서 사라짐 | 동기화 범위가 `now` 기준이라 당일 일정이 스캔에서 제외 → 삭제로 오인 | `sync_window.py` 도입 (오늘 00:00~), `event_exists()` 확인 추가 |
