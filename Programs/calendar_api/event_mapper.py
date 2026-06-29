"""
Google Calendar ↔ Outlook (MS Graph) 이벤트 형식 변환 및 UTC 정규화.

전일 일정(all-day)은 날짜 문자열만 사용하고 UTC 변환을 적용하지 않는다.
타임존이 다를 경우 날짜가 하루 어긋나는 문제를 방지하기 위함이다.
"""

import re
from datetime import datetime, timezone
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

# MS Graph가 반환하는 Windows 타임존 이름을 IANA 이름으로 변환
_WINDOWS_TO_IANA: dict[str, str] = {
    "Korea Standard Time": "Asia/Seoul",
    "Tokyo Standard Time": "Asia/Tokyo",
    "China Standard Time": "Asia/Shanghai",
    "Singapore Standard Time": "Asia/Singapore",
    "India Standard Time": "Asia/Kolkata",
    "AUS Eastern Standard Time": "Australia/Sydney",
    "Eastern Standard Time": "America/New_York",
    "Central Standard Time": "America/Chicago",
    "Mountain Standard Time": "America/Denver",
    "Pacific Standard Time": "America/Los_Angeles",
    "GMT Standard Time": "Europe/London",
    "W. Europe Standard Time": "Europe/Berlin",
    "UTC": "UTC",
}


def _resolve_tz(tz_name: str) -> ZoneInfo:
    """IANA 또는 Windows 타임존 이름으로 ZoneInfo 반환. 미지원 시 UTC 폴백."""
    try:
        return ZoneInfo(tz_name)
    except (ZoneInfoNotFoundError, KeyError):
        iana = _WINDOWS_TO_IANA.get(tz_name, "UTC")
        try:
            return ZoneInfo(iana)
        except (ZoneInfoNotFoundError, KeyError):
            return ZoneInfo("UTC")


def normalize_to_utc(dt_str: str, tz_name: str = "UTC") -> str:
    """
    datetime 문자열 → 'YYYY-MM-DDTHH:MM:SSZ' (UTC).

    Outlook의 .0000000 마이크로초 형식과 RFC 3339 오프셋 포함 형식 모두 처리.
    타임존 정보가 없는 경우 tz_name을 적용한다.
    """
    # Outlook: "2025-06-29T09:00:00.0000000" → 소수점 이하 제거
    clean = re.sub(r"\.\d+", "", dt_str)
    # "Z" → "+00:00" (Python < 3.11 fromisoformat 호환)
    clean = clean.replace("Z", "+00:00")

    try:
        dt = datetime.fromisoformat(clean)
    except ValueError:
        return dt_str  # 파싱 실패 시 원본 반환

    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=_resolve_tz(tz_name))

    return dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def is_google_deleted(event: dict) -> bool:
    return event.get("status") == "cancelled"


def is_outlook_deleted(event: dict) -> bool:
    return "@removed" in event


def match_key_for_google(event: dict) -> str | None:
    """최초 동기화 짝짓기 키: '제목||UTC시작시각' 또는 '제목||날짜||allday'"""
    title = event.get("summary", "")
    start = event.get("start", {})
    if "date" in start:
        return f"{title}||{start['date']}||allday"
    if "dateTime" in start:
        utc = normalize_to_utc(start["dateTime"], start.get("timeZone", "UTC"))
        return f"{title}||{utc}"
    return None


def match_key_for_outlook(event: dict) -> str | None:
    """최초 동기화 짝짓기 키: '제목||UTC시작시각' 또는 '제목||날짜||allday'"""
    title = event.get("subject", "")
    start = event.get("start", {})
    is_all_day = event.get("isAllDay", False)

    if is_all_day:
        # Prefer UTC 헤더로 받은 전일 일정의 dateTime은 "YYYY-MM-DDT00:00:00..."
        date_str = start.get("dateTime", "")[:10]
        return f"{title}||{date_str}||allday"
    if start.get("dateTime"):
        utc = normalize_to_utc(start["dateTime"], start.get("timeZone", "UTC"))
        return f"{title}||{utc}"
    return None


def google_event_to_outlook(event: dict) -> dict:
    """Google 이벤트 → Outlook 생성/수정 요청 body."""
    start = event.get("start", {})
    end = event.get("end", {})

    if "date" in start:
        body = {
            "subject": event.get("summary", ""),
            "isAllDay": True,
            "start": {"dateTime": f"{start['date']}T00:00:00", "timeZone": "UTC"},
            "end": {"dateTime": f"{end['date']}T00:00:00", "timeZone": "UTC"},
        }
    else:
        utc_start = normalize_to_utc(start["dateTime"], start.get("timeZone", "UTC"))
        utc_end = normalize_to_utc(end["dateTime"], end.get("timeZone", "UTC"))
        # Outlook은 'Z' 없이 순수 ISO 문자열 + timeZone 필드로 지정
        body = {
            "subject": event.get("summary", ""),
            "start": {"dateTime": utc_start.rstrip("Z"), "timeZone": "UTC"},
            "end": {"dateTime": utc_end.rstrip("Z"), "timeZone": "UTC"},
        }

    description = event.get("description", "")
    if description:
        body["body"] = {"contentType": "text", "content": description}

    location = event.get("location", "")
    if location:
        body["location"] = {"displayName": location}

    return body


def outlook_event_to_google(event: dict) -> dict:
    """Outlook 이벤트 → Google 생성/수정 요청 body."""
    start = event.get("start", {})
    end = event.get("end", {})
    is_all_day = event.get("isAllDay", False)

    body: dict = {"summary": event.get("subject", "")}

    if is_all_day:
        # Prefer UTC 헤더로 받은 전일 일정: dateTime에서 날짜만 추출
        body["start"] = {"date": start.get("dateTime", "")[:10]}
        body["end"] = {"date": end.get("dateTime", "")[:10]}
    else:
        utc_start = normalize_to_utc(start["dateTime"], start.get("timeZone", "UTC"))
        utc_end = normalize_to_utc(end["dateTime"], end.get("timeZone", "UTC"))
        body["start"] = {"dateTime": utc_start, "timeZone": "UTC"}
        body["end"] = {"dateTime": utc_end, "timeZone": "UTC"}

    content = event.get("body", {}).get("content", "")
    if content:
        body["description"] = content

    display_name = event.get("location", {}).get("displayName", "")
    if display_name:
        body["location"] = display_name

    return body
