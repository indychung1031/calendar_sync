"""
event_mapper.py 단위 테스트.

순수 함수 테스트 — 외부 API 호출 없음.
"""

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent))

from calendar_api.event_mapper import (
    GOOGLE_COLOR_TO_GSYNC,
    GSYNC_TO_GOOGLE_COLOR,
    google_event_to_outlook,
    is_google_deleted,
    is_outlook_deleted,
    match_key_for_google,
    match_key_for_outlook,
    normalize_to_utc,
    outlook_event_to_google,
)


# ── normalize_to_utc ──────────────────────────────────────────────────────

class TestNormalizeToUtc:
    def test_rfc3339_with_offset(self):
        # KST(+09:00) → UTC
        result = normalize_to_utc("2025-06-29T09:00:00+09:00")
        assert result == "2025-06-29T00:00:00Z"

    def test_rfc3339_utc_z(self):
        result = normalize_to_utc("2025-06-29T00:00:00Z")
        assert result == "2025-06-29T00:00:00Z"

    def test_outlook_microseconds_format(self):
        # Outlook: ".0000000" 소수점 제거 후 UTC 변환
        result = normalize_to_utc("2025-06-29T09:00:00.0000000", "Asia/Seoul")
        assert result == "2025-06-29T00:00:00Z"

    def test_windows_tz_korea(self):
        result = normalize_to_utc("2025-06-29T09:00:00.0000000", "Korea Standard Time")
        assert result == "2025-06-29T00:00:00Z"

    def test_no_timezone_defaults_utc(self):
        result = normalize_to_utc("2025-06-29T05:00:00")
        assert result == "2025-06-29T05:00:00Z"

    def test_unknown_tz_fallback_utc(self):
        result = normalize_to_utc("2025-06-29T05:00:00", "Unknown/Zone")
        assert result == "2025-06-29T05:00:00Z"


# ── is_deleted ────────────────────────────────────────────────────────────

class TestIsDeleted:
    def test_google_cancelled(self):
        assert is_google_deleted({"id": "x", "status": "cancelled"})

    def test_google_confirmed_not_deleted(self):
        assert not is_google_deleted({"id": "x", "status": "confirmed"})

    def test_outlook_removed(self):
        assert is_outlook_deleted({"id": "x", "@removed": {"reason": "deleted"}})

    def test_outlook_active_not_deleted(self):
        assert not is_outlook_deleted({"id": "x", "subject": "Meeting"})


# ── match_key ─────────────────────────────────────────────────────────────

class TestMatchKey:
    def test_google_regular_event(self):
        event = {
            "summary": "회의",
            "start": {"dateTime": "2025-06-29T09:00:00+09:00", "timeZone": "Asia/Seoul"},
        }
        key = match_key_for_google(event)
        assert key == "회의||2025-06-29T00:00:00Z"

    def test_google_allday_event(self):
        event = {"summary": "휴가", "start": {"date": "2025-06-29"}}
        key = match_key_for_google(event)
        assert key == "휴가||2025-06-29||allday"

    def test_outlook_regular_event(self):
        event = {
            "subject": "회의",
            "isAllDay": False,
            "start": {"dateTime": "2025-06-29T00:00:00.0000000", "timeZone": "UTC"},
        }
        key = match_key_for_outlook(event)
        assert key == "회의||2025-06-29T00:00:00Z"

    def test_outlook_allday_event(self):
        event = {
            "subject": "휴가",
            "isAllDay": True,
            "start": {"dateTime": "2025-06-29T00:00:00.0000000", "timeZone": "UTC"},
        }
        key = match_key_for_outlook(event)
        assert key == "휴가||2025-06-29||allday"

    def test_google_missing_start_returns_none(self):
        assert match_key_for_google({"summary": "x", "start": {}}) is None


# ── google_event_to_outlook ───────────────────────────────────────────────

class TestGoogleToOutlook:
    def test_regular_event_fields(self):
        event = {
            "summary": "팀 회의",
            "start": {"dateTime": "2025-06-29T09:00:00+09:00", "timeZone": "Asia/Seoul"},
            "end": {"dateTime": "2025-06-29T10:00:00+09:00", "timeZone": "Asia/Seoul"},
            "description": "주간 업무 공유",
            "location": "회의실 A",
        }
        result = google_event_to_outlook(event)

        assert result["subject"] == "팀 회의"
        assert result["start"]["dateTime"] == "2025-06-29T00:00:00"
        assert result["start"]["timeZone"] == "UTC"
        assert result["end"]["dateTime"] == "2025-06-29T01:00:00"
        assert result["body"]["content"] == "주간 업무 공유"
        assert result["location"]["displayName"] == "회의실 A"
        assert "isAllDay" not in result

    def test_allday_event(self):
        event = {
            "summary": "연차",
            "start": {"date": "2025-06-29"},
            "end": {"date": "2025-06-30"},
        }
        result = google_event_to_outlook(event)

        assert result["isAllDay"] is True
        assert result["start"]["dateTime"] == "2025-06-29T00:00:00"
        assert result["end"]["dateTime"] == "2025-06-30T00:00:00"
        assert result["start"]["timeZone"] == "UTC"

    def test_empty_description_not_included(self):
        event = {
            "summary": "x",
            "start": {"dateTime": "2025-06-29T09:00:00Z"},
            "end": {"dateTime": "2025-06-29T10:00:00Z"},
        }
        result = google_event_to_outlook(event)
        assert "body" not in result

    def test_no_z_suffix_in_datetime(self):
        event = {
            "summary": "x",
            "start": {"dateTime": "2025-06-29T00:00:00Z"},
            "end": {"dateTime": "2025-06-29T01:00:00Z"},
        }
        result = google_event_to_outlook(event)
        # Outlook은 'Z' 없는 순수 ISO + timeZone 필드로 지정
        assert not result["start"]["dateTime"].endswith("Z")


# ── outlook_event_to_google ───────────────────────────────────────────────

class TestOutlookToGoogle:
    def test_regular_event_fields(self):
        event = {
            "subject": "팀 회의",
            "isAllDay": False,
            "start": {"dateTime": "2025-06-29T00:00:00.0000000", "timeZone": "UTC"},
            "end": {"dateTime": "2025-06-29T01:00:00.0000000", "timeZone": "UTC"},
            "body": {"contentType": "text", "content": "주간 업무 공유"},
            "location": {"displayName": "회의실 A"},
        }
        result = outlook_event_to_google(event)

        assert result["summary"] == "팀 회의"
        assert result["start"]["dateTime"] == "2025-06-29T00:00:00Z"
        assert result["start"]["timeZone"] == "UTC"
        assert result["end"]["dateTime"] == "2025-06-29T01:00:00Z"
        assert result["description"] == "주간 업무 공유"
        assert result["location"] == "회의실 A"

    def test_allday_event(self):
        event = {
            "subject": "연차",
            "isAllDay": True,
            "start": {"dateTime": "2025-06-29T00:00:00.0000000", "timeZone": "UTC"},
            "end": {"dateTime": "2025-06-30T00:00:00.0000000", "timeZone": "UTC"},
        }
        result = outlook_event_to_google(event)

        assert result["start"] == {"date": "2025-06-29"}
        assert result["end"] == {"date": "2025-06-30"}
        assert "dateTime" not in result["start"]

    def test_z_suffix_in_datetime(self):
        event = {
            "subject": "x",
            "isAllDay": False,
            "start": {"dateTime": "2025-06-29T00:00:00.0000000", "timeZone": "UTC"},
            "end": {"dateTime": "2025-06-29T01:00:00.0000000", "timeZone": "UTC"},
        }
        result = outlook_event_to_google(event)
        assert result["start"]["dateTime"].endswith("Z")

    def test_empty_body_not_included(self):
        event = {
            "subject": "x",
            "isAllDay": False,
            "start": {"dateTime": "2025-06-29T00:00:00.0000000", "timeZone": "UTC"},
            "end": {"dateTime": "2025-06-29T01:00:00.0000000", "timeZone": "UTC"},
            "body": {"contentType": "text", "content": ""},
        }
        result = outlook_event_to_google(event)
        assert "description" not in result


# ── 색상 동기화 ──────────────────────────────────────────────────────────

class TestColorSync:
    def _base_google_event(self, color_id: str | None = None) -> dict:
        event = {
            "summary": "색상 테스트",
            "start": {"dateTime": "2025-06-29T09:00:00Z"},
            "end": {"dateTime": "2025-06-29T10:00:00Z"},
        }
        if color_id:
            event["colorId"] = color_id
        return event

    def _base_outlook_event(self, categories: list[str] | None = None) -> dict:
        event = {
            "subject": "색상 테스트",
            "isAllDay": False,
            "start": {"dateTime": "2025-06-29T09:00:00.0000000", "timeZone": "UTC"},
            "end": {"dateTime": "2025-06-29T10:00:00.0000000", "timeZone": "UTC"},
        }
        if categories is not None:
            event["categories"] = categories
        return event

    def test_google_tomato_maps_to_gsync_tomato(self):
        result = google_event_to_outlook(self._base_google_event(color_id="10"))
        assert result.get("categories") == ["GSync-Tomato"]

    def test_google_peacock_maps_to_gsync_peacock(self):
        result = google_event_to_outlook(self._base_google_event(color_id="7"))
        assert result.get("categories") == ["GSync-Peacock"]

    def test_google_no_color_has_no_categories(self):
        result = google_event_to_outlook(self._base_google_event())
        assert "categories" not in result

    def test_outlook_gsync_tomato_maps_to_colorid_10(self):
        result = outlook_event_to_google(self._base_outlook_event(["GSync-Tomato"]))
        assert result.get("colorId") == "10"

    def test_outlook_first_gsync_category_wins(self):
        """GSync 카테고리가 여러 개여도 첫 번째만 사용."""
        result = outlook_event_to_google(
            self._base_outlook_event(["GSync-Tomato", "GSync-Peacock"])
        )
        assert result.get("colorId") == "10"

    def test_outlook_non_gsync_category_ignored(self):
        """사용자 기존 카테고리는 colorId에 영향 없음."""
        result = outlook_event_to_google(self._base_outlook_event(["업무", "중요"]))
        assert "colorId" not in result

    def test_outlook_mixed_categories_picks_gsync(self):
        """사용자 카테고리와 GSync 카테고리 혼합 시 GSync 우선."""
        result = outlook_event_to_google(
            self._base_outlook_event(["업무", "GSync-Sage"])
        )
        assert result.get("colorId") == "2"

    def test_all_11_google_colors_have_mapping(self):
        """Google 11가지 colorId 전부 매핑 테이블에 존재."""
        for i in range(1, 12):
            assert str(i) in GOOGLE_COLOR_TO_GSYNC

    def test_reverse_mapping_covers_all(self):
        """역변환 테이블이 정방향과 동일한 수의 항목 보유."""
        assert len(GSYNC_TO_GOOGLE_COLOR) == len(GOOGLE_COLOR_TO_GSYNC)

    def test_round_trip_color(self):
        """Google colorId → Outlook category → Google colorId 왕복 일관성."""
        for color_id in GOOGLE_COLOR_TO_GSYNC:
            g_event = self._base_google_event(color_id=color_id)
            outlook_body = google_event_to_outlook(g_event)
            google_body = outlook_event_to_google({**self._base_outlook_event(), **outlook_body})
            assert google_body.get("colorId") == color_id, f"colorId={color_id} 왕복 실패"


# ── 왕복 변환 일관성 ──────────────────────────────────────────────────────

class TestRoundTrip:
    def test_google_to_outlook_to_google_regular(self):
        """Google → Outlook → Google 변환 후 핵심 필드 보존."""
        original = {
            "summary": "왕복 테스트",
            "start": {"dateTime": "2025-06-29T09:00:00+09:00", "timeZone": "Asia/Seoul"},
            "end": {"dateTime": "2025-06-29T10:00:00+09:00", "timeZone": "Asia/Seoul"},
            "description": "설명",
            "location": "서울",
        }
        intermediate = google_event_to_outlook(original)

        # Outlook → Google 변환을 위해 isAllDay 추가
        intermediate["isAllDay"] = intermediate.get("isAllDay", False)
        restored = outlook_event_to_google(intermediate)

        assert restored["summary"] == original["summary"]
        assert restored["start"]["dateTime"] == "2025-06-29T00:00:00Z"
        assert restored["description"] == original["description"]
        assert restored["location"] == original["location"]
