"""Google Calendar API 래퍼 (bounded-window 방식)."""

import logging
from datetime import datetime, timedelta, timezone

from calendar_api.sync_window import sync_bounds
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials

logger = logging.getLogger(__name__)

_MAX_RESULTS = 2500


class GoogleCalendarClient:
    def __init__(
        self,
        credentials: Credentials,
        calendar_id: str,
        horizon_days: int,
    ):
        self._service = build("calendar", "v3", credentials=credentials)
        self._cal_id = calendar_id
        self._horizon_days = horizon_days

    def _window(self) -> tuple[str, str]:
        """동기화 윈도우: [로컬 오늘 00:00, 오늘+horizon+1일 00:00) ISO 문자열."""
        window_start, window_end = sync_bounds(self._horizon_days)
        # Google API: timeMax는 exclusive — 자정 경계 사용
        start_local = window_start.astimezone()
        end_local = window_end.astimezone()
        return start_local.isoformat(), end_local.isoformat()

    def list_events(self, updated_min: str | None = None) -> list[dict]:
        """
        동기화 범위 이벤트 조회.

        updated_min 지정 시 증분 조회 (변경·삭제 포함),
        미지정 시 현재 유효한 이벤트 전체 조회.
        """
        time_min, time_max = self._window()
        params: dict = {
            "calendarId": self._cal_id,
            "timeMin": time_min,
            "timeMax": time_max,
            "singleEvents": True,   # 반복 일정을 개별 인스턴스로 전개
            "showDeleted": updated_min is not None,
            "maxResults": _MAX_RESULTS,
        }
        if updated_min:
            params["updatedMin"] = updated_min

        events: list[dict] = []
        page_token: str | None = None

        while True:
            if page_token:
                params["pageToken"] = page_token
            try:
                result = self._service.events().list(**params).execute()
            except HttpError as e:
                logger.error("Google Calendar 이벤트 조회 실패: %s", e)
                raise

            events.extend(result.get("items", []))
            page_token = result.get("nextPageToken")
            if not page_token:
                break

        return events

    def get_event(self, event_id: str) -> dict | None:
        """ID로 일정 조회 (동기화 윈도우 무관). 없으면 None."""
        try:
            return (
                self._service.events()
                .get(calendarId=self._cal_id, eventId=event_id)
                .execute()
            )
        except HttpError as e:
            if e.resp.status in (404, 410):
                return None
            logger.warning("Google 이벤트 조회 실패 (id=%s): %s", event_id, e)
            raise

    def check_event_status(self, event_id: str) -> str:
        """Google 일정 상태: active | cancelled | missing | unknown."""
        try:
            event = self.get_event(event_id)
            if event is None:
                return "missing"
            if event.get("status") == "cancelled":
                return "cancelled"
            return "active"
        except HttpError:
            logger.warning(
                "Google 존재 확인 실패 (id=%s) — 삭제 전파 보류", event_id
            )
            return "unknown"

    def create_event(self, body: dict) -> dict:
        try:
            return (
                self._service.events()
                .insert(calendarId=self._cal_id, body=body)
                .execute()
            )
        except HttpError as e:
            logger.error("Google 이벤트 생성 실패: %s", e)
            raise

    def update_event(self, event_id: str, body: dict) -> dict:
        try:
            return (
                self._service.events()
                .update(calendarId=self._cal_id, eventId=event_id, body=body)
                .execute()
            )
        except HttpError as e:
            logger.error("Google 이벤트 수정 실패 (id=%s): %s", event_id, e)
            raise

    def list_calendars(self) -> list[dict]:
        """사용자 Google 캘린더 목록 조회 (--list-calendars 용)."""
        try:
            result = self._service.calendarList().list().execute()
            return result.get("items", [])
        except HttpError as e:
            logger.error("Google 캘린더 목록 조회 실패: %s", e)
            raise

    def delete_event(self, event_id: str) -> None:
        try:
            self._service.events().delete(
                calendarId=self._cal_id, eventId=event_id
            ).execute()
        except HttpError as e:
            if e.resp.status == 410:
                # 이미 삭제된 이벤트: 정상 처리
                logger.debug("Google 이벤트 이미 삭제됨 (id=%s)", event_id)
                return
            logger.error("Google 이벤트 삭제 실패 (id=%s): %s", event_id, e)
            raise
