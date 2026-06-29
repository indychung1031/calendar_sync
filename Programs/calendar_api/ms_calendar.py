"""MS Graph Calendar API 래퍼 (calendarView/delta 방식)."""

import logging
import time
from datetime import datetime, timedelta, timezone

import requests

logger = logging.getLogger(__name__)

_GRAPH_BASE = "https://graph.microsoft.com/v1.0"
_SELECT_FIELDS = (
    "id,subject,start,end,body,location,isAllDay,lastModifiedDateTime"
)
# MS Graph 응답 시각을 UTC로 받도록 요청
_PREFER_UTC = 'outlook.timezone="UTC"'
_MAX_RETRIES = 3


class MSCalendarClient:
    def __init__(
        self,
        access_token: str,
        calendar_id: str,
        horizon_days: int,
    ):
        self._headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
            "Prefer": _PREFER_UTC,
        }
        self._cal_id = calendar_id  # 비어 있으면 기본 캘린더 사용
        self._horizon_days = horizon_days

    # ── URL 헬퍼 ──────────────────────────────────────────────────────────

    def _cal_base(self) -> str:
        if self._cal_id:
            return f"{_GRAPH_BASE}/me/calendars/{self._cal_id}"
        return f"{_GRAPH_BASE}/me/calendar"

    def _view_delta_url(self) -> str:
        return f"{self._cal_base()}/calendarView/delta"

    def _events_url(self) -> str:
        return f"{self._cal_base()}/events"

    # ── 이벤트 조회 ───────────────────────────────────────────────────────

    def list_events_initial(self) -> tuple[list[dict], str]:
        """
        전체 초기 조회. (이벤트 목록, deltaLink) 반환.

        calendarView/delta는 startDateTime·endDateTime 필수.
        """
        now = datetime.now(timezone.utc)
        end = now + timedelta(days=self._horizon_days)
        params = {
            "startDateTime": now.isoformat(),
            "endDateTime": end.isoformat(),
            "$select": _SELECT_FIELDS,
        }
        return self._paginate_delta(self._view_delta_url(), params)

    def list_events_delta(self, delta_link: str) -> tuple[list[dict], str]:
        """증분 조회. deltaLink URL 자체에 파라미터 포함됨."""
        return self._paginate_delta(delta_link, {})

    def _paginate_delta(
        self, url: str, params: dict
    ) -> tuple[list[dict], str]:
        events: list[dict] = []
        delta_link: str = ""
        first_request = True

        while url:
            resp = self._get(url, params=params if first_request else None)
            resp.raise_for_status()
            data = resp.json()
            events.extend(data.get("value", []))
            first_request = False

            if "@odata.deltaLink" in data:
                delta_link = data["@odata.deltaLink"]
                break
            url = data.get("@odata.nextLink", "")

        return events, delta_link

    # ── CRUD ──────────────────────────────────────────────────────────────

    def create_event(self, body: dict) -> dict:
        resp = self._post(self._events_url(), json=body)
        resp.raise_for_status()
        return resp.json()

    def update_event(self, event_id: str, body: dict) -> dict:
        url = f"{_GRAPH_BASE}/me/events/{event_id}"
        resp = self._patch(url, json=body)
        resp.raise_for_status()
        return resp.json()

    def delete_event(self, event_id: str) -> None:
        url = f"{_GRAPH_BASE}/me/events/{event_id}"
        resp = self._delete(url)
        if resp.status_code == 404:
            logger.debug("MS 이벤트 이미 삭제됨 (id=%s)", event_id)
            return
        resp.raise_for_status()

    # ── HTTP 헬퍼 (재시도 + 429 처리) ─────────────────────────────────────

    def _get(self, url: str, **kwargs) -> requests.Response:
        return self._request("GET", url, **kwargs)

    def _post(self, url: str, **kwargs) -> requests.Response:
        return self._request("POST", url, **kwargs)

    def _patch(self, url: str, **kwargs) -> requests.Response:
        return self._request("PATCH", url, **kwargs)

    def _delete(self, url: str, **kwargs) -> requests.Response:
        return self._request("DELETE", url, **kwargs)

    def _request(self, method: str, url: str, **kwargs) -> requests.Response:
        for attempt in range(_MAX_RETRIES):
            resp = requests.request(
                method, url, headers=self._headers, **kwargs
            )

            if resp.status_code == 429:
                # Rate limit: Retry-After 헤더 우선, 없으면 지수 백오프
                wait = int(resp.headers.get("Retry-After", 2 ** attempt))
                logger.warning("MS Graph rate limit. %d초 대기 후 재시도", wait)
                time.sleep(wait)
                continue

            if resp.status_code == 503:
                wait = 2 ** attempt
                logger.warning("MS Graph 503. %d초 대기 후 재시도", wait)
                time.sleep(wait)
                continue

            return resp

        logger.error("MS Graph 요청 최대 재시도 초과: %s %s", method, url)
        return resp  # 마지막 응답 반환 (호출 측에서 raise_for_status)
