"""Outlook 로컬 캘린더 래퍼 (win32com COM 방식).

MS Graph API 대신 로컬에 설치된 Outlook 앱을 직접 제어한다.
POP/SMTP, IMAP, Exchange 등 모든 Outlook 계정 유형에서 동작하며
별도 Azure 앱 등록이나 인증 절차가 필요 없다.
"""

import logging
import re
from datetime import datetime, timedelta, timezone

from calendar_api.sync_window import sync_bounds

import pythoncom
import win32com.client

logger = logging.getLogger(__name__)

_OL_FOLDER_CALENDAR = 9    # olFolderCalendar
_OL_APPOINTMENT_ITEM = 1   # olAppointmentItem (CreateItem 용)
_OL_APPOINTMENT_CLASS = 26 # olAppointment (item.Class 비교 용)

_HAS_TZ = re.compile(r"[Zz]|[+-]\d{2}:\d{2}$")


def _pywint_to_utc(dt) -> str:
    """pywintypes.datetime(로컬 시각) → UTC ISO 문자열 ('YYYY-MM-DDTHH:MM:SSZ')."""
    naive = datetime(dt.year, dt.month, dt.day, dt.hour, dt.minute, dt.second)
    utc = naive.astimezone().astimezone(timezone.utc)
    return utc.strftime("%Y-%m-%dT%H:%M:%SZ")


def _to_com_local(dt_str: str, tz_name: str = "UTC") -> str:
    """datetime 문자열 → Outlook COM 설정용 로컬 시각 문자열 ('YYYY-MM-DD HH:MM:SS').

    타임존 정보가 없으면 tz_name을 UTC로 처리한다.
    """
    from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

    clean = re.sub(r"\.\d+", "", dt_str)
    if not _HAS_TZ.search(clean):
        try:
            tz = ZoneInfo(tz_name)
        except (ZoneInfoNotFoundError, KeyError):
            tz = timezone.utc
        dt = datetime.fromisoformat(clean).replace(tzinfo=tz)
    else:
        clean = clean.replace("Z", "+00:00")
        dt = datetime.fromisoformat(clean)
    return dt.astimezone().strftime("%Y-%m-%d %H:%M:%S")


class OutlookComClient:
    """Outlook COM 인터페이스를 통해 로컬 캘린더를 읽고 쓴다.

    delta query를 지원하지 않으므로 항상 전체 스캔한다.
    sync_engine은 supports_delta = False를 보고 삭제 감지 로직을 별도로 실행한다.
    """

    supports_delta: bool = False  # SyncEngine이 삭제 감지 방식 선택에 사용

    def __init__(self, calendar_name: str = "", horizon_days: int = 365):
        """
        calendar_name: 빈 문자열이면 Outlook 기본 캘린더 사용.
                       폴더 이름 지정 시 해당 서브 폴더 사용.
        """
        pythoncom.CoInitialize()
        self._app = win32com.client.Dispatch("Outlook.Application")
        self._ns = self._app.GetNamespace("MAPI")
        self._horizon_days = horizon_days
        self._calendar = self._get_calendar_folder(calendar_name)
        logger.info(
            "Outlook COM 연결 완료 (캘린더: %s)",
            calendar_name or "기본 캘린더",
        )

    def _get_calendar_folder(self, name: str):
        default = self._ns.GetDefaultFolder(_OL_FOLDER_CALENDAR)
        if not name:
            return default
        try:
            return default.Folders[name]
        except Exception:
            logger.warning("캘린더 폴더 '%s' 없음 → 기본 캘린더 사용", name)
            return default

    # ── 이벤트 조회 (항상 전체 스캔) ─────────────────────────────────────────

    def list_events(self) -> list[dict]:
        """동기화 범위 내 모든 이벤트 반환 (오늘 ~ horizon_days 후).

        Restrict 필터는 로케일에 따라 동작이 달라지므로 Python에서 직접 필터링한다.
        Start 기준 정렬 후 범위를 벗어나면 break하여 성능을 확보한다.
        """
        window_start, window_end = sync_bounds(self._horizon_days)

        items = self._calendar.Items
        # IncludeRecurrences는 Sort 이전에 설정해야 반복 일정이 전개됨
        items.IncludeRecurrences = True
        items.Sort("[Start]")

        # Restrict 필터로 날짜 범위 1차 축소 (오늘 00:00부터 — '지금'이 아님)
        start_str = window_start.strftime("%m/%d/%Y %H:%M")
        end_str = window_end.strftime("%m/%d/%Y %H:%M")
        filter_str = f"[Start] >= '{start_str}' AND [Start] < '{end_str}'"
        try:
            restricted = items.Restrict(filter_str)
        except Exception:
            restricted = items  # Restrict 실패 시 전체 컬렉션으로 폴백

        events: list[dict] = []
        for item in restricted:
            try:
                if item.Class != _OL_APPOINTMENT_CLASS:
                    continue
                # pywintypes.datetime은 naive datetime이므로 직접 비교 가능
                s = item.Start
                item_start = datetime(s.year, s.month, s.day, s.hour, s.minute, s.second)
                if item_start < window_start:
                    continue
                if item_start >= window_end:
                    break
                events.append(self._item_to_dict(item))
            except Exception as e:
                logger.warning("Outlook 이벤트 변환 실패: %s", e)
        return events

    def list_events_initial(self) -> tuple[list[dict], str]:
        """최초 동기화용. (이벤트 목록, 더미 delta_link) 반환."""
        return self.list_events(), "outlook_com_v1"

    def list_events_delta(self, delta_link: str) -> tuple[list[dict], str]:
        """증분 동기화용. COM은 delta가 없으므로 항상 전체 스캔."""
        return self.list_events(), delta_link  # delta_link 값 유지 (COM 마커)

    def _item_to_dict(self, item) -> dict:
        """Outlook AppointmentItem → 표준 이벤트 dict 변환."""
        is_all_day = bool(item.AllDayEvent)

        if is_all_day:
            # 전일 일정: 로컬 날짜 그대로 사용 (UTC 변환 시 날짜 어긋남 방지)
            s = item.Start
            e = item.End
            date_str = f"{s.year:04d}-{s.month:02d}-{s.day:02d}"
            end_date = f"{e.year:04d}-{e.month:02d}-{e.day:02d}"
            start_field = {"dateTime": f"{date_str}T00:00:00", "timeZone": "UTC"}
            end_field = {"dateTime": f"{end_date}T00:00:00", "timeZone": "UTC"}
        else:
            start_field = {"dateTime": _pywint_to_utc(item.Start), "timeZone": "UTC"}
            end_field = {"dateTime": _pywint_to_utc(item.End), "timeZone": "UTC"}

        categories = [
            c.strip()
            for c in (item.Categories or "").split(",")
            if c.strip()
        ]

        return {
            "id": item.EntryID,
            "subject": item.Subject or "",
            "isAllDay": is_all_day,
            "start": start_field,
            "end": end_field,
            "body": {"contentType": "text", "content": item.Body or ""},
            "location": {"displayName": item.Location or ""},
            "lastModifiedDateTime": _pywint_to_utc(item.LastModificationTime),
            "categories": categories,
        }

    # ── CRUD ──────────────────────────────────────────────────────────────────

    def create_event(self, body: dict) -> dict:
        """새 Outlook 이벤트 생성 후 dict 반환."""
        appt = self._app.CreateItem(_OL_APPOINTMENT_ITEM)
        self._apply_body(appt, body)
        appt.Save()
        logger.debug("Outlook 이벤트 생성: %s", body.get("subject", ""))
        return self._item_to_dict(appt)

    def update_event(self, entry_id: str, body: dict) -> dict:
        """기존 Outlook 이벤트 수정 후 dict 반환."""
        appt = self._ns.GetItemFromID(entry_id)
        self._apply_body(appt, body)
        appt.Save()
        logger.debug("Outlook 이벤트 수정: %s", body.get("subject", ""))
        return self._item_to_dict(appt)

    def get_event(self, entry_id: str) -> dict | None:
        """EntryID로 일정 조회 (동기화 윈도우 무관). 없으면 None."""
        try:
            appt = self._ns.GetItemFromID(entry_id)
            if appt.Class != _OL_APPOINTMENT_CLASS:
                logger.warning(
                    "Outlook EntryID %s 가 일정 타입이 아님 (class=%s)",
                    entry_id,
                    appt.Class,
                )
                return None
            return self._item_to_dict(appt)
        except Exception as e:
            err = str(e).lower()
            if any(
                token in err
                for token in ("not found", "could not find", "8004010f", "0x8004010f")
            ):
                return None
            logger.warning("Outlook 이벤트 조회 실패 (id=%s): %s", entry_id, e)
            raise

    def check_event_status(self, entry_id: str) -> str:
        """Outlook 일정 존재 여부: exists | missing | unknown."""
        try:
            appt = self._ns.GetItemFromID(entry_id)
            if appt.Class == _OL_APPOINTMENT_CLASS:
                return "exists"
            logger.warning(
                "Outlook EntryID %s 가 일정 타입이 아님 (class=%s)", entry_id, appt.Class
            )
            return "unknown"
        except Exception as e:
            err = str(e).lower()
            # MAPI_E_NOT_FOUND 등 — 실제 삭제로 간주
            if any(
                token in err
                for token in ("not found", "could not find", "8004010f", "0x8004010f")
            ):
                return "missing"
            logger.warning(
                "Outlook 존재 확인 실패 (id=%s) — 삭제 전파 보류: %s", entry_id, e
            )
            return "unknown"

    def event_exists(self, entry_id: str) -> bool:
        """하위 호환. exists일 때만 True."""
        return self.check_event_status(entry_id) == "exists"

    def delete_event(self, entry_id: str) -> None:
        """Outlook 이벤트 삭제 (이미 삭제된 경우 무시)."""
        try:
            appt = self._ns.GetItemFromID(entry_id)
            appt.Delete()
            logger.debug("Outlook 이벤트 삭제: %s", entry_id)
        except Exception as e:
            logger.debug("Outlook 이벤트 이미 삭제됨 (id=%s): %s", entry_id, e)

    def _apply_body(self, appt, body: dict) -> None:
        """body dict 내용을 AppointmentItem에 적용."""
        appt.Subject = body.get("subject", "")

        is_all_day = body.get("isAllDay", False)
        start = body.get("start", {})
        end = body.get("end", {})

        if is_all_day:
            appt.AllDayEvent = True
            # 전일 일정: 날짜 문자열만 설정 (시각 변환 불필요)
            appt.Start = start.get("dateTime", "")[:10]
            appt.End = end.get("dateTime", "")[:10]
        else:
            appt.AllDayEvent = False
            start_tz = start.get("timeZone", "UTC")
            end_tz = end.get("timeZone", "UTC")
            appt.Start = _to_com_local(start.get("dateTime", ""), start_tz)
            appt.End = _to_com_local(end.get("dateTime", ""), end_tz)

        content = body.get("body", {}).get("content", "")
        appt.Body = content or ""

        location = body.get("location", {}).get("displayName", "")
        appt.Location = location or ""

        # GSync-* 이외 기존 카테고리 보존 + 신규 GSync-* 적용
        existing_raw = appt.Categories or ""
        existing_cats = [c.strip() for c in existing_raw.split(",") if c.strip()]
        non_gsync = [c for c in existing_cats if not c.startswith("GSync-")]
        new_gsync = body.get("categories", [])
        all_cats = non_gsync + new_gsync
        appt.Categories = ", ".join(all_cats)

    # ── 카테고리 관리 ─────────────────────────────────────────────────────────

    def ensure_gsync_categories(self) -> None:
        """COM 방식에서는 카테고리가 이벤트 저장 시 자동 생성된다.

        색상 지정은 Outlook에서 수동으로 설정해야 한다:
        홈 탭 → 범주 → 모든 범주 → GSync-* 항목에 색상 지정.
        """
        logger.info(
            "Outlook COM 방식: GSync-* 카테고리는 동기화 시 자동 생성됩니다. "
            "색상은 Outlook '모든 범주'에서 수동 지정하세요."
        )
