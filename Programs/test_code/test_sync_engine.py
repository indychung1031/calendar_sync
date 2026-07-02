"""
sync_engine.py 단위 테스트.

Google/MS 클라이언트는 Mock으로 대체. API 호출 없음.
"""

import sys
from pathlib import Path
from unittest.mock import MagicMock, call, patch

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent))

from sync.state import EventMap, LastSync, SyncStats
from sync.sync_engine import SyncEngine


def _make_engine(tmp_path: Path) -> tuple[SyncEngine, MagicMock, MagicMock]:
    google = MagicMock()
    ms = MagicMock()
    event_map = EventMap(path=tmp_path / "event_map.json")
    last_sync = LastSync()
    engine = SyncEngine(google, ms, event_map, last_sync)
    return engine, google, ms


def _google_event(
    gid: str,
    title: str = "회의",
    utc_start: str = "2025-06-29T00:00:00Z",
    utc_end: str = "2025-06-29T01:00:00Z",
    updated: str = "2025-06-29T00:30:00Z",
    status: str = "confirmed",
) -> dict:
    return {
        "id": gid,
        "summary": title,
        "start": {"dateTime": utc_start, "timeZone": "UTC"},
        "end": {"dateTime": utc_end, "timeZone": "UTC"},
        "updated": updated,
        "status": status,
    }


def _ms_event(
    mid: str,
    title: str = "회의",
    utc_start: str = "2025-06-29T00:00:00.0000000",
    utc_end: str = "2025-06-29T01:00:00.0000000",
    last_modified: str = "2025-06-29T00:30:00Z",
    removed: bool = False,
) -> dict:
    event: dict = {
        "id": mid,
        "subject": title,
        "isAllDay": False,
        "start": {"dateTime": utc_start, "timeZone": "UTC"},
        "end": {"dateTime": utc_end, "timeZone": "UTC"},
        "lastModifiedDateTime": last_modified,
        "body": {"contentType": "text", "content": ""},
        "location": {"displayName": ""},
    }
    if removed:
        event["@removed"] = {"reason": "deleted"}
    return event


# ── 최초 동기화 ───────────────────────────────────────────────────────────

class TestInitialSync:
    def test_pair_matching_events(self, tmp_path):
        """같은 제목+시각의 이벤트는 생성 없이 매핑만 등록."""
        engine, google, ms = _make_engine(tmp_path)

        google.list_events.return_value = [_google_event("g1")]
        ms.list_events_initial.return_value = (
            [_ms_event("o1")],
            "https://delta.link",
        )

        engine.run()

        # 짝이 맞으면 create 호출 없음
        google.create_event.assert_not_called()
        ms.create_event.assert_not_called()

        entry = engine._map.get("g1")
        assert entry is not None
        assert entry["outlook_id"] == "o1"

    def test_unmatched_google_creates_in_outlook(self, tmp_path):
        """Outlook에 없는 Google 이벤트는 Outlook에 생성."""
        engine, google, ms = _make_engine(tmp_path)

        google.list_events.return_value = [_google_event("g1", title="신규")]
        ms.list_events_initial.return_value = ([], "https://delta.link")
        ms.create_event.return_value = _ms_event("o_new", last_modified="2025-06-29T00:35:00Z")

        engine.run()

        ms.create_event.assert_called_once()
        assert engine.stats.added == 1

    def test_unmatched_ms_creates_in_google(self, tmp_path):
        """Google에 없는 MS 이벤트는 Google에 생성."""
        engine, google, ms = _make_engine(tmp_path)

        google.list_events.return_value = []
        ms.list_events_initial.return_value = (
            [_ms_event("o1", title="신규")],
            "https://delta.link",
        )
        google.create_event.return_value = _google_event("g_new", updated="2025-06-29T00:35:00Z")

        engine.run()

        google.create_event.assert_called_once()
        assert engine.stats.added == 1

    def test_deleted_events_skipped(self, tmp_path):
        """삭제된 이벤트는 최초 동기화에서 무시."""
        engine, google, ms = _make_engine(tmp_path)

        google.list_events.return_value = [
            _google_event("g1", status="cancelled")
        ]
        ms.list_events_initial.return_value = (
            [_ms_event("o1", removed=True)],
            "https://delta.link",
        )

        engine.run()

        google.create_event.assert_not_called()
        ms.create_event.assert_not_called()

    def test_delta_link_saved(self, tmp_path):
        """최초 동기화 후 deltaLink 저장."""
        engine, google, ms = _make_engine(tmp_path)
        google.list_events.return_value = []
        ms.list_events_initial.return_value = ([], "https://my.delta.link")

        engine.run()

        assert engine._last_sync.ms_delta_link == "https://my.delta.link"


# ── 증분 동기화 ───────────────────────────────────────────────────────────

class TestIncrementalSync:
    def _setup_existing_mapping(self, engine: SyncEngine) -> None:
        engine._map.set("g1", "o1", "2025-06-29T00:30:00Z", "2025-06-29T00:30:00Z")
        engine._last_sync.last_sync_at = "2025-06-29T00:30:00Z"
        engine._last_sync.google_updated_min = "2025-06-29T00:30:00Z"
        engine._last_sync.ms_delta_link = "https://delta.link"

    def test_skip_unchanged_google_event(self, tmp_path):
        """수정 시각이 동일하면 무한 루프 방지로 건너뜀."""
        engine, google, ms = _make_engine(tmp_path)
        self._setup_existing_mapping(engine)

        # updated가 저장된 값과 동일 → 우리가 동기화한 결과물
        google.list_events.return_value = [
            _google_event("g1", updated="2025-06-29T00:30:00Z")
        ]
        ms.list_events_delta.return_value = ([], "https://delta.link2")

        engine.run()
        ms.update_event.assert_not_called()

    def test_apply_google_change_to_outlook(self, tmp_path):
        """Google 이벤트 변경 시 Outlook 업데이트."""
        engine, google, ms = _make_engine(tmp_path)
        self._setup_existing_mapping(engine)

        # updated가 저장값과 다름 → 실제 변경
        google.list_events.return_value = [
            _google_event("g1", updated="2025-06-29T01:00:00Z")
        ]
        ms.list_events_delta.return_value = ([], "https://delta.link2")
        ms.update_event.return_value = _ms_event("o1", last_modified="2025-06-29T01:01:00Z")

        engine.run()
        ms.update_event.assert_called_once()
        assert engine.stats.updated == 1

    def test_apply_ms_change_to_google(self, tmp_path):
        """MS 이벤트 변경 시 Google 업데이트."""
        engine, google, ms = _make_engine(tmp_path)
        self._setup_existing_mapping(engine)

        google.list_events.return_value = []
        ms.list_events_delta.return_value = (
            [_ms_event("o1", last_modified="2025-06-29T02:00:00Z")],
            "https://delta.link2",
        )
        google.update_event.return_value = _google_event("g1", updated="2025-06-29T02:01:00Z")

        engine.run()
        google.update_event.assert_called_once()
        assert engine.stats.updated == 1

    def test_conflict_google_wins(self, tmp_path):
        """양쪽 변경 시 더 최신(Google) 쪽 승리."""
        engine, google, ms = _make_engine(tmp_path)
        self._setup_existing_mapping(engine)

        # Google이 더 최신 수정 시각
        google.list_events.return_value = [
            _google_event("g1", updated="2025-06-29T03:00:00Z")
        ]
        ms.list_events_delta.return_value = (
            [_ms_event("o1", last_modified="2025-06-29T02:00:00Z")],
            "https://delta.link2",
        )
        ms.update_event.return_value = _ms_event("o1", last_modified="2025-06-29T03:05:00Z")

        engine.run()

        # Outlook만 업데이트, Google은 업데이트 안 함
        ms.update_event.assert_called_once()
        google.update_event.assert_not_called()

    def test_conflict_ms_wins(self, tmp_path):
        """양쪽 변경 시 더 최신(MS) 쪽 승리."""
        engine, google, ms = _make_engine(tmp_path)
        self._setup_existing_mapping(engine)

        # MS가 더 최신 수정 시각
        google.list_events.return_value = [
            _google_event("g1", updated="2025-06-29T02:00:00Z")
        ]
        ms.list_events_delta.return_value = (
            [_ms_event("o1", last_modified="2025-06-29T03:00:00Z")],
            "https://delta.link2",
        )
        google.update_event.return_value = _google_event("g1", updated="2025-06-29T03:05:00Z")

        engine.run()

        # Google만 업데이트, Outlook은 업데이트 안 함
        google.update_event.assert_called_once()
        ms.update_event.assert_not_called()

    def test_delete_from_outlook_when_google_deleted(self, tmp_path):
        """Google cancelled 2회 확인 후 Outlook 삭제."""
        engine, google, ms = _make_engine(tmp_path)
        self._setup_existing_mapping(engine)
        google.check_event_status.return_value = "cancelled"

        google.list_events.return_value = [
            _google_event("g1", status="cancelled")
        ]
        ms.list_events_delta.return_value = ([], "https://delta.link2")

        engine.run()
        ms.delete_event.assert_not_called()
        assert engine._map.is_google_delete_pending("g1")

        google.list_events.return_value = []
        engine.run()

        ms.delete_event.assert_called_once_with("o1")
        assert engine._map.get("g1") is None
        assert engine.stats.deleted == 1

    def test_delete_from_google_when_ms_deleted(self, tmp_path):
        """MS @removed 삭제 시 2회 확인 후 Google도 삭제."""
        engine, google, ms = _make_engine(tmp_path)
        self._setup_existing_mapping(engine)
        ms.supports_delta = False
        ms.check_event_status.return_value = "missing"

        google.list_events.return_value = []
        ms.list_events_delta.return_value = (
            [_ms_event("o1", removed=True)],
            "https://delta.link2",
        )

        engine.run()
        google.delete_event.assert_not_called()
        assert engine._map.is_outlook_missing_pending("g1")

        ms.list_events_delta.return_value = ([], "https://delta.link2")
        engine.run()

        google.delete_event.assert_called_once_with("g1")
        assert engine._map.get("g1") is None
        assert engine.stats.deleted == 1

    def test_new_google_event_creates_in_outlook(self, tmp_path):
        """event_map에 없는 Google 이벤트 → Outlook 신규 생성."""
        engine, google, ms = _make_engine(tmp_path)
        # last_sync 설정 (증분 모드)
        engine._last_sync.last_sync_at = "2025-06-29T00:00:00Z"
        engine._last_sync.google_updated_min = "2025-06-29T00:00:00Z"
        engine._last_sync.ms_delta_link = "https://delta.link"

        google.list_events.return_value = [_google_event("g_new")]
        ms.list_events_delta.return_value = ([], "https://delta.link2")
        ms.create_event.return_value = _ms_event("o_new", last_modified="2025-06-29T00:35:00Z")

        engine.run()
        ms.create_event.assert_called_once()
        assert engine.stats.added == 1

    def test_error_counted_but_continues(self, tmp_path):
        """API 오류 발생 시 오류 카운트 증가 후 계속 진행."""
        engine, google, ms = _make_engine(tmp_path)
        engine._last_sync.last_sync_at = "2025-06-29T00:00:00Z"
        engine._last_sync.google_updated_min = "2025-06-29T00:00:00Z"
        engine._last_sync.ms_delta_link = "https://delta.link"

        google.list_events.return_value = [_google_event("g_new")]
        ms.list_events_delta.return_value = ([], "https://delta.link2")
        ms.create_event.side_effect = Exception("API 오류")

        engine.run()
        assert engine.stats.errors == 1
        assert engine.stats.added == 0


# ── COM 전체 스캔 삭제 감지 ───────────────────────────────────────────────

class TestOutlookComDeletion:
    """Outlook COM 방식(전체 스캔)에서 삭제 이벤트 감지 테스트."""

    def _setup_com_engine(self, tmp_path: Path):
        engine, google, ms = _make_engine(tmp_path)
        # COM 방식 표시 (delta 미지원)
        ms.supports_delta = False
        ms.list_events_delta.return_value = ([], "outlook_com_v1")
        # 기존 매핑 등록
        engine._map.set("g1", "o1", "2025-06-29T00:30:00Z", "2025-06-29T00:30:00Z")
        engine._last_sync.last_sync_at = "2025-06-29T00:30:00Z"
        engine._last_sync.google_updated_min = "2025-06-29T00:30:00Z"
        engine._last_sync.ms_delta_link = "outlook_com_v1"
        return engine, google, ms

    def test_detect_outlook_deletion_and_remove_from_google(self, tmp_path):
        """Outlook 이벤트가 2회 연속 missing이면 Google에서도 삭제."""
        engine, google, ms = self._setup_com_engine(tmp_path)

        ms.list_events_delta.return_value = ([], "outlook_com_v1")
        ms.check_event_status.return_value = "missing"
        google.list_events.return_value = []

        # 1차: 보류만
        engine.run()
        google.delete_event.assert_not_called()
        assert engine._map.get("g1") is not None
        assert engine._map.is_outlook_missing_pending("g1")

        # 2차: 삭제
        engine.run()
        ms.check_event_status.assert_called_with("o1")
        google.delete_event.assert_called_once_with("g1")
        assert engine._map.get("g1") is None
        assert engine.stats.deleted == 1

    def test_no_false_deletion_when_event_exists(self, tmp_path):
        """Outlook에 이벤트가 있으면 삭제하지 않음."""
        engine, google, ms = self._setup_com_engine(tmp_path)

        # o1이 현재 목록에 존재 → 삭제 없음
        ms.list_events_delta.return_value = (
            [_ms_event("o1", last_modified="2025-06-29T00:30:00Z")],
            "outlook_com_v1",
        )
        google.list_events.return_value = []

        engine.run()

        google.delete_event.assert_not_called()
        assert engine._map.get("g1") is not None

    def test_no_false_deletion_when_event_outside_scan_but_still_in_outlook(self, tmp_path):
        """스캔 목록에 없어도 Outlook에 존재하면 Google에서 삭제하지 않음."""
        engine, google, ms = self._setup_com_engine(tmp_path)

        ms.list_events_delta.return_value = ([], "outlook_com_v1")
        ms.check_event_status.return_value = "exists"
        google.list_events.return_value = []

        engine.run()

        ms.check_event_status.assert_called_once_with("o1")
        google.delete_event.assert_not_called()
        assert engine._map.get("g1") is not None

    def test_no_deletion_when_outlook_status_unknown(self, tmp_path):
        """COM 오류(unknown) 시 Google에서 삭제하지 않음."""
        engine, google, ms = self._setup_com_engine(tmp_path)

        ms.list_events_delta.return_value = ([], "outlook_com_v1")
        ms.check_event_status.return_value = "unknown"
        google.list_events.return_value = []

        engine.run()
        engine.run()

        google.delete_event.assert_not_called()
        assert engine._map.get("g1") is not None

    def test_google_delete_requires_two_confirmations(self, tmp_path):
        """Google cancelled도 API 재확인 + 2회 연속 후 Outlook 삭제."""
        engine, google, ms = _make_engine(tmp_path)
        ms.supports_delta = True
        engine._map.set("g1", "o1", "2025-06-29T00:30:00Z", "2025-06-29T00:30:00Z")
        engine._last_sync.last_sync_at = "2025-06-29T00:30:00Z"
        engine._last_sync.google_updated_min = "2025-06-29T00:30:00Z"

        cancelled = _google_event("g1", status="cancelled")
        google.list_events.return_value = [cancelled]
        google.check_event_status.return_value = "cancelled"
        ms.list_events_delta.return_value = ([], "link")

        engine.run()
        ms.delete_event.assert_not_called()
        assert engine._map.is_google_delete_pending("g1")

        google.list_events.return_value = []
        engine.run()
        ms.delete_event.assert_called_once_with("o1")
        assert engine._map.get("g1") is None
