"""sync_window.py 단위 테스트."""

import sys
from datetime import datetime, timedelta
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).parent.parent))

from calendar_api.sync_window import sync_bounds


class TestSyncBounds:
    @patch("calendar_api.sync_window.datetime")
    def test_window_starts_at_local_midnight_not_now(self, mock_dt):
        """동기화 시작은 '지금'이 아니라 오늘 00:00."""
        mock_dt.now.return_value = datetime(2025, 7, 2, 15, 30, 0)
        mock_dt.side_effect = lambda *a, **k: datetime(*a, **k)

        start, end = sync_bounds(365)

        assert start == datetime(2025, 7, 2, 0, 0, 0)
        assert end == datetime(2025, 7, 2, 0, 0, 0) + timedelta(days=366)

    @patch("calendar_api.sync_window.datetime")
    def test_all_day_event_on_today_included_all_day(self, mock_dt):
        """7/2 종일 일정은 7/2 15:30에도 범위 안에 포함."""
        mock_dt.now.return_value = datetime(2025, 7, 2, 15, 30, 0)
        mock_dt.side_effect = lambda *a, **k: datetime(*a, **k)

        start, _ = sync_bounds(365)
        event_start = datetime(2025, 7, 2, 0, 0, 0)

        assert event_start >= start
