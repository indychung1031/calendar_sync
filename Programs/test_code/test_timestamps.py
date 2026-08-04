"""timestamps.py 단위 테스트."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from sync.timestamps import modified_equal, modified_gte, parse_modified


class TestParseModified:
    def test_zulu_and_offset(self):
        a = parse_modified("2025-06-29T00:30:00Z")
        b = parse_modified("2025-06-29T00:30:00+00:00")
        assert a == b

    def test_fractional_seconds_stripped(self):
        a = parse_modified("2025-06-29T00:30:00.123456Z")
        b = parse_modified("2025-06-29T00:30:00Z")
        assert a == b


class TestCompare:
    def test_equal_ignores_fraction(self):
        assert modified_equal(
            "2025-06-29T00:30:00.123456Z", "2025-06-29T00:30:00Z"
        )

    def test_gte_fraction_vs_zulu(self):
        # 문자열 비교라면 '.' < 'Z' 라 False지만, datetime 비교면 True
        assert modified_gte(
            "2025-06-29T00:30:00.123456Z", "2025-06-29T00:30:00Z"
        )

    def test_gte_later(self):
        assert modified_gte(
            "2025-06-29T01:00:00Z", "2025-06-29T00:30:00Z"
        )
        assert not modified_gte(
            "2025-06-29T00:30:00Z", "2025-06-29T01:00:00Z"
        )
