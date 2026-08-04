"""수정 시각(ISO) 정규화 및 비교."""

from __future__ import annotations

import re
from datetime import datetime, timezone


def parse_modified(ts: str | None) -> datetime | None:
    """수정 시각 문자열을 UTC aware datetime으로 파싱. 실패 시 None."""
    if not ts:
        return None
    clean = re.sub(r"\.\d+", "", ts.strip())
    clean = clean.replace("Z", "+00:00")
    try:
        dt = datetime.fromisoformat(clean)
    except ValueError:
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc)


def modified_equal(a: str | None, b: str | None) -> bool:
    """동일 시각이면 True (소수초·Z 표기 차이 무시)."""
    da, db = parse_modified(a), parse_modified(b)
    if da is None or db is None:
        return (a or "") == (b or "")
    return da == db


def modified_gte(a: str | None, b: str | None) -> bool:
    """a >= b (시각). 파싱 실패 시 문자열 비교로 폴백."""
    da, db = parse_modified(a), parse_modified(b)
    if da is None or db is None:
        return (a or "") >= (b or "")
    return da >= db
