"""동기화 날짜 범위 계산 (Google·Outlook 공통)."""

from datetime import datetime, timedelta


def sync_bounds(horizon_days: int) -> tuple[datetime, datetime]:
    """로컬 오늘 00:00(포함) ~ horizon_days일 후 자정(미포함) 범위.

    '오늘 이후'는 '지금 이후'가 아니라 '오늘 날짜부터'를 의미한다.
    당일 일정이 자정을 지나도 동기화·삭제 오인되지 않도록 한다.
    """
    local_now = datetime.now()
    window_start = local_now.replace(hour=0, minute=0, second=0, microsecond=0)
    window_end = window_start + timedelta(days=horizon_days + 1)
    return window_start, window_end
