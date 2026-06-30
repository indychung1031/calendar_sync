"""
Outlook ↔ Google 캘린더 양방향 동기화 진입점.

사용법:
  python main.py                 # 일반 동기화 (최초 또는 증분)
  python main.py --reauth        # Microsoft 재인증 후 동기화
  python main.py --rebuild-map   # event_map 재구성 (중복 발생 시 복구용)
  python main.py --list-calendars  # Google 캘린더 목록과 ID 출력
"""

import argparse
import logging
import sys
from pathlib import Path

# Programs/ 기준으로 패키지 import가 동작하도록 경로 추가
sys.path.insert(0, str(Path(__file__).parent))

from auth.google_auth import get_credentials
from auth.ms_auth import MSAuthExpiredError, get_access_token
from calendar_api.google_calendar import GoogleCalendarClient
from calendar_api.ms_calendar import MSCalendarClient
from config import Config
from sync.state import EventMap, LastSync, acquire_lock, release_lock
from sync.sync_engine import SyncEngine

_LOG_DIR = Path(__file__).parent / "logs"


def setup_logging() -> None:
    _LOG_DIR.mkdir(exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(_LOG_DIR / "sync.log", encoding="utf-8"),
        ],
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Outlook ↔ Google 캘린더 동기화")
    parser.add_argument(
        "--reauth",
        action="store_true",
        help="Microsoft 인증 토큰을 초기화하고 재인증",
    )
    parser.add_argument(
        "--rebuild-map",
        action="store_true",
        dest="rebuild_map",
        help="event_map을 재구성 (중복 이벤트 발생 후 복구용)",
    )
    parser.add_argument(
        "--list-calendars",
        action="store_true",
        dest="list_calendars",
        help="Google 캘린더 목록과 ID를 출력 (GOOGLE_CALENDAR_ID 설정 참고용)",
    )
    return parser.parse_args()


def main() -> None:
    setup_logging()
    logger = logging.getLogger(__name__)
    args = parse_args()

    # ── 동시 실행 방지 ────────────────────────────────────────────────────
    if not acquire_lock():
        logger.warning("다른 동기화 프로세스 실행 중. 종료합니다.")
        sys.exit(0)

    try:
        _run(args, logger)
    finally:
        release_lock()


def _run(args: argparse.Namespace, logger: logging.Logger) -> None:
    # ── 설정 로드 ─────────────────────────────────────────────────────────
    try:
        cfg = Config.load()
    except ValueError as e:
        logger.critical("설정 오류: %s", e)
        sys.exit(1)

    # ── Google 인증 ───────────────────────────────────────────────────────
    try:
        google_creds = get_credentials(
            cfg.google_credentials_path,
            cfg.google_token_path,
        )
    except FileNotFoundError as e:
        logger.critical("%s", e)
        sys.exit(1)
    except Exception as e:
        logger.critical("Google 인증 실패: %s", e)
        sys.exit(1)

    # ── Microsoft 인증 ────────────────────────────────────────────────────
    try:
        ms_token = get_access_token(
            cfg.ms_client_id,
            cfg.ms_tenant_id,
            cfg.ms_token_path,
            force_reauth=args.reauth,
        )
    except MSAuthExpiredError:
        # 메시지는 ms_auth.py에서 CRITICAL로 이미 기록됨
        sys.exit(1)
    except Exception as e:
        logger.critical("Microsoft 인증 실패: %s", e)
        sys.exit(1)

    # ── 클라이언트 초기화 ─────────────────────────────────────────────────
    google_client = GoogleCalendarClient(
        google_creds, cfg.google_calendar_id, cfg.sync_recurring_horizon_days
    )
    ms_client = MSCalendarClient(
        ms_token, cfg.ms_calendar_id, cfg.sync_recurring_horizon_days
    )

    # ── 상태 로드 ─────────────────────────────────────────────────────────
    event_map = EventMap.load()
    last_sync = LastSync.load()

    engine = SyncEngine(google_client, ms_client, event_map, last_sync)

    # ── 실행 ──────────────────────────────────────────────────────────────
    if args.list_calendars:
        calendars = google_client.list_calendars()
        print("\n=== Google 캘린더 목록 ===")
        for cal in calendars:
            print(f"  이름: {cal.get('summary', '')}")
            print(f"  ID:   {cal.get('id', '')}")
            print()
        print("→ .env 파일의 GOOGLE_CALENDAR_ID에 원하는 ID를 입력하세요.")
        return

    if args.rebuild_map:
        logger.info("--rebuild-map 모드: event_map 재구성")
        try:
            engine.rebuild_map()
        except Exception as e:
            logger.error("rebuild-map 실패: %s", e)
            sys.exit(1)
        return

    try:
        stats = engine.run()
        logger.info("동기화 완료: %s", stats.summary())
    except Exception as e:
        logger.error("동기화 중 예외 발생: %s", e, exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()
