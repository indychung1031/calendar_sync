import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv

# Programs/ 기준 프로젝트 루트의 .env 로드
_PROJECT_ROOT = Path(__file__).parent.parent
load_dotenv(_PROJECT_ROOT / ".env")


def _req(key: str) -> str:
    val = os.getenv(key, "").strip()
    if not val:
        raise ValueError(f"필수 환경 변수 누락: {key} (.env 파일 확인)")
    return val


@dataclass
class Config:
    google_credentials_path: Path
    google_token_path: Path
    google_calendar_id: str

    ms_calendar_id: str  # 빈 문자열 = Outlook 기본 캘린더

    sync_recurring_horizon_days: int

    @classmethod
    def load(cls) -> "Config":
        return cls(
            google_credentials_path=_PROJECT_ROOT / _req("GOOGLE_CREDENTIALS_PATH"),
            google_token_path=_PROJECT_ROOT / _req("GOOGLE_TOKEN_PATH"),
            google_calendar_id=os.getenv("GOOGLE_CALENDAR_ID", "primary"),
            ms_calendar_id=os.getenv("MS_CALENDAR_ID", "").strip(),
            sync_recurring_horizon_days=int(
                os.getenv("SYNC_RECURRING_HORIZON_DAYS", "365")
            ),
        )
