"""Google Calendar OAuth 2.0 인증 (Installed App flow)."""

import logging
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

_SCOPES = ["https://www.googleapis.com/auth/calendar"]
logger = logging.getLogger(__name__)


def get_credentials(
    credentials_path: Path,
    token_path: Path,
    force_reauth: bool = False,
) -> Credentials:
    """
    저장된 토큰으로 자격증명 반환. 없거나 만료 시 브라우저 인증 실행.

    force_reauth=True 이면 기존 토큰을 무시하고 재인증한다.
    """
    creds: Credentials | None = None

    if not force_reauth and token_path.exists():
        try:
            creds = Credentials.from_authorized_user_file(str(token_path), _SCOPES)
        except Exception as e:
            logger.warning("Google 토큰 파일 읽기 실패, 재인증 진행: %s", e)

    if creds and creds.valid:
        return creds

    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
            _save_token(creds, token_path)
            logger.info("Google 액세스 토큰 갱신 완료")
            return creds
        except Exception as e:
            logger.warning("Google 토큰 갱신 실패, 재인증 진행: %s", e)

    # 브라우저 인증 (최초 실행 또는 갱신 불가)
    if not credentials_path.exists():
        raise FileNotFoundError(
            f"Google credentials.json 파일 없음: {credentials_path}\n"
            "Google Cloud Console에서 OAuth 2.0 클라이언트 ID를 다운로드하세요."
        )

    flow = InstalledAppFlow.from_client_secrets_file(str(credentials_path), _SCOPES)
    creds = flow.run_local_server(port=0)
    _save_token(creds, token_path)
    logger.info("Google 인증 완료, 토큰 저장: %s", token_path)
    return creds


def _save_token(creds: Credentials, token_path: Path) -> None:
    token_path.write_text(creds.to_json(), encoding="utf-8")
