"""Microsoft MSAL 인증 (Public client + device code flow)."""

import logging
from pathlib import Path

import msal

logger = logging.getLogger(__name__)

_SCOPES = ["Calendars.ReadWrite"]
_GRAPH_BASE = "https://graph.microsoft.com"


class MSAuthExpiredError(Exception):
    """MSAL refresh token 만료 (90일 미사용)."""


def get_access_token(
    client_id: str,
    tenant_id: str,
    token_path: Path,
    force_reauth: bool = False,
) -> str:
    """
    유효한 MS Graph access token 반환.

    token_path에 캐시된 토큰이 있으면 자동 갱신 시도.
    없거나 refresh token 만료 시 device code flow로 재인증.
    force_reauth=True 이면 캐시를 무시하고 재인증한다.
    """
    cache = _load_cache(token_path, force_reauth)
    app = msal.PublicClientApplication(
        client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        token_cache=cache,
    )

    # 캐시된 계정으로 silent 갱신 시도
    if not force_reauth:
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(_SCOPES, account=accounts[0])
            if result and "access_token" in result:
                _persist_cache(cache, token_path)
                return result["access_token"]

            # silent 실패 원인 분석
            if result:
                error = result.get("error", "")
                if error in ("invalid_grant", "interaction_required"):
                    # Refresh token 만료 (90일 미사용)
                    logger.critical(
                        "Microsoft refresh token이 만료되었습니다 (90일 미사용). "
                        "재인증이 필요합니다: python main.py --reauth"
                    )
                    raise MSAuthExpiredError(
                        "Microsoft refresh token 만료. --reauth 플래그로 재인증하세요."
                    )

    # Device code flow (최초 인증 또는 재인증)
    flow = app.initiate_device_flow(scopes=_SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(f"Device code flow 시작 실패: {flow}")

    print("\n" + flow["message"] + "\n")  # 사용자에게 인증 URL/코드 안내
    result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        raise RuntimeError(
            f"MS 인증 실패: {result.get('error_description', result.get('error'))}"
        )

    _persist_cache(cache, token_path)
    logger.info("Microsoft 인증 완료, 토큰 저장: %s", token_path)
    return result["access_token"]


def _load_cache(token_path: Path, force_reauth: bool) -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    if not force_reauth and token_path.exists():
        try:
            cache.deserialize(token_path.read_text(encoding="utf-8"))
        except Exception as e:
            logger.warning("MS 토큰 캐시 읽기 실패, 재인증 진행: %s", e)
    return cache


def _persist_cache(cache: msal.SerializableTokenCache, token_path: Path) -> None:
    if cache.has_state_changed:
        token_path.write_text(cache.serialize(), encoding="utf-8")
