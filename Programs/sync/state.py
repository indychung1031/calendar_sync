"""
로컬 동기화 상태 관리: last_sync.json, event_map.json, sync.lock.
"""

import json
import os
import shutil
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path

_SYNC_DIR = Path(__file__).parent

LOCK_PATH = _SYNC_DIR / "sync.lock"
LAST_SYNC_PATH = _SYNC_DIR / "last_sync.json"
EVENT_MAP_PATH = _SYNC_DIR / "event_map.json"

# 2시간 경과한 lock 파일은 비정상 종료로 간주하고 무시
_LOCK_TIMEOUT_SECONDS = 7200


# ── 잠금 파일 ──────────────────────────────────────────────────────────────

def acquire_lock(lock_path: Path = LOCK_PATH) -> bool:
    """잠금 파일 생성. 유효한 잠금이 이미 존재하면 False 반환."""
    if lock_path.exists():
        try:
            data = json.loads(lock_path.read_text(encoding="utf-8"))
            started_at = datetime.fromisoformat(data["started_at"])
            elapsed = (datetime.now(timezone.utc) - started_at).total_seconds()
            if elapsed < _LOCK_TIMEOUT_SECONDS:
                return False
        except (json.JSONDecodeError, KeyError, ValueError):
            pass  # 손상된 잠금 파일은 무시

    lock_path.write_text(
        json.dumps(
            {"pid": os.getpid(), "started_at": datetime.now(timezone.utc).isoformat()},
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    return True


def release_lock(lock_path: Path = LOCK_PATH) -> None:
    lock_path.unlink(missing_ok=True)


# ── 마지막 동기화 상태 ─────────────────────────────────────────────────────

@dataclass
class LastSync:
    last_sync_at: str = ""
    google_updated_min: str = ""   # 다음 Google 증분 조회의 updatedMin
    ms_delta_link: str = ""        # 다음 MS Graph 증분 조회 URL

    def is_initial(self) -> bool:
        return not self.last_sync_at

    def save(self, path: Path = LAST_SYNC_PATH) -> None:
        path.write_text(
            json.dumps(
                {
                    "last_sync_at": self.last_sync_at,
                    "google_updated_min": self.google_updated_min,
                    "ms_delta_link": self.ms_delta_link,
                },
                indent=2,
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )

    @classmethod
    def load(cls, path: Path = LAST_SYNC_PATH) -> "LastSync":
        if not path.exists():
            return cls()
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            return cls(
                last_sync_at=data.get("last_sync_at", ""),
                google_updated_min=data.get("google_updated_min", ""),
                ms_delta_link=data.get("ms_delta_link", ""),
            )
        except (json.JSONDecodeError, KeyError):
            return cls()


# ── 이벤트 매핑 테이블 ─────────────────────────────────────────────────────

class EventMap:
    """
    Google ↔ Outlook 이벤트 ID + 수정 시각 매핑 테이블.

    구조:
      {
        "<google_id>": {
          "outlook_id": "<outlook_id>",
          "google_modified": "<UTC ISO>",
          "outlook_modified": "<UTC ISO>"
        }
      }

    google_modified / outlook_modified에는 서버가 반환한 실제 수정 시각을 저장한다.
    API 응답값을 그대로 저장해야 무한 루프 방지 로직이 정상 동작한다.
    """

    def __init__(self, data: dict | None = None, path: Path = EVENT_MAP_PATH):
        self._data: dict = data or {}
        self._path = path
        self._bak_path = Path(str(path) + ".bak")

    @classmethod
    def load(cls, path: Path = EVENT_MAP_PATH) -> "EventMap":
        if not path.exists():
            return cls(path=path)
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            return cls(data=data, path=path)
        except json.JSONDecodeError:
            # 손상된 경우 .bak 복구 시도
            bak = Path(str(path) + ".bak")
            if bak.exists():
                try:
                    data = json.loads(bak.read_text(encoding="utf-8"))
                    return cls(data=data, path=path)
                except json.JSONDecodeError:
                    pass
            return cls(path=path)

    def save(self) -> None:
        # 덮어쓰기 전 백업
        if self._path.exists():
            shutil.copy2(self._path, self._bak_path)
        self._path.write_text(
            json.dumps(self._data, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )

    def get(self, google_id: str) -> dict | None:
        return self._data.get(google_id)

    def set(
        self,
        google_id: str,
        outlook_id: str,
        google_modified: str,
        outlook_modified: str,
    ) -> None:
        existing = self._data.get(google_id, {})
        self._data[google_id] = {
            "outlook_id": outlook_id,
            "google_modified": google_modified,
            "outlook_modified": outlook_modified,
            "outlook_missing_since": existing.get("outlook_missing_since"),
            "google_delete_pending_since": existing.get("google_delete_pending_since"),
        }

    def mark_outlook_missing(self, google_id: str) -> None:
        """Outlook 삭제 1차 감지 — 다음 동기화까지 보류."""
        entry = self._data.get(google_id)
        if entry and not entry.get("outlook_missing_since"):
            entry["outlook_missing_since"] = datetime.now(timezone.utc).isoformat()

    def clear_outlook_missing(self, google_id: str) -> None:
        entry = self._data.get(google_id)
        if entry:
            entry.pop("outlook_missing_since", None)

    def is_outlook_missing_pending(self, google_id: str) -> bool:
        return bool(self._data.get(google_id, {}).get("outlook_missing_since"))

    def mark_google_delete_pending(self, google_id: str) -> None:
        """Google 삭제 1차 감지 — 다음 동기화까지 보류."""
        entry = self._data.get(google_id)
        if entry and not entry.get("google_delete_pending_since"):
            entry["google_delete_pending_since"] = datetime.now(timezone.utc).isoformat()

    def clear_google_delete_pending(self, google_id: str) -> None:
        entry = self._data.get(google_id)
        if entry:
            entry.pop("google_delete_pending_since", None)

    def is_google_delete_pending(self, google_id: str) -> bool:
        return bool(self._data.get(google_id, {}).get("google_delete_pending_since"))

    def remove(self, google_id: str) -> None:
        self._data.pop(google_id, None)

    def find_by_outlook_id(self, outlook_id: str) -> str | None:
        """outlook_id로 google_id 역조회."""
        for gid, entry in self._data.items():
            if entry.get("outlook_id") == outlook_id:
                return gid
        return None

    def clear(self) -> None:
        self._data.clear()

    def items(self):
        return self._data.items()

    def __len__(self) -> int:
        return len(self._data)

    def __contains__(self, google_id: str) -> bool:
        return google_id in self._data


# ── 동기화 통계 ────────────────────────────────────────────────────────────

@dataclass
class SyncStats:
    added: int = 0
    updated: int = 0
    deleted: int = 0
    errors: int = 0

    def summary(self) -> str:
        return (
            f"추가 {self.added}건 / 수정 {self.updated}건 / "
            f"삭제 {self.deleted}건 / 오류 {self.errors}건"
        )
