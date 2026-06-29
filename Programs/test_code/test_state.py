"""
sync/state.py 단위 테스트.

파일시스템 작업은 pytest tmp_path 픽스처로 격리.
"""

import json
import os
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent))

from sync.state import (
    EventMap,
    LastSync,
    SyncStats,
    acquire_lock,
    release_lock,
)


# ── 잠금 파일 ─────────────────────────────────────────────────────────────

class TestLock:
    def test_acquire_when_no_lock_file(self, tmp_path):
        lock = tmp_path / "sync.lock"
        assert acquire_lock(lock) is True
        assert lock.exists()

    def test_refuse_when_valid_lock_exists(self, tmp_path):
        lock = tmp_path / "sync.lock"
        acquire_lock(lock)
        assert acquire_lock(lock) is False  # 이미 존재

    def test_acquire_when_stale_lock(self, tmp_path):
        lock = tmp_path / "sync.lock"
        # 3시간 전 시작 시각으로 lock 생성
        old_time = datetime(2020, 1, 1, 0, 0, 0, tzinfo=timezone.utc)
        lock.write_text(
            json.dumps({"pid": 99999, "started_at": old_time.isoformat()}),
            encoding="utf-8",
        )
        assert acquire_lock(lock) is True

    def test_acquire_overwrites_corrupt_lock(self, tmp_path):
        lock = tmp_path / "sync.lock"
        lock.write_text("not valid json", encoding="utf-8")
        assert acquire_lock(lock) is True

    def test_release_removes_file(self, tmp_path):
        lock = tmp_path / "sync.lock"
        acquire_lock(lock)
        release_lock(lock)
        assert not lock.exists()

    def test_release_nonexistent_does_not_raise(self, tmp_path):
        lock = tmp_path / "nonexistent.lock"
        release_lock(lock)  # should not raise


# ── LastSync ──────────────────────────────────────────────────────────────

class TestLastSync:
    def test_load_empty_when_file_missing(self, tmp_path):
        path = tmp_path / "last_sync.json"
        ls = LastSync.load(path)
        assert ls.is_initial()
        assert ls.last_sync_at == ""

    def test_save_and_reload(self, tmp_path):
        path = tmp_path / "last_sync.json"
        ls = LastSync(
            last_sync_at="2025-06-29T00:00:00+00:00",
            google_updated_min="2025-06-29T00:00:00+00:00",
            ms_delta_link="https://graph.microsoft.com/delta?token=xxx",
        )
        ls.save(path)

        loaded = LastSync.load(path)
        assert loaded.last_sync_at == "2025-06-29T00:00:00+00:00"
        assert loaded.ms_delta_link == "https://graph.microsoft.com/delta?token=xxx"
        assert not loaded.is_initial()

    def test_load_corrupt_file_returns_empty(self, tmp_path):
        path = tmp_path / "last_sync.json"
        path.write_text("not json", encoding="utf-8")
        ls = LastSync.load(path)
        assert ls.is_initial()


# ── EventMap ──────────────────────────────────────────────────────────────

class TestEventMap:
    def test_load_empty_when_missing(self, tmp_path):
        path = tmp_path / "event_map.json"
        em = EventMap.load(path)
        assert len(em) == 0

    def test_set_and_get(self, tmp_path):
        path = tmp_path / "event_map.json"
        em = EventMap(path=path)
        em.set("g1", "o1", "2025-06-29T00:00:00Z", "2025-06-29T00:00:00Z")
        entry = em.get("g1")
        assert entry["outlook_id"] == "o1"
        assert entry["google_modified"] == "2025-06-29T00:00:00Z"

    def test_find_by_outlook_id(self, tmp_path):
        em = EventMap(path=tmp_path / "event_map.json")
        em.set("g1", "o1", "t", "t")
        em.set("g2", "o2", "t", "t")
        assert em.find_by_outlook_id("o2") == "g2"
        assert em.find_by_outlook_id("nonexistent") is None

    def test_remove(self, tmp_path):
        em = EventMap(path=tmp_path / "event_map.json")
        em.set("g1", "o1", "t", "t")
        em.remove("g1")
        assert em.get("g1") is None

    def test_save_creates_backup(self, tmp_path):
        path = tmp_path / "event_map.json"
        bak = Path(str(path) + ".bak")

        em = EventMap(path=path)
        em.set("g1", "o1", "t", "t")
        em.save()

        # 두 번째 저장 시 .bak 생성
        em.set("g2", "o2", "t", "t")
        em.save()

        assert bak.exists()

    def test_load_from_backup_when_corrupt(self, tmp_path):
        path = tmp_path / "event_map.json"
        bak = Path(str(path) + ".bak")

        # 정상 데이터를 .bak으로 저장
        bak.write_text(
            json.dumps({"g1": {"outlook_id": "o1", "google_modified": "t", "outlook_modified": "t"}}),
            encoding="utf-8",
        )
        # 메인 파일 손상
        path.write_text("corrupt data", encoding="utf-8")

        em = EventMap.load(path)
        assert em.get("g1") is not None

    def test_clear(self, tmp_path):
        em = EventMap(path=tmp_path / "event_map.json")
        em.set("g1", "o1", "t", "t")
        em.clear()
        assert len(em) == 0

    def test_contains(self, tmp_path):
        em = EventMap(path=tmp_path / "event_map.json")
        em.set("g1", "o1", "t", "t")
        assert "g1" in em
        assert "g2" not in em


# ── SyncStats ─────────────────────────────────────────────────────────────

class TestSyncStats:
    def test_summary_format(self):
        stats = SyncStats(added=3, updated=1, deleted=2, errors=0)
        summary = stats.summary()
        assert "3" in summary
        assert "추가" in summary
        assert "수정" in summary
        assert "삭제" in summary
