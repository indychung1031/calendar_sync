"""
양방향 동기화 핵심 엔진.

최초 동기화: 제목 + UTC 시작 시각으로 짝짓기 후 나머지 복사.
증분 동기화: 변경 이벤트만 처리. 무한 루프는 서버 반환 수정 시각 비교로 방지.
충돌: 양쪽 모두 변경 시 더 최신 수정 시각 쪽이 승리.
"""

import logging
from datetime import datetime, timezone

from calendar_api.event_mapper import (
    google_event_to_outlook,
    is_google_deleted,
    is_outlook_deleted,
    match_key_for_google,
    match_key_for_outlook,
    outlook_event_to_google,
)
from sync.state import EventMap, LastSync, SyncStats

logger = logging.getLogger(__name__)


class SyncEngine:
    def __init__(
        self,
        google_client,
        ms_client,
        event_map: EventMap,
        last_sync: LastSync,
    ):
        self._google = google_client
        self._ms = ms_client
        self._map = event_map
        self._last_sync = last_sync
        self.stats = SyncStats()

    # ── 진입점 ────────────────────────────────────────────────────────────

    def run(self) -> SyncStats:
        # 색상 동기화에 필요한 GSync 카테고리를 Outlook에 사전 준비
        self._ms.ensure_gsync_categories()

        if self._last_sync.is_initial():
            self._initial_sync()
        else:
            self._incremental_sync()
        return self.stats

    def rebuild_map(self) -> None:
        """
        event_map 재구성 (--rebuild-map 플래그).

        기존 매핑을 지우고 제목+시각 기반으로 재짝짓기만 실행한다.
        새 이벤트 생성은 하지 않는다.
        """
        logger.info("event_map 재구성 시작")
        self._map.clear()

        google_events = [
            e for e in self._google.list_events() if not is_google_deleted(e)
        ]
        ms_events, _ = self._ms.list_events_initial()
        ms_events = [e for e in ms_events if not is_outlook_deleted(e)]

        matched_g, matched_ms = self._pair_events(google_events, ms_events)
        self._map.save()

        unmatched_g = len(google_events) - matched_g
        unmatched_ms = len(ms_events) - matched_ms
        logger.info(
            "재구성 완료: 매핑 %d건, 미매칭 Google %d건 / MS %d건",
            len(self._map),
            unmatched_g,
            unmatched_ms,
        )
        if unmatched_g or unmatched_ms:
            logger.warning(
                "미매칭 이벤트는 다음 동기화 시 신규로 처리됩니다. "
                "캘린더에서 중복이 발생하면 수동 정리 후 다시 --rebuild-map 실행."
            )

    # ── 최초 동기화 ───────────────────────────────────────────────────────

    def _initial_sync(self) -> None:
        logger.info("최초 동기화 시작")
        now_utc = datetime.now(timezone.utc).isoformat()

        google_events = self._google.list_events()
        ms_events, delta_link = self._ms.list_events_initial()

        active_google = [e for e in google_events if not is_google_deleted(e)]
        active_ms = [e for e in ms_events if not is_outlook_deleted(e)]

        matched_g_ids, matched_ms_ids = self._pair_and_register(active_google, active_ms)

        # 짝 없는 Google 이벤트 → Outlook 생성
        for event in active_google:
            if event["id"] not in matched_g_ids:
                self._create_in_outlook(event)

        # 짝 없는 MS 이벤트 → Google 생성
        for event in active_ms:
            if event["id"] not in matched_ms_ids:
                self._create_in_google(event)

        self._save_state(now_utc, now_utc, delta_link)
        logger.info("최초 동기화 완료: %s", self.stats.summary())

    def _pair_and_register(
        self,
        google_events: list[dict],
        ms_events: list[dict],
    ) -> tuple[set[str], set[str]]:
        """
        제목 + UTC 시작 시각으로 짝짓기 후 event_map에 등록.
        반환: (매칭된 google_id 집합, 매칭된 outlook_id 집합)
        """
        matched_g: set[str] = set()
        matched_ms: set[str] = set()

        # MS 이벤트 인덱스: match_key → [event, ...]
        ms_index: dict[str, list[dict]] = {}
        for event in ms_events:
            key = match_key_for_outlook(event)
            if key:
                ms_index.setdefault(key, []).append(event)

        for g_event in google_events:
            key = match_key_for_google(g_event)
            if not key or key not in ms_index:
                continue

            candidates = ms_index[key]
            if len(candidates) > 1:
                logger.warning(
                    "중복 매칭 발견 (키=%s) — 수동 확인 필요, 자동 매핑 건너뜀", key
                )
                continue

            ms_event = candidates[0]
            self._map.set(
                g_event["id"],
                ms_event["id"],
                g_event.get("updated", ""),
                ms_event.get("lastModifiedDateTime", ""),
            )
            matched_g.add(g_event["id"])
            matched_ms.add(ms_event["id"])

        return matched_g, matched_ms

    def _pair_events(
        self, google_events: list[dict], ms_events: list[dict]
    ) -> tuple[int, int]:
        """짝짓기 후 매칭 건수 반환 (rebuild_map 전용)."""
        matched_g, matched_ms = self._pair_and_register(google_events, ms_events)
        return len(matched_g), len(matched_ms)

    # ── 증분 동기화 ───────────────────────────────────────────────────────

    def _incremental_sync(self) -> None:
        logger.info("증분 동기화 시작 (마지막 동기화: %s)", self._last_sync.last_sync_at)
        now_utc = datetime.now(timezone.utc).isoformat()

        google_changes = self._google.list_events(
            updated_min=self._last_sync.google_updated_min
        )
        ms_changes, new_delta_link = self._ms.list_events_delta(
            self._last_sync.ms_delta_link
        )

        # MS 변경을 outlook_id로 인덱싱 (충돌 해결에 사용)
        ms_change_by_id: dict[str, dict] = {e["id"]: e for e in ms_changes}

        # Google 변경 처리 (충돌 시 ms_change_by_id에서 제거)
        self._apply_google_changes(google_changes, ms_change_by_id)

        # 남은 MS 변경 처리 (Google에서 이미 처리된 항목 제외)
        self._apply_ms_changes(list(ms_change_by_id.values()))

        self._save_state(now_utc, now_utc, new_delta_link or self._last_sync.ms_delta_link)
        logger.info("증분 동기화 완료: %s", self.stats.summary())

    def _apply_google_changes(
        self,
        changes: list[dict],
        ms_change_by_id: dict[str, dict],
    ) -> None:
        for event in changes:
            g_id = event["id"]
            entry = self._map.get(g_id)

            if is_google_deleted(event):
                if entry:
                    self._delete_from_outlook(g_id, entry["outlook_id"])
                    ms_change_by_id.pop(entry["outlook_id"], None)
                continue

            current_g_modified = event.get("updated", "")

            if entry is None:
                self._create_in_outlook(event)
                continue

            if current_g_modified == entry["google_modified"]:
                # 우리가 동기화한 결과물 → 무한 루프 방지를 위해 건너뜀
                continue

            # 충돌 검사: 같은 이벤트가 MS 쪽도 변경됐는지 확인
            ms_counterpart = ms_change_by_id.get(entry["outlook_id"])
            if ms_counterpart and not is_outlook_deleted(ms_counterpart):
                ms_modified = ms_counterpart.get("lastModifiedDateTime", "")
                if current_g_modified >= ms_modified:
                    # Google이 더 최신 → Outlook 업데이트, MS 변경은 처리 완료로 표시
                    self._update_in_outlook(event, entry)
                    ms_change_by_id.pop(entry["outlook_id"])
                # else: MS가 더 최신 → ms_change_by_id에 남겨 _apply_ms_changes에서 처리
            else:
                self._update_in_outlook(event, entry)

    def _apply_ms_changes(self, changes: list[dict]) -> None:
        for event in changes:
            ms_id = event["id"]
            g_id = self._map.find_by_outlook_id(ms_id)
            entry = self._map.get(g_id) if g_id else None

            if is_outlook_deleted(event):
                if g_id:
                    self._delete_from_google(g_id, ms_id)
                continue

            current_ms_modified = event.get("lastModifiedDateTime", "")

            if g_id is None or entry is None:
                self._create_in_google(event)
                continue

            if current_ms_modified == entry["outlook_modified"]:
                continue

            self._update_in_google(event, g_id, entry)

    # ── 개별 작업 ─────────────────────────────────────────────────────────

    def _create_in_outlook(self, g_event: dict) -> None:
        try:
            body = google_event_to_outlook(g_event)
            created = self._ms.create_event(body)
            # 서버가 반환한 수정 시각을 저장해야 루프 방지가 정상 동작한다
            self._map.set(
                g_event["id"],
                created["id"],
                g_event.get("updated", ""),
                created.get("lastModifiedDateTime", ""),
            )
            self.stats.added += 1
            logger.debug("Outlook 생성: %s", g_event.get("summary", ""))
        except Exception as e:
            logger.error("Outlook 생성 실패 (google_id=%s): %s", g_event.get("id"), e)
            self.stats.errors += 1

    def _create_in_google(self, ms_event: dict) -> None:
        try:
            body = outlook_event_to_google(ms_event)
            created = self._google.create_event(body)
            self._map.set(
                created["id"],
                ms_event["id"],
                created.get("updated", ""),
                ms_event.get("lastModifiedDateTime", ""),
            )
            self.stats.added += 1
            logger.debug("Google 생성: %s", ms_event.get("subject", ""))
        except Exception as e:
            logger.error("Google 생성 실패 (outlook_id=%s): %s", ms_event.get("id"), e)
            self.stats.errors += 1

    def _update_in_outlook(self, g_event: dict, entry: dict) -> None:
        try:
            body = google_event_to_outlook(g_event)
            updated = self._ms.update_event(entry["outlook_id"], body)
            self._map.set(
                g_event["id"],
                entry["outlook_id"],
                g_event.get("updated", ""),
                updated.get("lastModifiedDateTime", ""),
            )
            self.stats.updated += 1
            logger.debug("Outlook 수정: %s", g_event.get("summary", ""))
        except Exception as e:
            logger.error(
                "Outlook 수정 실패 (google_id=%s): %s", g_event.get("id"), e
            )
            self.stats.errors += 1

    def _update_in_google(self, ms_event: dict, g_id: str, entry: dict) -> None:
        try:
            body = outlook_event_to_google(ms_event)
            updated = self._google.update_event(g_id, body)
            self._map.set(
                g_id,
                ms_event["id"],
                updated.get("updated", ""),
                ms_event.get("lastModifiedDateTime", ""),
            )
            self.stats.updated += 1
            logger.debug("Google 수정: %s", ms_event.get("subject", ""))
        except Exception as e:
            logger.error(
                "Google 수정 실패 (outlook_id=%s): %s", ms_event.get("id"), e
            )
            self.stats.errors += 1

    def _delete_from_outlook(self, g_id: str, outlook_id: str) -> None:
        try:
            self._ms.delete_event(outlook_id)
            self._map.remove(g_id)  # 즉시 제거 (다음 sync에서 404 루프 방지)
            self.stats.deleted += 1
            logger.debug("Outlook 삭제: outlook_id=%s", outlook_id)
        except Exception as e:
            logger.error("Outlook 삭제 실패 (outlook_id=%s): %s", outlook_id, e)
            self.stats.errors += 1

    def _delete_from_google(self, g_id: str, outlook_id: str) -> None:
        try:
            self._google.delete_event(g_id)
            self._map.remove(g_id)
            self.stats.deleted += 1
            logger.debug("Google 삭제: google_id=%s", g_id)
        except Exception as e:
            logger.error("Google 삭제 실패 (google_id=%s): %s", g_id, e)
            self.stats.errors += 1

    # ── 상태 저장 ─────────────────────────────────────────────────────────

    def _save_state(
        self, last_sync_at: str, google_updated_min: str, ms_delta_link: str
    ) -> None:
        self._map.save()
        self._last_sync.last_sync_at = last_sync_at
        self._last_sync.google_updated_min = google_updated_min
        self._last_sync.ms_delta_link = ms_delta_link
        self._last_sync.save()
