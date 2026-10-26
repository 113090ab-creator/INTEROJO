import sys
import io
import json
import unittest
from datetime import datetime
from pathlib import Path
from unittest import mock


sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
import check_app_snapshot as probe


class AppSnapshotProbeTests(unittest.TestCase):
    def test_caption_parser_reads_the_actual_set_and_operational_status(self):
        actual = probe.parse_rendered_snapshot("PLAN: 2026-10-26 · Set: aps_new · 상태: 최신")
        self.assertEqual(actual, {"set_id": "aps_new", "display_status": "최신"})

    def test_successful_http_page_without_snapshot_is_not_healthy(self):
        self.assertEqual(probe.parse_rendered_snapshot("생산현황 Loading"), {})

    def test_old_app_set_fails_even_after_successful_publish(self):
        ok, _ = probe.check_rendered_snapshot(
            {"set_id": "aps_new", "slot_key": "2026-10-26 AM"},
            {"set_id": "aps_old", "display_status": "최신"},
            datetime(2026, 10, 26, 11, tzinfo=probe.KST))
        self.assertFalse(ok)

    def test_matching_current_set_passes(self):
        ok, _ = probe.check_rendered_snapshot(
            {"set_id": "aps_new", "slot_key": "2026-10-26 AM"},
            {"set_id": "aps_new", "display_status": "최신"},
            datetime(2026, 10, 26, 11, tzinfo=probe.KST))
        self.assertTrue(ok)

    def test_matching_but_overdue_published_set_fails(self):
        ok, _ = probe.check_rendered_snapshot(
            {"set_id": "aps_old", "slot_key": "2026-10-25 PM"},
            {"set_id": "aps_old", "display_status": "최신"},
            datetime(2026, 10, 26, 11, tzinfo=probe.KST))
        self.assertFalse(ok)

    def test_waiting_during_the_allowed_refresh_window_is_not_a_false_alarm(self):
        for hour in (8, 9):
            with self.subTest(hour=hour):
                ok, _ = probe.check_rendered_snapshot(
                    {"set_id": "aps_previous", "slot_key": "2026-10-25 PM"},
                    {"set_id": "aps_previous", "display_status": "갱신 중"},
                    datetime(2026, 10, 26, hour, 30, tzinfo=probe.KST))
                self.assertTrue(ok)

    def test_unhealthy_status_on_published_current_set_fails(self):
        ok, _ = probe.check_rendered_snapshot(
            {"set_id": "aps_new", "slot_key": "2026-10-26 AM"},
            {"set_id": "aps_new", "display_status": "갱신 실패"},
            datetime(2026, 10, 26, 11, tzinfo=probe.KST))
        self.assertFalse(ok)

    def test_deadlines_cover_evening_and_before_morning(self):
        for hour, slot in ((7, "2026-10-25 PM"), (10, "2026-10-26 AM"),
                           (17, "2026-10-26 AM"), (18, "2026-10-26 PM")):
            with self.subTest(hour=hour):
                self.assertEqual(probe.required_slot(datetime(2026, 10, 26, hour, tzinfo=probe.KST)), slot)

    def test_probe_adopts_a_publish_that_finishes_during_the_browser_check(self):
        expected = {"set_id": "aps_morning", "slot_key": "2026-10-26 AM"}
        current = {"set_id": "aps_evening", "slot_key": "2026-10-26 PM"}
        with mock.patch.object(probe.urllib.request, "urlopen", return_value=io.BytesIO(json.dumps(current).encode())):
            probe.refresh_expected_pointer(expected)
        self.assertEqual(expected, current)

    def test_probe_does_not_regress_when_the_remote_index_is_stale(self):
        expected = {"set_id": "aps_evening", "slot_key": "2026-10-26 PM"}
        current = {"set_id": "aps_morning", "slot_key": "2026-10-26 AM"}
        with mock.patch.object(probe.urllib.request, "urlopen", return_value=io.BytesIO(json.dumps(current).encode())):
            probe.refresh_expected_pointer(expected)
        self.assertEqual(expected["set_id"], "aps_evening")


if __name__ == "__main__":
    unittest.main()
