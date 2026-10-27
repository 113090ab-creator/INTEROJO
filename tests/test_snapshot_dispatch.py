import sys
import unittest
from datetime import datetime
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
import dispatch_snapshot_refresh as scheduler


class SnapshotDispatchTests(unittest.TestCase):
    def test_target_matches_operational_day_boundaries(self):
        for hour, expected in ((7, "2026-10-26 PM"), (8, "2026-10-27 AM"),
                               (15, "2026-10-27 AM"), (16, "2026-10-27 PM")):
            with self.subTest(hour=hour):
                self.assertEqual(scheduler.target_slot(datetime(2026, 10, 27, hour, tzinfo=scheduler.KST)), expected)

    @mock.patch.object(scheduler, "target_slot", return_value="2026-10-27 PM")
    def test_check_only_does_not_dispatch(self, _):
        github = mock.Mock()
        github.request.return_value = {"state": "active"}
        github.pointer.return_value = {"slot_key": "2026-10-27 AM"}
        self.assertEqual(scheduler.run(github, check_only=True), 0)
        self.assertEqual(github.request.call_args_list, [mock.call("actions/workflows/refresh_snapshot.yml")])

    @mock.patch.object(scheduler, "target_slot", return_value="2026-10-27 PM")
    def test_current_publish_does_not_start_another_refresh(self, _):
        github = mock.Mock()
        github.request.return_value = {"state": "active"}
        github.pointer.return_value = {"slot_key": "2026-10-27 PM"}
        self.assertEqual(scheduler.run(github), 0)
        self.assertEqual(github.request.call_count, 1)

    @mock.patch.object(scheduler, "target_slot", return_value="2026-10-27 PM")
    def test_existing_refresh_is_watched_and_live_check_failure_is_reported(self, _):
        github = mock.Mock()
        github.pointer.return_value = {"slot_key": "2026-10-27 AM"}
        github.request.side_effect = [
            {"state": "active"},
            {"workflow_runs": [{"id": 42, "status": "in_progress", "head_branch": "main"}]},
            {"status": "completed", "conclusion": "failure"},
        ]
        with self.assertRaisesRegex(RuntimeError, "live verification failed"):
            scheduler.run(github)
        self.assertFalse(any(len(call.args) > 1 for call in github.request.call_args_list))

    @mock.patch.object(scheduler.time, "sleep")
    @mock.patch.object(scheduler, "target_slot", return_value="2026-10-27 PM")
    def test_dispatch_waits_for_a_new_run_then_requires_target_publication(self, _, sleep):
        github = mock.Mock()
        github.pointer.side_effect = [{"slot_key": "2026-10-27 AM"}, {"slot_key": "2026-10-27 PM", "set_id": "new"}]
        github.request.side_effect = [
            {"state": "active"},
            {"workflow_runs": [{"id": 50, "status": "completed", "head_branch": "main"}]},
            {},
            {"workflow_runs": [{"id": 10, "status": "completed", "head_branch": "main"}]},
            {"workflow_runs": [{"id": 51, "status": "in_progress", "head_branch": "main"}]},
            {"status": "completed", "conclusion": "success"},
        ]
        self.assertEqual(scheduler.run(github), 0)
        self.assertIn(mock.call("actions/runs/51"), github.request.call_args_list)
        self.assertNotIn(mock.call("actions/runs/10"), github.request.call_args_list)

    @mock.patch.object(scheduler, "target_slot", return_value="2026-10-27 PM")
    def test_success_without_target_publication_is_an_error(self, _):
        github = mock.Mock()
        github.pointer.return_value = {"slot_key": "2026-10-27 AM"}
        github.request.side_effect = [
            {"state": "active"},
            {"workflow_runs": [{"id": 42, "status": "in_progress", "head_branch": "main"}]},
            {"status": "completed", "conclusion": "success"},
        ]
        with self.assertRaisesRegex(RuntimeError, "without publishing target"):
            scheduler.run(github)


if __name__ == "__main__":
    unittest.main()
