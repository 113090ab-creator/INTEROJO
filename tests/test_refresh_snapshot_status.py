import json
import os
import sys
import tempfile
import unittest
from datetime import datetime
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))
sys.path.insert(0, str(PROJECT_ROOT / "scripts"))

import app  # noqa: E402
import refresh_snapshot  # noqa: E402


def clear_app_snapshot_caches() -> None:
    for func_name in ("read_cloud_snapshot_context_cached", "read_cloud_snapshot_csv"):
        func = getattr(app, func_name, None)
        clear = getattr(func, "clear", None)
        if callable(clear):
            clear()


class SlotDecisionApp:
    def __init__(self, published_slot: str, target_slot: str) -> None:
        self.published_slot = published_slot
        self.target_slot = target_slot

    def get_operational_target_snapshot_slot(self) -> dict[str, object]:
        return {"slot_key": self.target_slot}

    def get_published_snapshot_slot_key(self) -> str:
        return self.published_slot

    def snapshot_slot_is_at_least(self, slot_key: object, target_slot_key: object) -> bool:
        return app.snapshot_slot_is_at_least(slot_key, target_slot_key)


class RefreshSnapshotStatusTests(unittest.TestCase):
    def setUp(self) -> None:
        self.original_snapshot_dir = app.CLOUD_SNAPSHOT_DIR
        self.original_backend = os.environ.get("SNAPSHOT_STORAGE_BACKEND")
        os.environ["SNAPSHOT_STORAGE_BACKEND"] = "local"
        self.temp_dir = tempfile.TemporaryDirectory()
        app.CLOUD_SNAPSHOT_DIR = Path(self.temp_dir.name)
        clear_app_snapshot_caches()

    def tearDown(self) -> None:
        app.CLOUD_SNAPSHOT_DIR = self.original_snapshot_dir
        if self.original_backend is None:
            os.environ.pop("SNAPSHOT_STORAGE_BACKEND", None)
        else:
            os.environ["SNAPSHOT_STORAGE_BACKEND"] = self.original_backend
        clear_app_snapshot_caches()
        self.temp_dir.cleanup()

    def read_status(self) -> dict[str, object]:
        path = app.CLOUD_SNAPSHOT_DIR / app.CLOUD_SNAPSHOT_REFRESH_STATUS_NAME
        return json.loads(path.read_text(encoding="utf-8"))

    def write_status_file(self, payload: dict[str, object]) -> None:
        path = app.CLOUD_SNAPSHOT_DIR / app.CLOUD_SNAPSHOT_REFRESH_STATUS_NAME
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    def test_target_slot_needs_refresh_when_am_missing_at_0820(self) -> None:
        fake = SlotDecisionApp("2026-09-08 PM", "2026-09-09 AM")
        self.assertFalse(refresh_snapshot.target_snapshot_set_is_published(fake, "2026-09-09 AM"))

    def test_target_slot_needs_refresh_when_late_am_run_arrives_at_1015(self) -> None:
        fake = SlotDecisionApp("2026-09-08 PM", "2026-09-09 AM")
        self.assertFalse(refresh_snapshot.target_snapshot_set_is_published(fake, "2026-09-09 AM"))

    def test_target_slot_skips_when_am_already_published(self) -> None:
        fake = SlotDecisionApp("2026-09-09 AM", "2026-09-09 AM")
        self.assertTrue(refresh_snapshot.target_snapshot_set_is_published(fake, "2026-09-09 AM"))

    def test_target_slot_needs_refresh_when_pm_missing_after_late_run(self) -> None:
        fake = SlotDecisionApp("2026-09-09 AM", "2026-09-09 PM")
        self.assertFalse(refresh_snapshot.target_snapshot_set_is_published(fake, "2026-09-09 PM"))

    def test_target_slot_skips_when_pm_already_published(self) -> None:
        fake = SlotDecisionApp("2026-09-09 PM", "2026-09-09 PM")
        self.assertTrue(refresh_snapshot.target_snapshot_set_is_published(fake, "2026-09-09 PM"))

    def test_same_slot_metadata_does_not_regress_to_missing_value(self) -> None:
        self.write_status_file(
            {
                "checked_at": "2026-09-09 09:39:31",
                "status": app.REFRESH_STATUS_FAILED,
                "slot_key": "2026-09-09 07:30",
                "api_updated_at": "2026-09-09 07:55:00",
                "wip_api_updated_at": "2026-09-09 08:27:49",
            }
        )

        refresh_snapshot.write_status(
            app,
            app.REFRESH_STATUS_WAITING_FOR_WIP,
            slot_key="2026-09-09 AM",
            api_updated_at="2026-09-09 07:55:00",
            wip_api_updated_at="-",
        )

        status = self.read_status()
        self.assertEqual(status["status"], app.REFRESH_STATUS_WAITING_FOR_WIP)
        self.assertEqual(status["wip_api_updated_at"], "2026-09-09 08:27:49")

    def test_new_target_slot_does_not_inherit_previous_slot_metadata(self) -> None:
        self.write_status_file(
            {
                "checked_at": "2026-09-08 20:51:51",
                "status": app.REFRESH_STATUS_PUBLISHED,
                "slot_key": "2026-09-08 PM",
                "api_updated_at": "2026-09-08 15:52:09",
                "wip_api_updated_at": "2026-09-08 16:13:07",
            }
        )

        refresh_snapshot.write_status(
            app,
            app.REFRESH_STATUS_WAITING_FOR_PLAN,
            slot_key="2026-09-09 AM",
            api_updated_at="-",
            wip_api_updated_at="-",
        )

        status = self.read_status()
        self.assertEqual(status["slot_key"], "2026-09-09 AM")
        self.assertEqual(status["api_updated_at"], "-")
        self.assertEqual(status["wip_api_updated_at"], "-")

    def test_checking_status_is_not_persisted(self) -> None:
        refresh_snapshot.write_status(
            app,
            app.REFRESH_STATUS_CHECKING,
            slot_key="2026-09-09 AM",
            api_updated_at="2026-09-09 07:55:00",
            wip_api_updated_at="-",
        )

        self.assertFalse((app.CLOUD_SNAPSHOT_DIR / app.CLOUD_SNAPSHOT_REFRESH_STATUS_NAME).exists())

    def test_waiting_status_can_replace_transient_checking(self) -> None:
        refresh_snapshot.write_status(
            app,
            app.REFRESH_STATUS_CHECKING,
            slot_key="2026-09-09 AM",
            api_updated_at="2026-09-09 07:55:00",
            wip_api_updated_at="-",
        )
        refresh_snapshot.write_status(
            app,
            app.REFRESH_STATUS_WAITING_FOR_WIP,
            slot_key="2026-09-09 AM",
            api_updated_at="2026-09-09 07:55:00",
            wip_api_updated_at="-",
        )

        self.assertEqual(self.read_status()["status"], app.REFRESH_STATUS_WAITING_FOR_WIP)

    def test_failed_status_can_replace_transient_checking(self) -> None:
        refresh_snapshot.write_status(
            app,
            app.REFRESH_STATUS_CHECKING,
            slot_key="2026-09-09 AM",
            api_updated_at="2026-09-09 07:55:00",
            wip_api_updated_at="-",
        )
        refresh_snapshot.write_status(
            app,
            app.REFRESH_STATUS_FAILED,
            slot_key="2026-09-09 AM",
            api_updated_at="2026-09-09 07:55:00",
            wip_api_updated_at="-",
            reason="test failure",
        )

        self.assertEqual(self.read_status()["status"], app.REFRESH_STATUS_FAILED)

    def test_operational_target_slot_examples(self) -> None:
        cases = [
            ("2026-09-09 08:20:00", "2026-09-09 AM"),
            ("2026-09-09 10:15:00", "2026-09-09 AM"),
            ("2026-09-09 18:30:00", "2026-09-09 PM"),
        ]
        for value, expected in cases:
            with self.subTest(value=value):
                now = datetime.strptime(value, "%Y-%m-%d %H:%M:%S").replace(tzinfo=app.DISPLAY_TZ)
                self.assertEqual(app.get_operational_target_snapshot_slot(now)["slot_key"], expected)


if __name__ == "__main__":
    unittest.main()
