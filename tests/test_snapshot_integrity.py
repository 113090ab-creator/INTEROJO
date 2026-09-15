import json
import sys
import tempfile
import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "scripts"))

import validate_snapshot_integrity  # noqa: E402


class SnapshotIntegrityTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.snapshot_dir = Path(self.temp_dir.name)
        self.set_id = "aps_20260915_pm_plan155834_wip161621_41958bce"
        self.manifest_path = f"sets/{self.set_id}/manifest.json"
        (self.snapshot_dir / "sets" / self.set_id).mkdir(parents=True)
        self.write_json(
            self.snapshot_dir / self.manifest_path,
            {
                "set_id": self.set_id,
                "slot_key": "2026-09-15 PM",
                "plan_updated_at": "2026-09-15 15:58:34",
                "wip_updated_at": "2026-09-15 16:16:21",
                "created_at": "2026-09-15 18:15:15",
            },
        )
        self.write_json(
            self.snapshot_dir / "current_snapshot_set.json",
            {
                "set_id": self.set_id,
                "slot_key": "2026-09-15 PM",
                "manifest_path": self.manifest_path,
                "plan_updated_at": "2026-09-15 15:58:34",
                "wip_updated_at": "2026-09-15 16:16:21",
                "published_at": "2026-09-15 18:15:19",
            },
        )
        self.write_json(
            self.snapshot_dir / "aps_snapshot_refresh_status.json",
            {
                "status": "PUBLISHED",
                "slot_key": "2026-09-15 PM",
                "set_id": self.set_id,
            },
        )

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def write_json(self, path: Path, payload: dict[str, object]) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    def validate(self) -> None:
        validate_snapshot_integrity.validate_snapshot_integrity(self.snapshot_dir)

    def test_valid_snapshot_set_passes(self) -> None:
        self.validate()

    def test_conflict_marker_in_status_fails(self) -> None:
        (self.snapshot_dir / "aps_snapshot_refresh_status.json").write_text(
            '{\n<<<<<<< HEAD\n  "status": "PUBLISHED"\n=======\n  "status": "FAILED"\n>>>>>>> branch\n}\n',
            encoding="utf-8",
        )

        with self.assertRaises(validate_snapshot_integrity.SnapshotIntegrityError):
            self.validate()

    def test_invalid_current_json_fails(self) -> None:
        (self.snapshot_dir / "current_snapshot_set.json").write_text("{not-json", encoding="utf-8")

        with self.assertRaises(validate_snapshot_integrity.SnapshotIntegrityError):
            self.validate()

    def test_published_status_must_match_current_set(self) -> None:
        self.write_json(
            self.snapshot_dir / "aps_snapshot_refresh_status.json",
            {
                "status": "PUBLISHED",
                "slot_key": "2026-09-15 AM",
                "set_id": "aps_20260915_am_plan074833_wip080309_bcc6dcef",
            },
        )

        with self.assertRaises(validate_snapshot_integrity.SnapshotIntegrityError):
            self.validate()


if __name__ == "__main__":
    unittest.main()
