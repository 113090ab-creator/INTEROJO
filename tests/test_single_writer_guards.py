import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))
sys.path.insert(0, str(PROJECT_ROOT / "scripts"))

import app  # noqa: E402
import refresh_cloud_snapshots  # noqa: E402


def clear_app_snapshot_caches() -> None:
    for func_name in ("read_cloud_snapshot_context_cached", "read_cloud_snapshot_csv"):
        func = getattr(app, func_name, None)
        clear = getattr(func, "clear", None)
        if callable(clear):
            clear()


def minimal_shortage_df() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "거래처": ["고객"],
            "이니셜": ["AA"],
            "품목코드": ["P12345"],
            "제품명": ["제품"],
            "납기일": ["2026-09-10"],
            "부족수량": [10],
            "R코드": ["R12345"],
            "Q코드": ["Q12345"],
            "공정재고 합계": [3],
        }
    )


def minimal_file_info(plan_updated_at: str, wip_updated_at: str) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "재고파일": [app.format_wip_inventory_snapshot_source_label(wip_updated_at)],
            "수요파일": [f"APS API ({app.format_reference_timestamp(plan_updated_at)})"],
            "행수(현황표)": [1],
        }
    )


class SingleWriterGuardTests(unittest.TestCase):
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

    def test_legacy_aps_refresh_entrypoint_is_blocked(self) -> None:
        script = PROJECT_ROOT / "scripts" / "refresh_aps_shortage_snapshots.py"
        result = subprocess.run(
            [sys.executable, str(script), "--scheduled", "--only-if-stale"],
            cwd=PROJECT_ROOT,
            capture_output=True,
            text=True,
            timeout=10,
        )

        self.assertNotEqual(result.returncode, 0)
        self.assertIn("legacy APS flat snapshot refresher is retired", result.stderr)
        self.assertIn("scripts", result.stderr)

    def test_app_flat_writes_are_blocked_outside_vss_publish(self) -> None:
        plan_updated_at = "2026-09-09 15:53:12"
        wip_updated_at = "2026-09-09 16:09:13"
        wip = pd.DataFrame({"품목코드": ["R12345"], "창고": ["사출창고"], "재공코드": ["R12345"], "재고량": [1]})

        with self.assertRaisesRegex(RuntimeError, "VSS-managed production snapshots"):
            app.write_cloud_snapshot_csv("sets/test/shortage_snapshot.csv.gz", minimal_shortage_df())
        with self.assertRaisesRegex(RuntimeError, "VSS-managed production snapshots"):
            app.write_cloud_shortage_snapshot(
                minimal_shortage_df(),
                minimal_file_info(plan_updated_at, wip_updated_at),
                pd.DataFrame({"공정창고": ["사출창고"]}),
                plan_updated_at,
                "전체",
            )
        with self.assertRaisesRegex(RuntimeError, "VSS-managed production snapshots"):
            app.write_cloud_wip_inventory_snapshot(
                wip,
                wip_updated_at,
                app.format_wip_inventory_snapshot_source_label(wip_updated_at),
            )
        with self.assertRaisesRegex(RuntimeError, "VSS-managed production snapshot metadata"):
            app.write_cloud_snapshot_meta_value("data_updated_at", plan_updated_at)

    def test_vss_publish_can_update_current_pointer_and_flat_compat(self) -> None:
        plan_updated_at = "2026-09-09 15:53:12"
        wip_updated_at = "2026-09-09 16:09:13"
        wip = pd.DataFrame({"품목코드": ["R12345"], "창고": ["사출창고"], "재공코드": ["R12345"], "재고량": [1]})
        shortage_results = {
            "전체": (
                minimal_shortage_df(),
                minimal_file_info(plan_updated_at, wip_updated_at),
                pd.DataFrame({"공정창고": ["사출창고"]}),
            )
        }

        manifest = app.write_validated_snapshot_set(
            plan_updated_at,
            wip_updated_at,
            ["전체"],
            {"전체": pd.DataFrame({"oper_id": ["80"], "item_id": ["P12345"], "plan_qty": [10]})},
            wip,
            app.format_wip_inventory_snapshot_source_label(wip_updated_at),
            shortage_results,
            update_flat_compat=True,
        )
        clear_app_snapshot_caches()

        current = json.loads((app.CLOUD_SNAPSHOT_DIR / app.CURRENT_SNAPSHOT_SET_NAME).read_text(encoding="utf-8"))
        self.assertEqual(current["set_id"], manifest["set_id"])
        self.assertTrue((app.CLOUD_SNAPSHOT_DIR / app.WIP_INVENTORY_SNAPSHOT_FILE).exists())
        self.assertTrue((app.CLOUD_SNAPSHOT_DIR / "shortage_snapshot.csv.gz").exists())
        self.assertEqual(app.get_cloud_shortage_snapshot_updated_at("전체"), plan_updated_at)
        self.assertEqual(app.get_cloud_wip_inventory_snapshot_updated_at(), wip_updated_at)

    def test_refresh_cloud_snapshots_blocks_vss_managed_outputs(self) -> None:
        with self.assertRaisesRegex(RuntimeError, "VSS-managed APS snapshot writes are blocked"):
            refresh_cloud_snapshots.write_snapshot("shortage_snapshot.csv.gz", minimal_shortage_df())
        with self.assertRaisesRegex(RuntimeError, "VSS-managed APS snapshot writes are blocked"):
            refresh_cloud_snapshots.write_snapshot("sets/example/shortage_snapshot.csv.gz", minimal_shortage_df())


if __name__ == "__main__":
    unittest.main()
