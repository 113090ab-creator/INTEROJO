import json
import os
import sys
import tempfile
import unittest
from pathlib import Path

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

import app  # noqa: E402


def clear_app_snapshot_caches() -> None:
    for func_name in ("read_cloud_snapshot_context_cached", "read_cloud_snapshot_csv"):
        func = getattr(app, func_name, None)
        clear = getattr(func, "clear", None)
        if callable(clear):
            clear()


def minimal_shortage_df(rows: int = 1) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "거래처": ["고객"] * rows,
            "이니셜": ["AA"] * rows,
            "품목코드": ["P12345"] * rows,
            "제품명": ["제품"] * rows,
            "납기일": ["2026-09-10"] * rows,
            "부족수량": [10] * rows,
            "R코드": ["R12345"] * rows,
            "Q코드": ["Q12345"] * rows,
            "공정재고 합계": [3] * rows,
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


class ValidatedSnapshotSetTests(unittest.TestCase):
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

    def publish_current_set(self, published_at: str = "2026-09-05 10:39:13") -> dict[str, object]:
        plan_updated_at = "2026-09-05 08:01:23"
        wip_updated_at = "2026-09-05 08:15:16"
        manifest = app.write_validated_snapshot_set(
            plan_updated_at,
            wip_updated_at,
            ["전체"],
            {"전체": pd.DataFrame({"oper_id": ["80"], "item_id": ["P12345"], "plan_qty": [10]})},
            pd.DataFrame({"품목코드": ["R12345"], "창고": ["사출창고"], "재공코드": ["R12345"], "재고량": [1]}),
            app.format_wip_inventory_snapshot_source_label(wip_updated_at),
            {
                "전체": (
                    minimal_shortage_df(),
                    minimal_file_info(plan_updated_at, wip_updated_at),
                    pd.DataFrame({"공정창고": ["사출창고"]}),
                )
            },
            update_flat_compat=False,
        )
        current_path = app.CLOUD_SNAPSHOT_DIR / app.CURRENT_SNAPSHOT_SET_NAME
        current = json.loads(current_path.read_text(encoding="utf-8"))
        current["published_at"] = published_at
        current_path.write_text(json.dumps(current, ensure_ascii=False, indent=2), encoding="utf-8")
        clear_app_snapshot_caches()
        return manifest

    def write_refresh_status(self, payload: dict[str, object]) -> None:
        status_path = app.CLOUD_SNAPSHOT_DIR / app.CLOUD_SNAPSHOT_REFRESH_STATUS_NAME
        status_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        clear_app_snapshot_caches()

    def test_slot_comparison_waiting_and_ready_states(self) -> None:
        self.assertEqual(
            app.compare_aps_snapshot_slots("2026-09-05 08:01:23", "2026-09-04 16:10:53"),
            app.REFRESH_STATUS_WAITING_FOR_WIP,
        )
        self.assertEqual(
            app.compare_aps_snapshot_slots("2026-09-05 08:01:23", "2026-09-05 08:15:16"),
            app.REFRESH_STATUS_READY,
        )
        self.assertEqual(
            app.compare_aps_snapshot_slots("2026-09-05 08:01:23", "2026-09-05 16:10:53"),
            app.REFRESH_STATUS_WAITING_FOR_PLAN,
        )

    def test_shortage_builder_uses_staged_wip_not_published_wip(self) -> None:
        raw_plan = pd.DataFrame(
            {
                "res_site_id": ["C"] * 2,
                "cust_name": ["고객"] * 2,
                "so_id": ["SO1"] * 2,
                "initial": ["AA"] * 2,
                "item_id": ["P75085"] * 2,
                "demand_item_name": ["제품"] * 2,
                "demand_qty": [10, 10],
                "oper_id": ["10", "80"],
                "plan_qty": [4, 10],
                "due_date": ["2026-09-10", "2026-09-10"],
            }
        )
        staged_wip = pd.DataFrame(
            {
                "품목코드": ["R75085", "P75085"],
                "창고": ["사출창고", "누수규격검사"],
                "재공코드": ["R75085", "P75085"],
                "재고량": [7, 2],
            }
        )
        original_loader = app.load_all_item_inventory_source_with_label

        def fail_if_published_wip_is_read(_base_dir):
            raise AssertionError("published WIP snapshot should not be read during staged build")

        app.load_all_item_inventory_source_with_label = fail_if_published_wip_is_read
        try:
            df, file_info_df, _ = app.build_api_shortage_data_from_frames(
                raw_plan,
                "unit-test-staged-wip",
                str(PROJECT_ROOT),
                "C관",
                plan_updated_at="2026-09-05 08:01:23",
                inventory_df=staged_wip,
                inventory_source_label=app.format_wip_inventory_snapshot_source_label("2026-09-05 08:15:16"),
            )
        finally:
            app.load_all_item_inventory_source_with_label = original_loader

        self.assertFalse(df.empty)
        self.assertEqual(file_info_df.iloc[0]["재고파일"], app.format_wip_inventory_snapshot_source_label("2026-09-05 08:15:16"))
        self.assertEqual(float(df.iloc[0]["사출창고"]), 7.0)

    def test_validated_set_publish_and_current_pointer_are_atomic_inputs(self) -> None:
        plan_updated_at = "2026-09-05 08:01:23"
        wip_updated_at = "2026-09-05 08:15:16"
        wip = pd.DataFrame(
            {
                "품목코드": ["R12345"],
                "창고": ["사출창고"],
                "재공코드": ["R12345"],
                "재고량": [7],
            }
        )
        plan_frames = {"전체": pd.DataFrame({"oper_id": ["80"], "item_id": ["P12345"], "plan_qty": [10]})}
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
            plan_frames,
            wip,
            app.format_wip_inventory_snapshot_source_label(wip_updated_at),
            shortage_results,
            update_flat_compat=False,
        )
        clear_app_snapshot_caches()

        current = app.get_published_snapshot_set_pointer()
        loaded_shortage, loaded_info, _ = app.load_cloud_shortage_snapshot("전체")
        self.assertEqual(current["set_id"], manifest["set_id"])
        self.assertEqual(app.get_cloud_shortage_snapshot_updated_at("전체"), plan_updated_at)
        self.assertEqual(app.get_cloud_wip_inventory_snapshot_updated_at(), wip_updated_at)
        self.assertEqual(len(loaded_shortage), 1)
        self.assertIn(app.format_reference_timestamp(wip_updated_at), loaded_info.iloc[0]["재고파일"])

    def test_invalid_or_partial_set_does_not_move_current_pointer(self) -> None:
        current_path = app.CLOUD_SNAPSHOT_DIR / app.CURRENT_SNAPSHOT_SET_NAME
        self.assertFalse(current_path.exists())
        with self.assertRaises(ValueError):
            app.write_validated_snapshot_set(
                "2026-09-05 08:01:23",
                "2026-09-05 16:10:53",
                ["전체"],
                {"전체": pd.DataFrame({"x": [1]})},
                pd.DataFrame({"품목코드": ["R12345"], "창고": ["사출창고"], "재공코드": ["R12345"], "재고량": [1]}),
                app.format_wip_inventory_snapshot_source_label("2026-09-05 16:10:53"),
                {"전체": (minimal_shortage_df(), minimal_file_info("2026-09-05 08:01:23", "2026-09-05 16:10:53"), pd.DataFrame())},
                update_flat_compat=False,
            )
        self.assertFalse(current_path.exists())

        with self.assertRaises(ValueError):
            app.write_validated_snapshot_set(
                "2026-09-05 08:01:23",
                "2026-09-05 08:15:16",
                ["C관", "A관"],
                {"C관": pd.DataFrame({"x": [1]})},
                pd.DataFrame({"품목코드": ["R12345"], "창고": ["사출창고"], "재공코드": ["R12345"], "재고량": [1]}),
                app.format_wip_inventory_snapshot_source_label("2026-09-05 08:15:16"),
                {"C관": (minimal_shortage_df(), minimal_file_info("2026-09-05 08:01:23", "2026-09-05 08:15:16"), pd.DataFrame())},
                update_flat_compat=False,
            )
        self.assertFalse(current_path.exists())

    def test_write_failure_does_not_move_current_pointer(self) -> None:
        current_path = app.CLOUD_SNAPSHOT_DIR / app.CURRENT_SNAPSHOT_SET_NAME
        plan_updated_at = "2026-09-05 08:01:23"
        wip_updated_at = "2026-09-05 08:15:16"
        original_writer = app.write_cloud_snapshot_csv

        def fail_shortage_write(name, df):
            if "shortage_snapshot" in str(name):
                return False
            return original_writer(name, df)

        app.write_cloud_snapshot_csv = fail_shortage_write
        try:
            with self.assertRaises(RuntimeError):
                app.write_validated_snapshot_set(
                    plan_updated_at,
                    wip_updated_at,
                    ["전체"],
                    {"전체": pd.DataFrame({"x": [1]})},
                    pd.DataFrame({"품목코드": ["R12345"], "창고": ["사출창고"], "재공코드": ["R12345"], "재고량": [1]}),
                    app.format_wip_inventory_snapshot_source_label(wip_updated_at),
                    {
                        "전체": (
                            minimal_shortage_df(),
                            minimal_file_info(plan_updated_at, wip_updated_at),
                            pd.DataFrame({"공정창고": ["사출창고"]}),
                        )
                    },
                    update_flat_compat=False,
                )
        finally:
            app.write_cloud_snapshot_csv = original_writer

        self.assertFalse(current_path.exists())

    def test_waiting_status_display_keeps_previous_validated_set(self) -> None:
        status = {"status": app.REFRESH_STATUS_WAITING_FOR_WIP, "slot_key": "2026-09-05 AM"}
        self.assertEqual(app.get_snapshot_refresh_display_status(status), "갱신 중")

    def test_stale_failed_status_after_published_set_is_ignored(self) -> None:
        self.publish_current_set()
        self.write_refresh_status(
            {
                "checked_at": "2026-09-05 09:59:15",
                "status": app.REFRESH_STATUS_FAILED,
                "api_updated_at": "2026-09-05 08:01:23",
                "wip_api_updated_at": "2026-09-04 16:10:53",
                "slot_key": "2026-09-05 AM",
                "reason": "old failure",
            }
        )

        status = app.get_snapshot_refresh_status()
        self.assertEqual(status["status"], app.REFRESH_STATUS_PUBLISHED)
        self.assertEqual(app.get_snapshot_refresh_display_status(status), "최신")
        self.assertEqual(app.build_snapshot_refresh_failure_message(), "")

    def test_newer_waiting_status_displays_refreshing_and_keeps_current_set(self) -> None:
        self.publish_current_set()
        self.write_refresh_status(
            {
                "checked_at": "2026-09-05 10:45:00",
                "status": app.REFRESH_STATUS_WAITING_FOR_WIP,
                "api_updated_at": "2026-09-05 16:01:00",
                "wip_api_updated_at": "2026-09-05 08:15:16",
                "slot_key": "2026-09-05 PM",
            }
        )

        status = app.get_snapshot_refresh_status()
        loaded_shortage, _, _ = app.load_cloud_shortage_snapshot("전체")
        self.assertEqual(status["status"], app.REFRESH_STATUS_WAITING_FOR_WIP)
        self.assertEqual(app.get_snapshot_refresh_display_status(status), "갱신 중")
        self.assertEqual(len(loaded_shortage), 1)

    def test_newer_failed_status_displays_failure_and_keeps_current_set(self) -> None:
        self.publish_current_set()
        self.write_refresh_status(
            {
                "checked_at": "2026-09-05 10:50:00",
                "status": app.REFRESH_STATUS_FAILED,
                "api_updated_at": "2026-09-05 16:01:00",
                "wip_api_updated_at": "2026-09-05 08:15:16",
                "slot_key": "2026-09-05 PM",
                "reason": "current failure",
            }
        )

        status = app.get_snapshot_refresh_status()
        loaded_shortage, _, _ = app.load_cloud_shortage_snapshot("전체")
        self.assertEqual(status["status"], app.REFRESH_STATUS_FAILED)
        self.assertEqual(app.get_snapshot_refresh_display_status(status), "갱신 실패")
        self.assertIn("데이터 갱신 실패", app.build_snapshot_refresh_failure_message())
        self.assertEqual(len(loaded_shortage), 1)

    def test_published_status_is_displayed_as_latest_not_internal_value(self) -> None:
        manifest = self.publish_current_set()
        self.write_refresh_status(
            {
                "checked_at": "2026-09-05 10:39:13",
                "published_at": "2026-09-05 10:39:13",
                "status": app.REFRESH_STATUS_PUBLISHED,
                "set_id": manifest["set_id"],
                "api_updated_at": "2026-09-05 08:01:23",
                "wip_api_updated_at": "2026-09-05 08:15:16",
                "slot_key": "2026-09-05 AM",
            }
        )

        self.assertEqual(app.get_cloud_snapshot_status_label(), "최신")
        self.assertEqual(app.get_snapshot_refresh_display_status(app.get_snapshot_refresh_status()), "최신")


if __name__ == "__main__":
    unittest.main()
