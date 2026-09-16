import sys
import unittest
import uuid
from pathlib import Path
from unittest import mock

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

import app  # noqa: E402


def build_plan_rows(order_no: str = "SO-001", plan_customer: str = "PLAN CUSTOMER", initial: str = "국내 SAMPLE") -> pd.DataFrame:
    return pd.DataFrame(
        {
            app.ORDER_NO_COL: [order_no],
            app.PLAN_CUSTOMER_COL: [plan_customer],
            "거래처": [plan_customer],
            "이니셜": [initial],
            "품목코드": ["P2026A-01.00"],
            "부족수량": [10],
            "사출생산필요수량": [4],
            "공정재고 합계": [7],
        }
    )


def build_order_raw(
    order_no: str = "SO-001",
    customer: str = "ORDER CUSTOMER",
    country_code: str = "KR",
    country_name: str = "대한민국",
    extracted_at: str = "2026-09-16T09:00:00",
) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "od_no": [order_no],
            "cust_nm": [customer],
            "e_nation": [country_code],
            "e_nation_nm": [country_name],
            "od_dt": ["2026-09-16"],
            "deli_date": ["2026-10-01"],
            "gd_deli_date": ["2026-10-02"],
            "stts": ["S"],
            "stts_label": ["저장"],
            "confirm_yn": ["1"],
            "confirm_yn_label": ["완료"],
            "extracted_at": [extracted_at],
        }
    )


class OrderStatusReferenceTests(unittest.TestCase):
    def tearDown(self) -> None:
        if hasattr(app.load_order_status_master_cached, "clear"):
            app.load_order_status_master_cached.clear()

    def test_order_no_normalization_preserves_full_exact_key(self) -> None:
        self.assertEqual(app.normalize_order_no_value(" 00123.0 "), "00123")
        self.assertEqual(app.normalize_order_no_value("so-001"), "SO-001")
        self.assertEqual(app.normalize_order_no_value("1.23E+5"), "123000")

    def test_order_status_date_range_uses_plan_order_dates(self) -> None:
        demand = pd.DataFrame(
            {
                app.ORDER_NO_COL: ["202609150001", "R202608060001", ""],
                app.ORDER_RECEIVED_DATE_COL: ["2026-04-17", "", ""],
            }
        )

        date_from, date_to = app.derive_order_status_date_range(demand)
        params = app.build_order_status_query_params(date_from, date_to)

        self.assertEqual((date_from, date_to), ("2026-04-17", "2026-09-15"))
        self.assertEqual(params["date_from"], "2026-04-17")
        self.assertEqual(params["date_to"], "2026-09-15")
        self.assertEqual(params["limit"], app.PLAN_API_DEFAULT_ROW_LIMIT)

    def test_duplicate_order_rows_are_collapsed_without_conflict(self) -> None:
        raw = pd.concat(
            [
                build_order_raw(extracted_at="2026-09-16T09:00:00"),
                build_order_raw(extracted_at="2026-09-16T10:00:00"),
            ],
            ignore_index=True,
        )

        master = app.build_order_status_master_from_raw(raw)

        self.assertEqual(len(master), 1)
        self.assertEqual(master.loc[0, app.ORDER_INFO_STATUS_COL], app.ORDER_INFO_STATUS_MATCHED)
        self.assertEqual(master.loc[0, app.ORDER_CUSTOMER_COL], "ORDER CUSTOMER")

    def test_unmatched_order_keeps_plan_customer_fallback(self) -> None:
        master = app.build_order_status_master_from_raw(build_order_raw(order_no="SO-OTHER"))

        enriched = app.apply_order_status_reference(build_plan_rows(order_no="SO-MISSING"), master)

        self.assertEqual(enriched.loc[0, app.ORDER_INFO_STATUS_COL], app.ORDER_INFO_STATUS_UNMATCHED)
        self.assertEqual(enriched.loc[0, app.FINAL_CUSTOMER_COL], "PLAN CUSTOMER")
        self.assertEqual(enriched.loc[0, app.CUSTOMER_NAME_SOURCE_COL], app.CUSTOMER_SOURCE_PLAN_API)

    def test_country_conflict_marks_duplicate_review_without_guessing(self) -> None:
        raw = pd.concat(
            [
                build_order_raw(country_code="CN", country_name="중국"),
                build_order_raw(country_code="JP", country_name="일본"),
            ],
            ignore_index=True,
        )

        master = app.build_order_status_master_from_raw(raw)
        enriched = app.apply_order_status_reference(build_plan_rows(initial="해외 SAMPLE"), master)

        self.assertEqual(master.loc[0, app.ORDER_INFO_STATUS_COL], app.ORDER_INFO_STATUS_DUPLICATE_REVIEW)
        self.assertEqual(enriched.loc[0, app.COUNTRY_CONFIRMATION_STATUS_COL], app.COUNTRY_STATUS_DUPLICATE_REVIEW)
        self.assertEqual(enriched.loc[0, app.COUNTRY_DOMESTIC_EXPORT_COL], app.DOMESTIC_EXPORT_UNKNOWN)
        self.assertEqual(enriched.loc[0, app.REGION_COL], app.REGION_UNKNOWN)

    def test_customer_conflict_marks_duplicate_review_and_falls_back_to_plan(self) -> None:
        raw = pd.concat(
            [
                build_order_raw(customer="CUSTOMER A"),
                build_order_raw(customer="CUSTOMER B"),
            ],
            ignore_index=True,
        )

        master = app.build_order_status_master_from_raw(raw)
        enriched = app.apply_order_status_reference(build_plan_rows(plan_customer="한국알콘(주)"), master)

        self.assertEqual(master.loc[0, app.ORDER_INFO_STATUS_COL], app.ORDER_INFO_STATUS_DUPLICATE_REVIEW)
        self.assertEqual(enriched.loc[0, app.ORDER_CUSTOMER_COL], "")
        self.assertEqual(enriched.loc[0, app.FINAL_CUSTOMER_COL], "한국알콘(주)")
        self.assertEqual(enriched.loc[0, app.CUSTOMER_NAME_SOURCE_COL], app.CUSTOMER_SOURCE_PLAN_API)

    def test_order_customer_has_priority_over_plan_customer_and_flags_mismatch(self) -> None:
        master = app.build_order_status_master_from_raw(build_order_raw(customer="한국알콘(주)", country_code="JP", country_name="일본"))

        enriched = app.apply_order_status_reference(build_plan_rows(plan_customer="PLAN CUSTOMER", initial="해외 SAMPLE"), master)

        self.assertEqual(enriched.loc[0, app.FINAL_CUSTOMER_COL], "한국알콘(주)")
        self.assertEqual(enriched.loc[0, app.CUSTOMER_NAME_SOURCE_COL], app.CUSTOMER_SOURCE_ORDER_API)
        self.assertEqual(enriched.loc[0, app.CUSTOMER_CONFIRMATION_STATUS_COL], app.CUSTOMER_STATUS_MISMATCH)
        self.assertEqual(enriched.loc[0, app.CUSTOMER_GROUP_COL], "Alcon")

    def test_plan_customer_fallback_when_order_customer_is_blank(self) -> None:
        master = app.build_order_status_master_from_raw(build_order_raw(customer="", country_code="JP", country_name="일본"))

        enriched = app.apply_order_status_reference(build_plan_rows(plan_customer="Alensa s.r.o", initial="해외 SAMPLE"), master)

        self.assertEqual(enriched.loc[0, app.FINAL_CUSTOMER_COL], "Alensa s.r.o")
        self.assertEqual(enriched.loc[0, app.CUSTOMER_NAME_SOURCE_COL], app.CUSTOMER_SOURCE_PLAN_API)
        self.assertEqual(enriched.loc[0, app.CUSTOMER_GROUP_COL], "ALENSA")

    def test_country_domestic_export_classification(self) -> None:
        self.assertEqual(app.classify_country_domestic_export("KR", "")[0], app.DOMESTIC_EXPORT_DOMESTIC)
        self.assertEqual(app.classify_country_domestic_export("", "일본")[0], app.DOMESTIC_EXPORT_EXPORT)
        self.assertEqual(app.classify_country_domestic_export("", "")[0], app.DOMESTIC_EXPORT_UNKNOWN)

    def test_country_overrides_initial_for_domestic_export_comparison_values(self) -> None:
        master = app.build_order_status_master_from_raw(build_order_raw(country_code="JP", country_name="일본"))

        enriched = app.apply_order_status_reference(build_plan_rows(initial="국내 SAMPLE"), master)

        self.assertEqual(enriched.loc[0, app.COUNTRY_DOMESTIC_EXPORT_COL], app.DOMESTIC_EXPORT_EXPORT)
        self.assertEqual(enriched.loc[0, app.INITIAL_DOMESTIC_EXPORT_COL], app.DOMESTIC_EXPORT_DOMESTIC)

    def test_api_failure_uses_previous_normal_cache(self) -> None:
        cached_master = app.build_order_status_master_from_raw(build_order_raw(customer="한국알콘(주)"))
        if hasattr(app.load_order_status_master_cached, "clear"):
            app.load_order_status_master_cached.clear()

        with (
            mock.patch.object(app, "read_plan_endpoint_dataframe", return_value=(pd.DataFrame(), "timeout")),
            mock.patch.object(app, "read_order_status_master_cache", return_value=cached_master),
        ):
            master, error = app.load_order_status_master_cached(f"test-{uuid.uuid4().hex}")

        self.assertEqual(len(master), 1)
        self.assertIn("이전 정상 캐시", error)

    def test_order_merge_preserves_row_count_and_quantities(self) -> None:
        plan = pd.concat(
            [
                build_plan_rows(order_no="SO-001", plan_customer="PLAN A"),
                build_plan_rows(order_no="SO-002", plan_customer="PLAN B"),
            ],
            ignore_index=True,
        )
        plan.loc[1, "부족수량"] = 25
        plan.loc[1, "사출생산필요수량"] = 8
        plan.loc[1, "공정재고 합계"] = 13
        raw = pd.concat(
            [
                build_order_raw(order_no="SO-001", customer="한국알콘(주)"),
                build_order_raw(order_no="SO-002", customer="Alensa s.r.o", country_code="JP", country_name="일본"),
            ],
            ignore_index=True,
        )
        master = app.build_order_status_master_from_raw(raw)

        enriched = app.apply_order_status_reference(plan, master)

        self.assertEqual(len(enriched), len(plan))
        for col in ["부족수량", "사출생산필요수량", "공정재고 합계"]:
            self.assertEqual(float(enriched[col].sum()), float(plan[col].sum()))

    def test_duplicate_master_keys_stop_merge(self) -> None:
        duplicate_master = pd.DataFrame(
            {
                "_수주번호키": ["SO-001", "SO-001"],
                app.ORDER_NO_COL: ["SO-001", "SO-001"],
                app.ORDER_CUSTOMER_COL: ["A", "B"],
            }
        )

        with self.assertRaises(ValueError):
            app.apply_order_status_reference(build_plan_rows(), duplicate_master)


if __name__ == "__main__":
    unittest.main()
