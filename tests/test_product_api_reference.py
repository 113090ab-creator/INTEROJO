import sys
import unittest
from pathlib import Path
from unittest import mock

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

import app  # noqa: E402


class ProductApiReferenceTests(unittest.TestCase):
    def test_category_falls_back_to_api_name_without_replacing_product_identity(self) -> None:
        info = app.normalize_product_names_api_lookup(pd.DataFrame({
            "nm_cd": ["P3015", "S030", "S180"],
            "nm_nm": ["1-DAY58_Micure", "Clalen 1Day_30팩", "Clalen 1Day_80팩"],
            "full_gu_nm": ["1-Day_Sph"] * 3,
        }))
        demand = pd.DataFrame({
            "품목코드": ["P3015A-03.75", "P3015A-03.25"],
            "제품명": ["Clalen 1Day_30팩", "Clalen 1Day_80팩"],
        })
        result = app.apply_product_master_reference(demand, info)
        self.assertEqual(result["API 제품분류"].tolist(), ["1-Day_Sph"] * 2)
        self.assertEqual(result["분류별요약"].tolist(), ["1-Day_Sph"] * 2)
        self.assertEqual(result["제품명"].tolist(), demand["제품명"].tolist())
        self.assertEqual(result["API 제품명코드"].tolist(), ["P3015", "P3015"])
        self.assertTrue(result["제품분류 판단 근거"].eq("API full_gu_nm").all())

    def test_ambiguous_api_name_does_not_assign_category(self) -> None:
        info = app.normalize_product_names_api_lookup(pd.DataFrame({
            "nm_cd": ["S030", "A030"], "nm_nm": ["SAME", "SAME"],
            "full_gu_nm": ["1-Day_Sph", "1-Day_Toric"],
        }))
        result = app.apply_api_product_categories(pd.DataFrame({"품목코드": ["P9999A-01.00"], "제품명": ["SAME"], "API 제품분류": [0]}), info)
        self.assertEqual(result.loc[0, "API 제품분류"], "")

    def test_exact_code_category_takes_priority_over_name(self) -> None:
        info = app.normalize_product_names_api_lookup(pd.DataFrame({
            "nm_cd": ["P1234", "S030"], "nm_nm": ["CODE_NAME", "OLD_NAME"],
            "full_gu_nm": ["1-Day_Toric", "1-Day_Sph"],
        }))
        result = app.apply_api_product_categories(pd.DataFrame({"품목코드": ["P1234"], "제품명": ["OLD_NAME"]}), info)
        self.assertEqual(result.loc[0, "API 제품분류"], "1-Day_Toric")

    def test_snapshot_category_is_refreshed_from_api_and_zero_is_not_displayed(self) -> None:
        info = app.normalize_product_names_api_lookup(pd.DataFrame({
            "nm_cd": ["S180"], "nm_nm": ["Clalen 1Day_80팩"], "full_gu_nm": ["1-Day_Sph"],
        }))
        snapshot = pd.DataFrame({"품목코드": ["P3015A-03.75"], "제품명": ["Clalen 1Day_80팩"], "API 제품분류": [0]})
        with mock.patch.object(app, "load_product_names_api_lookup", return_value=info):
            result = app.enrich_loaded_shortage_snapshot_for_display(snapshot)
        self.assertEqual(result.loc[0, "API 제품분류"], "1-Day_Sph")
        display = app.format_numeric_columns_for_display(pd.DataFrame({"API 제품분류": [None, 0, "0.0", "1-Day_Sph"]}))
        self.assertEqual(display["API 제품분류"].tolist(), ["", "", "", "1-Day_Sph"])

    def test_product_api_lookup_keeps_full_nm_cd_without_first_five_dedupe(self) -> None:
        raw = pd.DataFrame(
            {
                "nm_cd": ["P1234", "P1234A"],
                "nm_nm": ["BASE", "VARIANT"],
                "full_gu_nm": ["FRP_Sph", "FRP_Toric"],
                "use_yn": ["Y", "Y"],
                "stts": ["S", "S"],
            }
        )

        lookup = app.normalize_product_names_api_lookup(raw)

        self.assertEqual(set(lookup["제품명코드"]), {"P1234", "P1234A"})
        self.assertEqual(dict(zip(lookup["제품명코드"], lookup["제품명_기준"])), {"P1234": "BASE", "P1234A": "VARIANT"})

    def test_demand_code_resolves_exact_then_variant_or_base_api_codes(self) -> None:
        code, reason = app.resolve_api_product_code_for_item("P1100", {"P1100", "P1200"})

        self.assertEqual(code, "P1100")
        self.assertIn("완전 일치", reason)

        code, reason = app.resolve_api_product_code_for_item("P1100S8.90-00.00", {"P1100", "P1200"})

        self.assertEqual(code, "P1100")
        self.assertIn("기본코드", reason)

        code, reason = app.resolve_api_product_code_for_item("P1234A-02.50", {"P1234", "P1234A"})

        self.assertEqual(code, "P1234A")
        self.assertIn("하이픈 앞", reason)

    def test_base_product_code_replaces_code_like_product_name(self) -> None:
        api_info = app.normalize_product_names_api_lookup(
            pd.DataFrame(
                {
                    "nm_cd": ["P0013"],
                    "nm_nm": ["1-Day_Contakt(5팩)"],
                    "full_gu_nm": ["1-Day_Color_Sph"],
                }
            )
        )
        demand = pd.DataFrame(
            {
                "품목코드": ["P0013A-02.50", "P0013A-02.00"],
                "제품명": ["P0013", "읽을 수 있는 변형 제품명"],
            }
        )

        enriched = app.apply_product_master_reference(demand, api_info)

        self.assertEqual(enriched.loc[0, "제품명"], "1-Day_Contakt(5팩)")
        self.assertEqual(enriched.loc[0, "API 제품명코드"], "P0013")
        self.assertEqual(enriched.loc[1, "제품명"], "읽을 수 있는 변형 제품명")
        self.assertEqual(enriched["API 매칭상태"].tolist(), [app.API_PRODUCT_STATUS_EXACT] * 2)

    def test_product_master_enrichment_marks_unmatched_and_preserves_original_name(self) -> None:
        api_info = app.normalize_product_names_api_lookup(
            pd.DataFrame(
                {
                    "nm_cd": ["P1100"],
                    "nm_nm": ["API_NAME"],
                    "full_gu_nm": ["1-Day_Sph"],
                    "use_yn": ["Y"],
                    "stts": ["S"],
                }
            )
        )
        demand = pd.DataFrame(
            {
                "품목코드": ["P1100", "P1100S8.90-00.00", "P9999-01.00"],
                "제품명": ["OLD_NAME", "VARIANT_NAME", "UNMATCHED_NAME"],
                "거래처": ["PIA Co.,Ltd.", "PIA Co.,Ltd.", "Unknown Customer"],
            }
        )

        enriched = app.apply_product_master_reference(demand, api_info)

        self.assertEqual(enriched.loc[0, "제품명"], "API_NAME")
        self.assertEqual(enriched.loc[0, "API 제품명코드"], "P1100")
        self.assertEqual(enriched.loc[0, app.PLAN_CUSTOMER_COL], "PIA Co.,Ltd.")
        self.assertEqual(enriched.loc[0, app.CUSTOMER_CONFIRMATION_STATUS_COL], app.CUSTOMER_STATUS_CONFIRMED)
        self.assertEqual(enriched.loc[0, "분류별요약"], "1-Day_Sph")
        self.assertEqual(enriched.loc[0, "제품분류 판단 근거"], "API full_gu_nm")
        self.assertEqual(enriched.loc[0, "시트분류"], "PIA 종합")
        self.assertEqual(enriched.loc[1, "제품명"], "VARIANT_NAME")
        self.assertEqual(enriched.loc[1, "API 매칭상태"], app.API_PRODUCT_STATUS_EXACT)
        self.assertEqual(enriched.loc[1, "API 매칭근거"], "품목코드 앞 5자리 기본코드가 API nm_cd와 일치")
        self.assertEqual(enriched.loc[1, "API 제품명코드"], "P1100")
        self.assertEqual(enriched.loc[2, "제품명"], "UNMATCHED_NAME")
        self.assertEqual(enriched.loc[2, "API 매칭상태"], app.API_PRODUCT_STATUS_UNMATCHED)

    def test_product_master_uses_plan_api_customer_when_product_api_has_no_customer(self) -> None:
        api_info = app.normalize_product_names_api_lookup(
            pd.DataFrame(
                {
                    "nm_cd": ["P2026"],
                    "nm_nm": ["API_PRODUCT"],
                    "full_gu_nm": ["1-Day_Sph"],
                    "use_yn": ["Y"],
                    "stts": ["S"],
                }
            )
        )
        demand = pd.DataFrame({"품목코드": ["P2026"], "제품명": ["OLD_PRODUCT"], "거래처": ["한국알콘"]})

        enriched = app.apply_product_master_reference(demand, api_info)

        self.assertEqual(enriched.loc[0, app.PLAN_CUSTOMER_COL], "한국알콘")
        self.assertEqual(enriched.loc[0, app.CUSTOMER_GROUP_COL], "Alcon")

    def test_customer_classification_uses_exact_match_only(self) -> None:
        category, reason = app.classify_sheet_with_reason(pd.Series({"거래처": "PIA unknown", "제품명": "PIA_KR TEST"}))

        self.assertEqual(category, app.UNCLASSIFIED_SHEET_CATEGORY)
        self.assertIn("완전 일치", reason)

    def test_requested_customer_exact_mappings_are_classified(self) -> None:
        expected = {
            "Alcon Services AG, Taiwan Branch": "Alcon",
            "ALCON India": "Alcon",
            "한국알콘(주)": "Alcon",
            "from-eyes Co.,ltd.": "from-eyes",
            "ESSILOR GROUP THE NETHERLAND BV": "ESSILOR",
            "OPTICAL SUPPLIES Co.": "OPTICAL SUPPLIES",
            "Coastal Contacts (Clearly)": "Coastal",
            "Alensa s.r.o": "ALENSA",
            "Hearts Optical": "HEARTS/TopTrend",
            "OPTIMAX": "MAXVUE/OPTIMAX",
        }

        for customer, group in expected.items():
            with self.subTest(customer=customer):
                category, reason = app.classify_sheet_with_reason(pd.Series({app.PLAN_CUSTOMER_COL: customer}))
                self.assertEqual(category, group)
                self.assertIn("완전 일치", reason)

    def test_blank_non_p_customer_inherits_only_from_exact_linked_p_row(self) -> None:
        work = pd.DataFrame(
            [
                {
                    "사이트코드": "A관",
                    "거래처": "ALCON India",
                    app.ORDER_NO_COL: "202609010001",
                    "접수일": "2026-09-01",
                    "이니셜": "D1926",
                    "품목코드": "P3219A-02.50ALG",
                    "제품명": "Alcon D_Allure Grey_W3C(APAC)",
                    "API납기일": "2026-08-17",
                },
                {
                    "사이트코드": "A관",
                    "거래처": "",
                    app.ORDER_NO_COL: "202609010001",
                    "접수일": "2026-09-01",
                    "이니셜": "D1926",
                    "품목코드": "Q3219-02.50ALG",
                    "제품명": "Alcon D_Allure Grey_W3C(APAC)",
                    "API납기일": "2026-08-17",
                },
                {
                    "사이트코드": "A관",
                    "거래처": "",
                    app.ORDER_NO_COL: "202609010001",
                    "접수일": "2026-09-01",
                    "이니셜": "D1926",
                    "품목코드": "Q3219-03.00ALG",
                    "제품명": "Alcon D_Allure Grey_W3C(APAC)",
                    "API납기일": "2026-08-17",
                },
                {
                    "사이트코드": "A관",
                    "거래처": "",
                    app.ORDER_NO_COL: "",
                    "접수일": "2026-09-01",
                    "이니셜": "D1926",
                    "품목코드": "R3219-02.50ALG",
                    "제품명": "Alcon D_Allure Grey_W3C(APAC)",
                    "API납기일": "2026-08-17",
                },
            ]
        )

        enriched = app.enrich_missing_customer_from_linked_p_rows(work)

        self.assertEqual(enriched.loc[1, "거래처"], "ALCON India")
        self.assertEqual(enriched.loc[1, app.CUSTOMER_FILL_SOURCE_COL], app.LINKED_P_ROW_CUSTOMER_SOURCE)
        self.assertEqual(enriched.loc[2, "거래처"], "")
        self.assertEqual(enriched.loc[3, "거래처"], "")

    def test_loaded_snapshot_display_enrichment_keeps_plan_customer_separate(self) -> None:
        snapshot = pd.DataFrame(
            {
                "품목코드": ["P2026A-01.00", "P3093A-01.00"],
                "제품명": ["OLD_A", "OLD_B"],
                "거래처": ["한국알콘(주)", ""],
                "이니셜": ["PIA_SAMPLE", "해외 SAMPLE"],
                "API 매칭상태": [app.API_PRODUCT_STATUS_EXACT, app.API_PRODUCT_STATUS_EXACT],
                "부족수량": [10, 20],
            }
        )

        with (
            mock.patch.object(app, "load_product_names_api_lookup", return_value=app.empty_product_info_lookup()),
            mock.patch.object(app, "load_order_status_master", return_value=(app.empty_order_status_master(), "")),
        ):
            enriched = app.enrich_loaded_shortage_snapshot_for_display(snapshot)

        self.assertEqual(enriched.loc[0, app.PLAN_CUSTOMER_COL], "한국알콘(주)")
        self.assertEqual(enriched.loc[0, app.CUSTOMER_GROUP_COL], "Alcon")
        self.assertEqual(enriched.loc[0, app.OPERATION_SEGMENT_COL], "PIA")
        self.assertEqual(enriched.loc[0, app.CUSTOMER_CONFIRMATION_STATUS_COL], app.CUSTOMER_STATUS_CONFIRMED)
        self.assertEqual(enriched.loc[1, app.PLAN_CUSTOMER_COL], "")
        self.assertEqual(enriched.loc[1, app.CUSTOMER_GROUP_COL], app.UNCLASSIFIED_SHEET_CATEGORY)
        self.assertEqual(enriched.loc[1, app.OPERATION_SEGMENT_COL], "기타")
        self.assertEqual(enriched.loc[1, app.CUSTOMER_CONFIRMATION_STATUS_COL], app.CUSTOMER_STATUS_SOURCE_MISSING)
        self.assertEqual(enriched.loc[1, "API 매칭상태"], app.API_PRODUCT_STATUS_EXACT)

    def test_loaded_snapshot_display_replaces_code_like_product_name(self) -> None:
        snapshot = pd.DataFrame(
            {
                "품목코드": ["P0013A-02.50"],
                "제품명": ["P0013"],
                "거래처": ["OPTIC CENTER"],
                "API 매칭상태": [app.API_PRODUCT_STATUS_UNMATCHED],
                "부족수량": [100],
            }
        )
        info = app.normalize_product_names_api_lookup(
            pd.DataFrame(
                {
                    "nm_cd": ["P0013"],
                    "nm_nm": ["1-Day_Contakt(5팩)"],
                    "full_gu_nm": ["1-Day_Color_Sph"],
                }
            )
        )

        with mock.patch.object(app, "load_product_names_api_lookup", return_value=info):
            enriched = app.enrich_loaded_shortage_snapshot_for_display(snapshot)

        self.assertEqual(enriched.loc[0, "제품명"], "1-Day_Contakt(5팩)")
        self.assertEqual(enriched.loc[0, "API 제품명코드"], "P0013")


if __name__ == "__main__":
    unittest.main()
