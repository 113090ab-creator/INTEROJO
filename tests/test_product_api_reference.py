import sys
import unittest
from pathlib import Path

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

import app  # noqa: E402


class ProductApiReferenceTests(unittest.TestCase):
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

    def test_demand_code_resolves_only_against_actual_api_codes(self) -> None:
        code, reason = app.resolve_api_product_code_for_item("P1100S8.90-00.00", {"P1100", "P1200"})

        self.assertEqual(code, "P1100")
        self.assertIn("API nm_cd", reason)

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
                "품목코드": ["P1100S8.90-00.00", "P9999-01.00"],
                "제품명": ["OLD_NAME", "UNMATCHED_NAME"],
                "거래처": ["PIA Co.,Ltd.", "Unknown Customer"],
            }
        )

        enriched = app.apply_product_master_reference(demand, api_info)

        self.assertEqual(enriched.loc[0, "제품명"], "API_NAME")
        self.assertEqual(enriched.loc[0, "API 제품명코드"], "P1100")
        self.assertEqual(enriched.loc[0, "분류별요약"], "1-Day_Sph")
        self.assertEqual(enriched.loc[0, "시트분류"], "PIA 종합")
        self.assertEqual(enriched.loc[1, "제품명"], "UNMATCHED_NAME")
        self.assertEqual(enriched.loc[1, "API 매칭상태"], app.API_PRODUCT_STATUS_UNMATCHED)

    def test_customer_classification_uses_exact_match_only(self) -> None:
        category, reason = app.classify_sheet_with_reason(pd.Series({"거래처": "PIA unknown", "제품명": "PIA_KR TEST"}))

        self.assertEqual(category, app.UNCLASSIFIED_SHEET_CATEGORY)
        self.assertIn("완전 일치", reason)


if __name__ == "__main__":
    unittest.main()
