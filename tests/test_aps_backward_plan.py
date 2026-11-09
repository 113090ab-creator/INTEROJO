import sys
import unittest
from pathlib import Path

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

import app  # noqa: E402


class ApsBackwardPlanTests(unittest.TestCase):
    def test_backward_qty_is_the_shortage_basis_and_existing_aliases_are_populated(self) -> None:
        raw = pd.DataFrame(
            {
                "res_site_id": ["S관(3공장)"],
                "cust_name": ["고객"],
                "demand_id": ["D1"],
                "initial": ["AA"],
                "item_id": ["R1234-01.00"],
                "demand_item_id": ["T1234-01.00"],
                "due_date": ["2026-11-09"],
                "oper_id": ["[10]사출조립"],
                "backward_qty": [123],
                "forward_qty": [100],
                "diff_qty": [23],
            }
        )

        normalized = app.normalize_aps_backward_plan_dataframe(raw)

        self.assertEqual(normalized.loc[0, "backward_qty"], 123)
        self.assertEqual(normalized.loc[0, "plan_qty"], 123)
        self.assertEqual(normalized.loc[0, "demand_qty"], 123)
        self.assertEqual(normalized.loc[0, "so_id"], "D1")
        self.assertEqual(normalized.loc[0, "oper_id"], "[10]사출조립")

    def test_missing_backward_qty_is_rejected(self) -> None:
        raw = pd.DataFrame({"item_id": ["R1234"], "oper_id": ["10"], "forward_qty": [100]})

        with self.assertRaisesRegex(ValueError, "backward_qty"):
            app.normalize_aps_backward_plan_dataframe(raw)


if __name__ == "__main__":
    unittest.main()
