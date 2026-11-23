import sys
import unittest
from pathlib import Path

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

import app  # noqa: E402


def build_row(
    *,
    item_code: str,
    customer: str,
    initial: str,
    order_no: str,
    power: str,
    product_name: str,
    injection_qty: float = 0,
) -> dict[str, object]:
    return {
        app.ORDER_NO_COL: order_no,
        "거래처": customer,
        "이니셜": initial,
        "품목코드": item_code,
        "제품명": product_name,
        "파워": power,
        "R코드": "R0108-04.75SOB",
        "Q코드": "Q0108-04.75SOB",
        "사출생산필요수량": injection_qty,
        app.SEPARATION_REQUIRED_QTY_COL: 0,
    }


class SyntheticPRowTests(unittest.TestCase):
    def test_does_not_borrow_p0108_from_another_initial(self) -> None:
        source = pd.DataFrame(
            [
                build_row(
                    item_code="R0108-04.75SOB",
                    customer="",
                    initial="국내",
                    order_no="국내_ORDER_1",
                    power="-04.75",
                    product_name="PIA_KEOPI BROWN",
                    injection_qty=29497,
                )
            ]
        )
        p_reference = pd.DataFrame(
            [
                build_row(
                    item_code="P0108A-04.75SOB",
                    customer="PIA Co.,Ltd.",
                    initial="AA1674",
                    order_no="PIA_ORDER_1",
                    power="-04.75",
                    product_name="PIA_KEOPI BROWN",
                )
            ]
        )

        result = app.build_synthetic_p_rows_for_process_scope(
            source,
            p_reference,
            source.columns.tolist(),
        )

        self.assertTrue(result.empty)

    def test_uses_same_initial_parent_when_process_only_row_needs_display_p(self) -> None:
        source = pd.DataFrame(
            [
                build_row(
                    item_code="R0108-04.75SOB",
                    customer="국내",
                    initial="국내",
                    order_no="국내_ORDER_2",
                    power="-04.75",
                    product_name="1-Day_iris_SoulBrown",
                    injection_qty=29497,
                )
            ]
        )
        p_reference = pd.DataFrame(
            [
                build_row(
                    item_code="P0010A-04.75SOB",
                    customer="국내",
                    initial="국내",
                    order_no="국내_ORDER_1",
                    power="-04.75",
                    product_name="1-Day_iris_SoulBrown",
                )
            ]
        )

        result = app.build_synthetic_p_rows_for_process_scope(
            source,
            p_reference,
            source.columns.tolist(),
        )

        self.assertEqual(result["품목코드"].tolist(), ["P0010A-04.75SOB"])
        self.assertEqual(result["제품명"].tolist(), ["1-Day_iris_SoulBrown"])


if __name__ == "__main__":
    unittest.main()
