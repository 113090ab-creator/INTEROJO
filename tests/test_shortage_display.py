import sys
import tempfile
import unittest
from pathlib import Path

from streamlit.testing.v1 import AppTest


PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))


DISPLAY_SCRIPT = """
import pandas as pd
import app

rows = []
for index, status in enumerate(
    [app.API_PRODUCT_STATUS_UNMATCHED, app.API_PRODUCT_STATUS_UNAVAILABLE, app.API_PRODUCT_STATUS_EXACT],
    start=1,
):
    row = dict.fromkeys(app.build_empty_shortage_dashboard_df().columns, "")
    row.update({
        "사이트코드": "C관",
        "거래처": "한국알콘(주)",
        "이니셜": "TEST",
        "품목코드": f"P{index}100A-01.00",
        "수요품목코드": f"T{index}100A-01.00",
        "R코드": "R0100-01.00",
        "Q코드": "Q0100-01.00",
        "제품명": f"원본 제품 {index}",
        "원본 제품명": f"원본 제품 {index}",
        "분류별요약": "1-Day_Sph",
        "API 매칭상태": status,
        "부족수량": 10 * index,
        "사출생산필요수량": index,
        "사출창고": index,
        "공정재고 합계": index,
        "납기일": "2026-10-20",
    })
    rows.append(row)

app.render_shortage_dashboard(
    pd.DataFrame(rows),
    "2026-10-15 07:33:51",
    source_label="Cloud 스냅샷 (Validated Set)",
    locked_site_filter="C관",
)
"""


class ShortageDisplayTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp_dir.cleanup)
        script_path = Path(self.temp_dir.name) / "shortage_display.py"
        script_path.write_text(DISPLAY_SCRIPT, encoding="utf-8")
        self.app_test = AppTest.from_file(script_path, default_timeout=20)

    def assert_displayed_demand(self, app_test: AppTest) -> None:
        self.assertEqual(len(app_test.exception), 0)
        table = app_test.dataframe[-1].value
        self.assertEqual(set(table["품목코드"]), {"P1100A-01.00", "P2100A-01.00", "P3100A-01.00"})
        self.assertEqual(set(table["제품명"]), {"원본 제품 1", "원본 제품 2", "원본 제품 3"})
        self.assertTrue(any("표시 3건 / 전체 3건" in caption.value for caption in app_test.caption))

    def test_production_view_keeps_demand_without_product_master_match(self) -> None:
        app_test = self.app_test.run()

        self.assert_displayed_demand(app_test)
        rendered = [element.value for element in app_test.markdown]
        self.assertTrue(any('>60<' in value for value in rendered))

    def test_common_item_view_keeps_demand_with_different_source_item_codes(self) -> None:
        self.app_test.session_state["shortage_view_selector_v4"] = "공용 품목 현황"
        app_test = self.app_test.run()

        self.assert_displayed_demand(app_test)


if __name__ == "__main__":
    unittest.main()
