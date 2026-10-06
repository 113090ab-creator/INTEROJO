import sys
import unittest
from pathlib import Path

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

import app  # noqa: E402


class ShortageExcelDownloadTests(unittest.TestCase):
    def test_download_frame_keeps_requested_columns_only(self) -> None:
        source = pd.DataFrame(
            {
                "거래처": ["PIA Co.,Ltd."],
                "이니셜": ["PIA_SAMPLE"],
                "생산코드": ["P1234A-01.00"],
                "사출코드": ["R1234"],
                "분리코드": ["Q1234"],
                "제품명": ["원본명"],
                "제품명코드": ["P1234"],
                "분류별요약": ["1-Day_Sph"],
                app.PIA_ORDER_CLASS_COL: ["신제품 / 8월 접수"],
                app.ORDER_RECEIVED_DATE_COL: ["2026-08-01"],
                "파워": ["-1.00"],
                "납기일": ["2026-08-15"],
                "생산부족수량": [10],
                "사출부족수량": [4],
                "사출재고": [1],
                "분리재고": [2],
                "검사접착재고": [3],
                "검사접착재작업창고": [4],
                "누수규격검사": [5],
                "공정재고합계": [15],
                "비고": [" 확인 필요 "],
                "내부 확인용": ["다운로드 제외"],
            }
        )

        result = app.build_shortage_excel_download_frame(source)

        self.assertEqual(result.columns.tolist(), app.SHORTAGE_EXCEL_DOWNLOAD_COLUMNS)
        self.assertEqual(result.loc[0, "품목코드"], "P1234A-01.00")
        self.assertEqual(result.loc[0, "R코드"], "R1234")
        self.assertEqual(result.loc[0, "Q코드"], "Q1234")
        self.assertEqual(result.loc[0, "원본 제품명"], "원본명")
        self.assertEqual(result.loc[0, "API 제품명코드"], "P1234")
        self.assertEqual(result.loc[0, "API 제품분류"], "1-Day_Sph")
        self.assertEqual(result.loc[0, "부족수량"], 10)
        self.assertEqual(result.loc[0, "사출 부족수량"], 4)
        self.assertEqual(result.loc[0, "누수규격검사 창고"], 5)
        self.assertEqual(result.loc[0, "공정재고 합계"], 15)
        self.assertNotIn("내부 확인용", result.columns)

    def test_download_frame_adds_missing_requested_columns(self) -> None:
        result = app.build_shortage_excel_download_frame(pd.DataFrame({"거래처": ["A"]}))

        self.assertEqual(result.columns.tolist(), app.SHORTAGE_EXCEL_DOWNLOAD_COLUMNS)
        self.assertEqual(result.loc[0, "거래처"], "A")
        self.assertEqual(result.loc[0, "품목코드"], "")
        self.assertEqual(result.loc[0, "비고"], "")


if __name__ == "__main__":
    unittest.main()
