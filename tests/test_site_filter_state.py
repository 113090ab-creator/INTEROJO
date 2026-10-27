import sys
import unittest
from pathlib import Path
from unittest import mock

from streamlit.elements.lib import policies
from streamlit.testing.v1 import AppTest


PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))


FILTER_SCRIPT = """
import pandas as pd
import streamlit as st
import app

rows = []
for index, site in enumerate(st.session_state.get("test_sites", ["A관", "C관"]), start=1):
    row = dict.fromkeys(app.build_empty_shortage_dashboard_df().columns, "")
    row.update({
        "사이트코드": site,
        "품목코드": f"P{index}100A-01.00",
        "이니셜": "TEST",
        "거래처": "테스트 거래처",
        "분류별요약": "1-Day_Sph",
        "부족수량": 10 * index,
    })
    rows.append(row)

filtered = app.apply_filters(pd.DataFrame(rows), "2026-10-27 07:26:12")
st.dataframe(filtered[["사이트코드", "품목코드"]])
"""


class SiteFilterStateTests(unittest.TestCase):
    def setUp(self) -> None:
        # Streamlit emits this warning only once per process. Reset that guard so
        # an earlier widget test cannot hide a duplicate initialization here.
        warning_guard = mock.patch.object(policies, "_shown_default_value_warning", False)
        warning_guard.start()
        self.addCleanup(warning_guard.stop)
        self.app_test = AppTest.from_string(FILTER_SCRIPT, default_timeout=20)

    def site_pills(self):
        return next(widget for widget in self.app_test.get("button_group") if widget.key == "flt_site_pills")

    def run_app(self, site: str | None = None) -> None:
        # Streamlit 1.53 AppTest serializes every ButtonGroup as a list, even
        # though single-selection pills return a scalar to the app.
        for widget in self.app_test.get("button_group"):
            if widget.key == "flt_site_pills":
                widget.set_value([site if site is not None else widget.value])
        self.app_test.run()

    def assert_filtered_sites(self, expected_sites: set[str]) -> None:
        self.assertEqual(len(self.app_test.exception), 0)
        self.assertEqual([warning.value for warning in self.app_test.warning], [])
        self.assertEqual(set(self.app_test.dataframe[-1].value["사이트코드"]), expected_sites)

    def test_first_load_selects_c_site_without_duplicate_default_warning(self) -> None:
        self.run_app()

        self.assertEqual(self.site_pills().value, "C관")
        self.assert_filtered_sites({"C관"})

    def test_user_can_change_site_and_keep_it_on_rerun(self) -> None:
        self.run_app()
        self.run_app("A관")
        self.run_app()

        self.assertEqual(self.site_pills().value, "A관")
        self.assert_filtered_sites({"A관"})

    def test_user_can_select_all_without_being_reset_to_c_site(self) -> None:
        self.run_app()
        self.run_app("전체")
        self.run_app()

        self.assertEqual(self.site_pills().value, "전체")
        self.assert_filtered_sites({"A관", "C관"})

    def test_removed_site_recovers_to_valid_default_without_warning(self) -> None:
        self.run_app()
        self.run_app("A관")
        self.app_test.session_state["test_sites"] = ["C관"]
        self.run_app()

        self.assertEqual(self.site_pills().value, "C관")
        self.assert_filtered_sites({"C관"})

    def test_no_c_site_falls_back_to_all_without_warning(self) -> None:
        self.app_test.session_state["test_sites"] = ["A관"]
        self.run_app()

        self.assertEqual(self.site_pills().value, "전체")
        self.assert_filtered_sites({"A관"})

    def test_legacy_all_selection_migrates_to_c_site_without_warning(self) -> None:
        self.app_test.session_state["flt_site_pills"] = "전체"
        self.run_app()

        self.assertEqual(self.site_pills().value, "C관")
        self.assert_filtered_sites({"C관"})


if __name__ == "__main__":
    unittest.main()
