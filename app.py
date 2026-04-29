import re
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="제품 부족수량 현황", layout="wide")

BASE_DIR = Path(__file__).resolve().parent

WAREHOUSE_MAP = {
    "사출창고": "사출창고",
    "분리창고": "분리창고",
    "검사접착": "검사접착창고",
    "누수규격검사": "누수규격검사 창고",
}
TARGET_WAREHOUSES = list(WAREHOUSE_MAP.keys())


def find_excel_files(base_dir: Path) -> tuple[Path, Path]:
    xlsx_files = [
        p
        for p in base_dir.glob("*.xlsx")
        if not p.name.startswith("~$")
    ]
    if len(xlsx_files) < 2:
        raise FileNotFoundError("xlsx 파일 2개(재고/수요)가 필요합니다.")

    # 용량이 큰 파일을 재고, 작은 파일을 수요로 가정
    xlsx_files.sort(key=lambda p: p.stat().st_size, reverse=True)
    return xlsx_files[0], xlsx_files[-1]


def normalize_process_to_warehouse(process_label: str) -> str | None:
    label = str(process_label).replace(" ", "")
    if "사출" in label:
        return "사출창고"
    if "분리" in label:
        return "분리창고"
    if "검사접착" in label or ("검사" in label and "접착" in label):
        return "검사접착창고"
    if "누수" in label or "규격검사" in label:
        return "누수규격검사 창고"
    return None


def extract_process_code_map(dem_path: Path) -> dict[str, str]:
    # 수요 시트의 1행(공정 라벨)과 2행(컬럼명)을 함께 읽어 공정코드 매핑 추출
    header_rows = pd.read_excel(dem_path, sheet_name=0, header=None, nrows=2)
    if header_rows.shape[0] < 2:
        return {}

    top_row = header_rows.iloc[0]
    second_row = header_rows.iloc[1]
    code_map: dict[str, str] = {}

    for idx, column_name in second_row.items():
        if "생산 수량" not in str(column_name):
            continue

        top_label = str(top_row.iloc[idx]).strip()
        if not top_label or top_label.lower() == "nan":
            continue

        warehouse_name = normalize_process_to_warehouse(top_label)
        if not warehouse_name:
            continue

        match = re.search(r"\[(.*?)\]", top_label)
        extracted_code = match.group(1).strip() if match else top_label
        code_map[warehouse_name] = extracted_code

    return code_map


@st.cache_data(show_spinner=False)
def load_data() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    inv_path, dem_path = find_excel_files(BASE_DIR)
    process_code_map = extract_process_code_map(dem_path)

    inv = pd.read_excel(inv_path, sheet_name=0)
    dem = pd.read_excel(dem_path, sheet_name=0, header=1)

    inv.columns = [str(c).strip() for c in inv.columns]
    dem.columns = [str(c).strip() for c in dem.columns]

    # 재고 기본 컬럼
    inv_df = pd.DataFrame(
        {
            "품목코드": inv.iloc[:, 1].astype(str).str.strip(),
            "창고": inv.iloc[:, 5].astype(str).str.strip(),
            "재고량": pd.to_numeric(inv.iloc[:, 0], errors="coerce").fillna(0),
        }
    )
    inv_df = inv_df[(inv_df["품목코드"] != "") & (inv_df["품목코드"].str.lower() != "nan")]

    # 수요 기본 컬럼
    dem_df = pd.DataFrame(
        {
            "거래처": dem.iloc[:, 1].astype(str).str.strip(),
            "이니셜": dem.iloc[:, 2].astype(str).str.strip(),
            "품목코드": dem.iloc[:, 3].astype(str).str.strip(),
            "생산수량": pd.to_numeric(dem.iloc[:, 5], errors="coerce").fillna(0),
        }
    )

    # 총합계 행 제거
    is_summary = (
        (dem_df["거래처"] == "총합계")
        | (dem_df["이니셜"] == "총합계")
        | (dem_df["품목코드"] == "총합계")
    )
    dem_df = dem_df[~is_summary]
    dem_df = dem_df[(dem_df["품목코드"] != "") & (dem_df["품목코드"].str.lower() != "nan")]

    # 생산수량 = 부족수량 (요청사항)
    grouped_demand = (
        dem_df.groupby(["이니셜", "거래처", "품목코드"], as_index=False)["생산수량"]
        .sum()
        .rename(columns={"생산수량": "부족수량"})
    )

    # 품목코드 x 공정창고 재고
    target_inv = inv_df[inv_df["창고"].isin(TARGET_WAREHOUSES)].copy()
    inv_pivot = target_inv.pivot_table(
        index="품목코드",
        columns="창고",
        values="재고량",
        aggfunc="sum",
        fill_value=0,
    ).reset_index()

    for raw_name in TARGET_WAREHOUSES:
        if raw_name not in inv_pivot.columns:
            inv_pivot[raw_name] = 0

    inv_pivot = inv_pivot.rename(columns=WAREHOUSE_MAP)
    inv_pivot["공정재고 합계"] = (
        inv_pivot["사출창고"]
        + inv_pivot["분리창고"]
        + inv_pivot["검사접착창고"]
        + inv_pivot["누수규격검사 창고"]
    )

    result = grouped_demand.merge(inv_pivot, on="품목코드", how="left")
    for col in ["사출창고", "분리창고", "검사접착창고", "누수규격검사 창고", "공정재고 합계"]:
        result[col] = result[col].fillna(0)

    process_map_df = pd.DataFrame(
        {
            "공정창고": ["사출창고", "분리창고", "검사접착창고", "누수규격검사 창고"],
            "수요정보 공정코드": [
                process_code_map.get("사출창고", "-"),
                process_code_map.get("분리창고", "-"),
                process_code_map.get("검사접착창고", "-"),
                process_code_map.get("누수규격검사 창고", "-"),
            ],
        }
    )

    return result, pd.DataFrame(
        {
            "재고파일": [inv_path.name],
            "수요파일": [dem_path.name],
            "행수(현황표)": [len(result)],
        }
    ), process_map_df


def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    st.sidebar.header("필터")

    initials = ["전체"] + sorted(df["이니셜"].dropna().unique().tolist())
    customers = ["전체"] + sorted(df["거래처"].dropna().unique().tolist())

    selected_initial = st.sidebar.selectbox("이니셜", initials, index=0)
    selected_customer = st.sidebar.selectbox("거래처", customers, index=0)
    code_query = st.sidebar.text_input("품목코드 검색", value="").strip()

    filtered = df.copy()
    if selected_initial != "전체":
        filtered = filtered[filtered["이니셜"] == selected_initial]
    if selected_customer != "전체":
        filtered = filtered[filtered["거래처"] == selected_customer]
    if code_query:
        filtered = filtered[filtered["품목코드"].str.contains(code_query, case=False, na=False)]

    return filtered


def main() -> None:
    st.title("이니셜/거래처/품목코드 기준 제품 부족수량 현황")
    st.caption("기준: 생산수량 = 부족수량")

    try:
        df, file_info, process_map_df = load_data()
    except Exception as exc:
        st.error(f"데이터 로드 실패: {exc}")
        st.stop()

    filtered = apply_filters(df)

    c1, c2, c3 = st.columns(3)
    c1.metric("현황 행 수", f"{len(filtered):,}")
    c2.metric("부족수량 합계", f"{filtered['부족수량'].sum():,.0f}")
    c3.metric("공정재고 합계", f"{filtered['공정재고 합계'].sum():,.0f}")

    st.dataframe(
        filtered[
            [
                "이니셜",
                "거래처",
                "품목코드",
                "부족수량",
                "사출창고",
                "분리창고",
                "검사접착창고",
                "누수규격검사 창고",
                "공정재고 합계",
            ]
        ].sort_values(["부족수량", "이니셜", "거래처"], ascending=[False, True, True]),
        use_container_width=True,
        height=700,
    )

    st.subheader("데이터 파일 정보")
    st.dataframe(file_info, use_container_width=True, hide_index=True)

    st.subheader("수요정보 공정코드 매핑")
    st.dataframe(process_map_df, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()

