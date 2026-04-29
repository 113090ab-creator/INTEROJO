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
    xlsx_files = [p for p in base_dir.glob("*.xlsx") if not p.name.startswith("~$")]
    if len(xlsx_files) < 2:
        raise FileNotFoundError("xlsx 파일 2개(재고/수요)가 필요합니다.")

    inventory_candidates = [p for p in xlsx_files if "재고" in p.name]
    inv_path = max(inventory_candidates, key=lambda p: p.stat().st_size) if inventory_candidates else max(
        xlsx_files, key=lambda p: p.stat().st_size
    )

    demand_candidates = [p for p in xlsx_files if p != inv_path]
    demand_named = [p for p in demand_candidates if "수요" in p.name]
    if demand_named:
        demand_candidates = demand_named

    full_process_candidates = [p for p in demand_candidates if "전공정" in p.stem]
    if full_process_candidates:
        dem_path = max(full_process_candidates, key=lambda p: p.stat().st_mtime)
    else:
        dem_path = max(demand_candidates, key=lambda p: p.stat().st_mtime)

    return inv_path, dem_path


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


def extract_demand_header_info(dem_path: Path) -> tuple[dict[str, str], list[int], list[int]]:
    header_rows = pd.read_excel(dem_path, sheet_name=0, header=None, nrows=2)
    if header_rows.shape[0] < 2:
        return {}, [], []

    top_row = header_rows.iloc[0]
    second_row = header_rows.iloc[1]

    code_map: dict[str, str] = {}
    qty_col_indices: list[int] = []
    total_qty_col_indices: list[int] = []

    for idx, column_name in second_row.items():
        if "생산 수량" not in str(column_name):
            continue

        idx = int(idx)
        qty_col_indices.append(idx)

        top_label = str(top_row.iloc[idx]).strip()
        if not top_label or top_label.lower() == "nan":
            continue
        if "총합계" in top_label:
            total_qty_col_indices.append(idx)
            continue

        warehouse_name = normalize_process_to_warehouse(top_label)
        if not warehouse_name:
            continue

        match = re.search(r"\[(.*?)\]", top_label)
        extracted_code = match.group(1).strip() if match else top_label
        code_map[warehouse_name] = extracted_code

    return code_map, qty_col_indices, total_qty_col_indices


def map_demand_code_to_process_code(demand_code: str, process_prefix: str) -> str:
    code = str(demand_code).strip()
    if not code or code.lower() == "nan":
        return code

    # 수요코드(P...)를 사출/분리 재고코드(R.../Q...) 체계로 변환
    letter_pattern = re.match(r"^P(\d{4})([A-Z])(.*)$", code)
    if letter_pattern:
        return f"{process_prefix}{letter_pattern.group(1)}{letter_pattern.group(3)}"
    if code.startswith("P"):
        return f"{process_prefix}{code[1:]}"
    return code


@st.cache_data(show_spinner=False)
def load_data() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    inv_path, dem_path = find_excel_files(BASE_DIR)
    process_code_map, qty_col_indices, total_qty_col_indices = extract_demand_header_info(dem_path)

    inv = pd.read_excel(inv_path, sheet_name=0)
    dem = pd.read_excel(dem_path, sheet_name=0, header=1)

    inv.columns = [str(c).strip() for c in inv.columns]
    dem.columns = [str(c).strip() for c in dem.columns]

    if total_qty_col_indices:
        total_qty_col = dem.columns[total_qty_col_indices[-1]]
        shortage_qty = pd.to_numeric(dem[total_qty_col], errors="coerce").fillna(0)
    else:
        qty_cols = [dem.columns[i] for i in qty_col_indices]
        if not qty_cols:
            raise ValueError("수요 파일에서 '생산 수량' 컬럼을 찾지 못했습니다.")
        shortage_qty = dem[qty_cols].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1)

    inv_df = pd.DataFrame(
        {
            "품목코드": inv.iloc[:, 1].astype(str).str.strip(),
            "창고": inv.iloc[:, 5].astype(str).str.strip(),
            "재고량": pd.to_numeric(inv.iloc[:, 0], errors="coerce").fillna(0),
        }
    )
    inv_df = inv_df[(inv_df["품목코드"] != "") & (inv_df["품목코드"].str.lower() != "nan")]

    dem_df = pd.DataFrame(
        {
            "거래처": dem.iloc[:, 1].astype(str).str.strip(),
            "이니셜": dem.iloc[:, 2].astype(str).str.strip(),
            "품목코드": dem.iloc[:, 3].astype(str).str.strip(),
            "생산수량": shortage_qty,
        }
    )

    is_summary = (
        (dem_df["거래처"] == "총합계")
        | (dem_df["이니셜"] == "총합계")
        | (dem_df["품목코드"] == "총합계")
    )
    dem_df = dem_df[~is_summary]
    dem_df = dem_df[(dem_df["품목코드"] != "") & (dem_df["품목코드"].str.lower() != "nan")]

    grouped_demand = (
        dem_df.groupby(["이니셜", "거래처", "품목코드"], as_index=False)["생산수량"]
        .sum()
        .rename(columns={"생산수량": "부족수량"})
    )

    target_inv = inv_df[inv_df["창고"].isin(TARGET_WAREHOUSES)].copy()
    stock_lookup: dict[str, dict[str, float]] = {}
    for raw_name, display_name in WAREHOUSE_MAP.items():
        stock_lookup[display_name] = (
            target_inv[target_inv["창고"] == raw_name]
            .groupby("품목코드")["재고량"]
            .sum()
            .to_dict()
        )

    code_stock = pd.DataFrame({"품목코드": grouped_demand["품목코드"].drop_duplicates()})
    code_stock["사출창고"] = code_stock["품목코드"].map(
        lambda x: stock_lookup["사출창고"].get(map_demand_code_to_process_code(x, "R"), 0)
    )
    code_stock["분리창고"] = code_stock["품목코드"].map(
        lambda x: stock_lookup["분리창고"].get(map_demand_code_to_process_code(x, "Q"), 0)
    )
    code_stock["검사접착창고"] = code_stock["품목코드"].map(
        lambda x: stock_lookup["검사접착창고"].get(x, 0)
    )
    code_stock["누수규격검사 창고"] = code_stock["품목코드"].map(
        lambda x: stock_lookup["누수규격검사 창고"].get(x, 0)
    )
    code_stock["공정재고 합계"] = (
        code_stock["사출창고"]
        + code_stock["분리창고"]
        + code_stock["검사접착창고"]
        + code_stock["누수규격검사 창고"]
    )

    result = grouped_demand.merge(code_stock, on="품목코드", how="left")
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
            "재고코드 매핑 규칙": [
                "P코드 -> R코드 변환",
                "P코드 -> Q코드 변환",
                "P코드 그대로 사용",
                "P코드 그대로 사용",
            ],
            "재고>0 품목수": [
                int((code_stock["사출창고"] > 0).sum()),
                int((code_stock["분리창고"] > 0).sum()),
                int((code_stock["검사접착창고"] > 0).sum()),
                int((code_stock["누수규격검사 창고"] > 0).sum()),
            ],
        }
    )

    file_info_df = pd.DataFrame(
        {
            "재고파일": [inv_path.name],
            "수요파일": [dem_path.name],
            "행수(현황표)": [len(result)],
        }
    )

    return result, file_info_df, process_map_df


def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    st.sidebar.header("필터")

    initials = ["전체"] + sorted(df["이니셜"].dropna().unique().tolist())
    customers = ["전체"] + sorted(df["거래처"].dropna().unique().tolist())

    selected_initial = st.sidebar.selectbox("이니셜", initials, index=0)
    selected_customer = st.sidebar.selectbox("거래처", customers, index=0)
    code_query = st.sidebar.text_input("품목코드 검색", value="").strip()
    only_with_stock = st.sidebar.checkbox("공정재고 있는 항목만", value=False)

    filtered = df.copy()
    if selected_initial != "전체":
        filtered = filtered[filtered["이니셜"] == selected_initial]
    if selected_customer != "전체":
        filtered = filtered[filtered["거래처"] == selected_customer]
    if code_query:
        filtered = filtered[filtered["품목코드"].str.contains(code_query, case=False, na=False)]
    if only_with_stock:
        filtered = filtered[filtered["공정재고 합계"] > 0]

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
    st.caption("사출/분리는 수요코드(P...)를 재고코드(R.../Q...)로 변환해 매핑합니다.")
    st.dataframe(process_map_df, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
