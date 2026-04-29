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


def extract_demand_header_info(dem_path: Path) -> tuple[dict[str, str], dict[str, int], list[int], list[int]]:
    header_rows = pd.read_excel(dem_path, sheet_name=0, header=None, nrows=2)
    if header_rows.shape[0] < 2:
        return {}, {}, [], []

    top_row = header_rows.iloc[0]
    second_row = header_rows.iloc[1]

    code_map: dict[str, str] = {}
    warehouse_qty_col_indices: dict[str, int] = {}
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
        warehouse_qty_col_indices[warehouse_name] = idx

    return code_map, warehouse_qty_col_indices, qty_col_indices, total_qty_col_indices


def map_demand_code_to_process_code(demand_code: str, process_prefix: str) -> str:
    code = str(demand_code).strip()
    if not code or code.lower() == "nan":
        return code

    letter_pattern = re.match(r"^P(\d{4})([A-Z])(.*)$", code)
    if letter_pattern:
        return f"{process_prefix}{letter_pattern.group(1)}{letter_pattern.group(3)}"
    if code.startswith("P"):
        return f"{process_prefix}{code[1:]}"
    return code


def extract_power_from_code(item_code: str) -> str:
    code = str(item_code).strip()
    match = re.search(r"([+-]\d{1,2}\.\d{2})", code)
    return match.group(1) if match else "-"


@st.cache_data(show_spinner=False)
def load_data() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    inv_path, dem_path = find_excel_files(BASE_DIR)
    process_code_map, warehouse_qty_col_indices, qty_col_indices, total_qty_col_indices = extract_demand_header_info(dem_path)

    inv = pd.read_excel(inv_path, sheet_name=0)
    dem = pd.read_excel(dem_path, sheet_name=0, header=1)

    inv.columns = [str(c).strip() for c in inv.columns]
    dem.columns = [str(c).strip() for c in dem.columns]

    leak_qty_idx = warehouse_qty_col_indices.get("누수규격검사 창고")
    leak_due_idx = leak_qty_idx + 1 if leak_qty_idx is not None and (leak_qty_idx + 1) < dem.shape[1] else None
    if leak_qty_idx is not None:
        shortage_qty = pd.to_numeric(dem.iloc[:, leak_qty_idx], errors="coerce").fillna(0)
    elif total_qty_col_indices:
        total_qty_col = dem.columns[total_qty_col_indices[-1]]
        shortage_qty = pd.to_numeric(dem[total_qty_col], errors="coerce").fillna(0)
    else:
        qty_cols = [dem.columns[i] for i in qty_col_indices]
        if not qty_cols:
            raise ValueError("수요 파일에서 '생산 수량' 컬럼을 찾지 못했습니다.")
        shortage_qty = dem[qty_cols].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1)

    if leak_due_idx is not None:
        leak_due_date = pd.to_datetime(dem.iloc[:, leak_due_idx], errors="coerce")
    else:
        leak_due_date = pd.Series(pd.NaT, index=dem.index, dtype="datetime64[ns]")

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
            "제품명": dem.iloc[:, 4].astype(str).str.strip(),
            "납기일": leak_due_date,
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
    dem_df = dem_df[dem_df["품목코드"].str.startswith("P")]
    dem_df = dem_df[dem_df["생산수량"] > 0]
    dem_df["제품명"] = dem_df["제품명"].replace({"nan": "", "None": ""})

    grouped_demand = (
        dem_df.groupby(["이니셜", "거래처", "품목코드"], as_index=False)
        .agg(
            {
                "생산수량": "sum",
                "납기일": "min",
                "제품명": lambda s: next((v for v in s if str(v).strip() and str(v).strip().lower() != "nan"), "-"),
            }
        )
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
    result["파워"] = result["품목코드"].map(extract_power_from_code)
    result["납기일"] = pd.to_datetime(result["납기일"], errors="coerce").dt.strftime("%Y-%m-%d")
    result["납기일"] = result["납기일"].fillna("-")

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
    st.subheader("필터")
    f1, f2, f3, f4 = st.columns([1.2, 1.4, 1.6, 1.0])

    initials = ["전체"] + sorted(df["이니셜"].dropna().unique().tolist())
    customers = ["전체"] + sorted(df["거래처"].dropna().unique().tolist())

    with f1:
        selected_initial = st.selectbox("이니셜", initials, index=0, key="flt_initial")
    with f2:
        selected_customer = st.selectbox("거래처", customers, index=0, key="flt_customer")
    with f3:
        code_query = st.text_input("품목코드 검색", value="", key="flt_code").strip()
    with f4:
        only_with_stock = st.checkbox("공정재고만", value=False, key="flt_only_stock")

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
    st.caption("기준: 누수규격검사 생산수량 = 부족수량, 품목코드는 P코드만 표시")

    try:
        df, _, _ = load_data()
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
                "거래처",
                "이니셜",
                "품목코드",
                "제품명",
                "파워",
                "납기일",
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

    # 요청사항: 파일/매핑 상세 정보는 화면에서 숨김


if __name__ == "__main__":
    main()
