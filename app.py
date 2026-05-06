import re
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

st.set_page_config(page_title="제품 부족수량 현황", layout="wide")

BASE_DIR = Path(__file__).resolve().parent
DISPLAY_TZ = ZoneInfo("Asia/Seoul")

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

    odv_candidates = [p for p in xlsx_files if p.stem.upper().startswith("ODV_WIP")]
    if not odv_candidates:
        odv_candidates = [p for p in xlsx_files if "ODV" in p.stem.upper() and "WIP" in p.stem.upper()]

    if odv_candidates:
        inv_path = max(odv_candidates, key=lambda p: p.stat().st_mtime)
    else:
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


def get_data_updated_at(base_dir: Path) -> str:
    try:
        inv_path, dem_path = find_excel_files(base_dir)
    except FileNotFoundError:
        return "-"

    latest_path = max([inv_path, dem_path], key=lambda p: p.stat().st_mtime)
    latest_dt = datetime.fromtimestamp(latest_path.stat().st_mtime, tz=DISPLAY_TZ)
    return latest_dt.strftime("%Y-%m-%d %H:%M:%S")


def pick_first_existing_column(columns: list[str], candidates: list[str]) -> str | None:
    for col in candidates:
        if col in columns:
            return col
    return None


def canonicalize_warehouse_label(raw_label: str) -> str:
    label = str(raw_label).strip()
    if not label or label.lower() == "nan":
        return ""

    if label in WAREHOUSE_MAP:
        return label

    display_to_key = {display: key for key, display in WAREHOUSE_MAP.items()}
    if label in display_to_key:
        return display_to_key[label]

    normalized = normalize_process_to_warehouse(label)
    if normalized is None:
        return label.replace(" ", "")

    return display_to_key.get(normalized, normalized)


def build_inventory_df(inv: pd.DataFrame) -> pd.DataFrame:
    inv.columns = [str(c).strip() for c in inv.columns]
    columns = inv.columns.tolist()

    qty_col = pick_first_existing_column(columns, ["총 재공 수량", "WIP_QTY", "재고량"])
    item_col = pick_first_existing_column(columns, ["제품 코드", "ITEM_ID", "제품코드", "품목코드"])
    warehouse_col = pick_first_existing_column(columns, ["버퍼 코드", "제품위치(창고)", "PROP02", "창고"])

    # Fallback for unknown layouts
    if qty_col is None:
        qty_col = columns[6] if len(columns) > 6 else columns[0]
    if item_col is None:
        item_col = columns[8] if len(columns) > 8 else (columns[1] if len(columns) > 1 else columns[0])
    if warehouse_col is None:
        warehouse_col = columns[10] if len(columns) > 10 else (columns[5] if len(columns) > 5 else columns[0])

    inv_df = pd.DataFrame(
        {
            "품목코드": inv[item_col].astype(str).str.strip(),
            "창고": inv[warehouse_col].astype(str).str.strip().map(canonicalize_warehouse_label),
            "재고량": pd.to_numeric(inv[qty_col], errors="coerce").fillna(0),
        }
    )

    inv_df = inv_df[(inv_df["품목코드"] != "") & (inv_df["품목코드"].str.lower() != "nan")]
    inv_df = inv_df[(inv_df["창고"] != "") & (inv_df["창고"].str.lower() != "nan")]
    return inv_df


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


def find_product_name_reference_file(base_dir: Path) -> Path | None:
    candidates = [
        p
        for p in base_dir.glob("*.xlsx")
        if not p.name.startswith("~$") and ("제품명" in p.stem and "기준" in p.stem)
    ]
    if not candidates:
        return None
    return max(candidates, key=lambda p: p.stat().st_mtime)


def load_product_reference_maps(base_dir: Path) -> tuple[dict[str, str], dict[str, str]]:
    ref_path = find_product_name_reference_file(base_dir)
    if ref_path is None:
        return {}, {}

    ref = pd.read_excel(ref_path, sheet_name=0)
    ref.columns = [str(c).strip() for c in ref.columns]

    code_col = "제품명코드" if "제품명코드" in ref.columns else ref.columns[0]
    name_col = "제품명" if "제품명" in ref.columns else ref.columns[1]
    group_col = "분류요약" if "분류요약" in ref.columns else None
    if group_col is None and "판매제품군" in ref.columns:
        group_col = "판매제품군"
    if group_col is None and "생산제품군" in ref.columns:
        group_col = "생산제품군"

    selected_cols = [code_col, name_col] + ([group_col] if group_col is not None else [])
    ref_df = ref[selected_cols].copy()
    ref_df[code_col] = ref_df[code_col].astype(str).str.strip()
    ref_df[name_col] = ref_df[name_col].astype(str).str.strip()
    if group_col is not None:
        ref_df[group_col] = ref_df[group_col].astype(str).str.strip()
    ref_df = ref_df[
        ref_df[code_col].str.startswith("P")
        & (ref_df[code_col].str.lower() != "nan")
        & (ref_df[name_col] != "")
        & (ref_df[name_col].str.lower() != "nan")
    ]
    ref_df["코드5"] = ref_df[code_col].str[:5]
    ref_df = ref_df.drop_duplicates(subset=["코드5"], keep="first")
    name_map = ref_df.set_index("코드5")[name_col].to_dict()

    if group_col is None:
        group_map: dict[str, str] = {}
    else:
        group_df = ref_df[(ref_df[group_col] != "") & (ref_df[group_col].str.lower() != "nan")]
        group_map = group_df.set_index("코드5")[group_col].to_dict()

    return name_map, group_map


def load_sheet2_group_map(base_dir: Path) -> dict[str, str]:
    ref_path = find_product_name_reference_file(base_dir)
    if ref_path is None:
        return {}

    sheet_names = pd.ExcelFile(ref_path).sheet_names
    if len(sheet_names) < 2:
        return {}

    sheet2 = pd.read_excel(ref_path, sheet_name=sheet_names[1])
    sheet2.columns = [str(c).strip() for c in sheet2.columns]
    if "코드" not in sheet2.columns or "시트이름" not in sheet2.columns:
        return {}

    df = sheet2[["코드", "시트이름"]].copy()
    df["코드"] = df["코드"].astype(str).str.strip()
    df["시트이름"] = df["시트이름"].astype(str).str.strip()
    df = df[
        df["코드"].str.startswith("P")
        & (df["코드"].str.lower() != "nan")
        & (df["시트이름"] != "")
        & (df["시트이름"].str.lower() != "nan")
    ]
    df["코드5"] = df["코드"].str[:5]
    df = df.drop_duplicates(subset=["코드5"], keep="first")
    return df.set_index("코드5")["시트이름"].to_dict()


def summarize_unique(values: pd.Series, head_count: int = 1) -> str:
    uniq = [v for v in values.astype(str).str.strip().tolist() if v and v.lower() != "nan"]
    # 순서를 유지한 unique
    uniq = list(dict.fromkeys(uniq))
    if not uniq:
        return "-"
    if len(uniq) <= head_count:
        return ", ".join(uniq)
    return f"{', '.join(uniq[:head_count])} 외 {len(uniq) - head_count}"


def format_pill_label(option: str, value_map: dict[str, float]) -> str:
    value = float(value_map.get(option, 0))
    return f"{option} ({value:,.0f})"


def build_thousand_separator_config(df: pd.DataFrame) -> dict[str, st.column_config.NumberColumn]:
    config: dict[str, st.column_config.NumberColumn] = {}
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            config[col] = st.column_config.NumberColumn(format="%,.0f")
    return config


def format_numeric_columns_for_display(df: pd.DataFrame) -> pd.DataFrame:
    display_df = df.copy()
    numeric_cols = display_df.select_dtypes(include="number").columns

    for col in numeric_cols:
        series = pd.to_numeric(display_df[col], errors="coerce")
        non_null = series.dropna()
        is_integer_like = non_null.empty or ((non_null % 1) == 0).all()

        if is_integer_like:
            display_df[col] = series.map(lambda x: "" if pd.isna(x) else f"{x:,.0f}")
        else:
            display_df[col] = series.map(lambda x: "" if pd.isna(x) else f"{x:,.2f}")

    return display_df


def split_query_terms(query: str) -> list[str]:
    return [term.strip() for term in str(query).split(",") if term.strip()]


def filter_with_terms(df: pd.DataFrame, column: str, query: str) -> pd.DataFrame:
    terms = split_query_terms(query)
    if not terms:
        return df
    pattern = "|".join(re.escape(term) for term in terms)
    return df[df[column].astype(str).str.contains(pattern, case=False, na=False)]


def build_qcode_summary(df: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "Q코드",
        "Q기준 제품명",
        "파워",
        "대표 이니셜",
        "대표 P코드",
        "부족수량 합계",
        "분리창고",
        "사출창고",
        "공정재고 합계",
    ]
    if df.empty:
        return pd.DataFrame(columns=columns)

    q_df = df.copy()
    q_df["Q코드"] = q_df["품목코드"].map(lambda x: map_demand_code_to_process_code(x, "Q"))
    q_df["파워"] = q_df["Q코드"].map(extract_power_from_code)

    summary = (
        q_df.groupby(["Q코드", "파워"], as_index=False)
        .agg(
            {
                "제품명": lambda s: summarize_unique(s, head_count=1),
                "이니셜": lambda s: summarize_unique(s, head_count=1),
                "품목코드": lambda s: summarize_unique(s, head_count=1),
                "부족수량": "sum",
                "분리창고": "max",
                "사출창고": "max",
                "공정재고 합계": "max",
            }
        )
        .rename(
            columns={
                "제품명": "Q기준 제품명",
                "이니셜": "대표 이니셜",
                "품목코드": "대표 P코드",
                "부족수량": "부족수량 합계",
            }
        )
        .sort_values(["부족수량 합계", "Q코드"], ascending=[False, True])
    )
    return summary[columns]


@st.cache_data(show_spinner=False)
def load_data() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    inv_path, dem_path = find_excel_files(BASE_DIR)
    product_name_map, product_group_map = load_product_reference_maps(BASE_DIR)
    sheet2_group_map = load_sheet2_group_map(BASE_DIR)
    process_code_map, warehouse_qty_col_indices, qty_col_indices, total_qty_col_indices = extract_demand_header_info(dem_path)

    inv = pd.read_excel(inv_path, sheet_name=0)
    dem = pd.read_excel(dem_path, sheet_name=0, header=1)

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

    inv_df = build_inventory_df(inv)

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
    grouped_demand["코드5"] = grouped_demand["품목코드"].str[:5]
    grouped_demand["제품명"] = grouped_demand["코드5"].map(product_name_map).fillna(grouped_demand["제품명"])
    grouped_demand["제품명"] = grouped_demand["제품명"].replace({"": "-", "nan": "-", "None": "-"}).fillna("-")
    grouped_demand["분류별요약"] = grouped_demand["코드5"].map(product_group_map).fillna("기타")
    grouped_demand["시트분류"] = grouped_demand["코드5"].map(sheet2_group_map).fillna("기타 해외")
    grouped_demand = grouped_demand.drop(columns=["코드5"])

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


def apply_filters(df: pd.DataFrame, updated_at: str) -> pd.DataFrame:
    st.subheader("필터")
    st.caption(f"업데이트: {updated_at}")

    r1c1, r1c2, r1c3, r1c4 = st.columns([1.0, 1.2, 1.8, 0.9])
    with r1c1:
        initial_query = st.text_input(
            "이니셜 검색",
            value="",
            key="flt_initial_query",
            placeholder="예: PIA, MON",
        ).strip()
    with r1c2:
        customer_query = st.text_input(
            "거래처 검색",
            value="",
            key="flt_customer_query",
            placeholder="예: 국내, 중국",
        ).strip()
    with r1c3:
        code_query = st.text_input(
            "품목코드 검색",
            value="",
            key="flt_code",
            placeholder="예: P1234, P5678",
        ).strip()
    with r1c4:
        only_with_stock = st.checkbox("공정재고만", value=False, key="flt_only_stock")
        exclude_safe_initial = st.checkbox("안전 이니셜 제외", value=False, key="flt_exclude_safe_initial")

    sheet_sum_map = (
        df.groupby("시트분류", as_index=True)["부족수량"].sum().sort_values(ascending=False).to_dict()
    )
    summary_sum_map = (
        df.groupby("분류별요약", as_index=True)["부족수량"].sum().sort_values(ascending=False).to_dict()
    )

    sheet_options = ["전체"] + list(sheet_sum_map.keys())
    summary_options = ["전체"] + list(summary_sum_map.keys())
    sheet_count_map = {"전체": float(df["부족수량"].sum()), **sheet_sum_map}
    summary_count_map = {"전체": float(df["부족수량"].sum()), **summary_sum_map}

    selected_sheet_option = st.pills(
        "시트 분류",
        options=sheet_options,
        default="전체",
        key="flt_sheet_pills",
        format_func=lambda x: format_pill_label(x, sheet_count_map),
    )
    selected_summary_option = st.pills(
        "분류별 요약",
        options=summary_options,
        default="전체",
        key="flt_summary_pills",
        format_func=lambda x: format_pill_label(x, summary_count_map),
    )

    filtered = df.copy()
    filtered = filter_with_terms(filtered, "이니셜", initial_query)
    filtered = filter_with_terms(filtered, "거래처", customer_query)
    if exclude_safe_initial:
        filtered = filtered[~filtered["이니셜"].astype(str).str.contains("안전", na=False)]
    if selected_sheet_option and selected_sheet_option != "전체":
        filtered = filtered[filtered["시트분류"] == selected_sheet_option]
    if selected_summary_option and selected_summary_option != "전체":
        filtered = filtered[filtered["분류별요약"] == selected_summary_option]
    filtered = filter_with_terms(filtered, "품목코드", code_query)
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

    updated_at = get_data_updated_at(BASE_DIR)
    filtered = apply_filters(df, updated_at)
    q_summary = build_qcode_summary(filtered)

    tab_p, tab_q = st.tabs(["P코드 기준 현황", "Q코드 기준 집계"])

    with tab_p:
        c1, c2, c3 = st.columns(3)
        c1.metric("현황 행 수", f"{len(filtered):,}")
        c2.metric("부족수량 합계", f"{filtered['부족수량'].sum():,.0f}")
        c3.metric("공정재고 합계", f"{filtered['공정재고 합계'].sum():,.0f}")

        p_table = filtered[
            [
                "거래처",
                "이니셜",
                "품목코드",
                "제품명",
                "분류별요약",
                "시트분류",
                "파워",
                "납기일",
                "부족수량",
                "사출창고",
                "분리창고",
                "검사접착창고",
                "누수규격검사 창고",
                "공정재고 합계",
            ]
        ].sort_values(["부족수량", "이니셜", "거래처"], ascending=[False, True, True])
        p_table_display = format_numeric_columns_for_display(p_table)

        st.dataframe(
            p_table_display,
            use_container_width=True,
            height=700,
        )

    with tab_q:
        q1, q2, q3 = st.columns(3)
        q1.metric("Q코드 수", f"{len(q_summary):,}")
        q2.metric("Q기준 부족수량 합계", f"{q_summary['부족수량 합계'].sum():,.0f}")
        q3.metric("Q기준 공정재고 합계", f"{q_summary['공정재고 합계'].sum():,.0f}")

        q_table = q_summary.copy()
        q_table_display = format_numeric_columns_for_display(q_table)
        st.dataframe(
            q_table_display,
            use_container_width=True,
            height=700,
        )

    # 요청사항: 파일/매핑 상세 정보는 화면에서 숨김


if __name__ == "__main__":
    main()
