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

    paths = [inv_path, dem_path]
    ref_path = find_product_name_reference_file(base_dir)
    if ref_path is not None:
        paths.append(ref_path)

    latest_path = max(paths, key=lambda p: p.stat().st_mtime)
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


def load_rq_code_maps(base_dir: Path) -> tuple[dict[str, str], dict[str, str], dict[str, str]]:
    ref_path = find_product_name_reference_file(base_dir)
    if ref_path is None:
        return {}, {}, {}

    sheet_names = pd.ExcelFile(ref_path).sheet_names
    if len(sheet_names) < 2:
        return {}, {}, {}

    sheet2 = pd.read_excel(ref_path, sheet_name=sheet_names[1])
    sheet2.columns = [str(c).strip() for c in sheet2.columns]
    required = {"코드", "Q코드", "R코드"}
    if not required.issubset(sheet2.columns):
        return {}, {}, {}

    name_col = "제품명" if "제품명" in sheet2.columns else None

    df = sheet2.copy()
    for col in ["코드", "Q코드", "R코드"]:
        df[col] = df[col].astype(str).str.strip()
    if name_col:
        df[name_col] = df[name_col].astype(str).str.strip()

    df = df[
        df["코드"].str.startswith("P")
        & (df["코드"].str.lower() != "nan")
        & (df["Q코드"].str.lower() != "nan")
        & (df["R코드"].str.lower() != "nan")
        & (df["Q코드"] != "")
        & (df["R코드"] != "")
    ]

    df["코드5"] = df["코드"].str[:5]
    df = df.drop_duplicates(subset=["코드5"], keep="first")

    q_map = df.set_index("코드5")["Q코드"].to_dict()
    r_map = df.set_index("코드5")["R코드"].to_dict()
    if name_col:
        r_name_map = df.set_index("코드5")[name_col].to_dict()
    else:
        r_name_map = {}
    return r_map, q_map, r_name_map


def build_data_refresh_key(base_dir: Path) -> str:
    inv_path, dem_path = find_excel_files(base_dir)
    ref_path = find_product_name_reference_file(base_dir)

    paths = [inv_path, dem_path]
    if ref_path is not None:
        paths.append(ref_path)

    parts = []
    for p in paths:
        stat = p.stat()
        parts.append(f"{p.name}:{stat.st_size}:{stat.st_mtime_ns}")
    return "|".join(parts)


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


def filter_with_terms_any(df: pd.DataFrame, columns: list[str], query: str) -> pd.DataFrame:
    terms = split_query_terms(query)
    if not terms:
        return df

    pattern = "|".join(re.escape(term) for term in terms)
    mask = pd.Series(False, index=df.index)
    for col in columns:
        mask = mask | df[col].astype(str).str.contains(pattern, case=False, na=False)
    return df[mask]


def add_rq_group_columns(df: pd.DataFrame) -> pd.DataFrame:
    enriched = df.copy()
    if "R코드" not in enriched.columns:
        enriched["R코드"] = enriched["품목코드"].map(lambda x: map_demand_code_to_process_code(x, "R"))
    if "Q코드" not in enriched.columns:
        enriched["Q코드"] = enriched["품목코드"].map(lambda x: map_demand_code_to_process_code(x, "Q"))
    if "R코드 제품명" not in enriched.columns:
        enriched["R코드 제품명"] = enriched.get("제품명", "-")
    enriched["RQ그룹"] = enriched["R코드"].astype(str) + " | " + enriched["Q코드"].astype(str)
    return enriched


def build_rq_group_summary(df: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "R코드",
        "Q코드",
        "R코드 제품명",
        "P코드 수",
        "제품명 예시",
        "P코드 예시",
        "부족수량 합계",
        "사출창고 합계",
        "분리창고 합계",
        "공정재고 합계",
        "사출 부족수량",
    ]
    if df.empty:
        return pd.DataFrame(columns=columns)

    grouped = (
        df.groupby(["R코드", "Q코드"], as_index=False)
        .agg(
            {
                "R코드 제품명": lambda s: summarize_unique(s, head_count=1),
                "제품명": lambda s: summarize_unique(s, head_count=3),
                "품목코드": lambda s: summarize_unique(s, head_count=5),
                "부족수량": "sum",
                "사출창고": "sum",
                "분리창고": "sum",
                "공정재고 합계": "sum",
            }
        )
        .rename(
            columns={
                "제품명": "제품명 예시",
                "품목코드": "P코드 예시",
                "부족수량": "부족수량 합계",
                "사출창고": "사출창고 합계",
                "분리창고": "분리창고 합계",
            }
        )
    )
    p_count = df.groupby(["R코드", "Q코드"])["품목코드"].nunique().rename("P코드 수").reset_index()
    grouped = grouped.merge(p_count, on=["R코드", "Q코드"], how="left")
    grouped["사출 부족수량"] = grouped["부족수량 합계"] - grouped["사출창고 합계"]
    grouped = grouped.sort_values(["부족수량 합계", "P코드 수"], ascending=[False, False])
    return grouped[columns]


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
def load_data(refresh_key: str) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    _ = refresh_key
    inv_path, dem_path = find_excel_files(BASE_DIR)
    product_name_map, product_group_map = load_product_reference_maps(BASE_DIR)
    sheet2_group_map = load_sheet2_group_map(BASE_DIR)
    r_code_map, q_code_map, r_name_map = load_rq_code_maps(BASE_DIR)
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
    grouped_demand["R코드"] = grouped_demand["코드5"].map(r_code_map)
    grouped_demand["Q코드"] = grouped_demand["코드5"].map(q_code_map)
    grouped_demand["R코드 제품명"] = grouped_demand["코드5"].map(r_name_map)
    grouped_demand["R코드"] = grouped_demand["R코드"].fillna(grouped_demand["품목코드"].map(lambda x: map_demand_code_to_process_code(x, "R")))
    grouped_demand["Q코드"] = grouped_demand["Q코드"].fillna(grouped_demand["품목코드"].map(lambda x: map_demand_code_to_process_code(x, "Q")))
    grouped_demand["R코드 제품명"] = grouped_demand["R코드 제품명"].fillna(grouped_demand["제품명"])
    grouped_demand["R코드 제품명"] = grouped_demand["R코드 제품명"].replace({"": "-", "nan": "-", "None": "-"}).fillna("-")
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

    r1c1, r1c2 = st.columns([3.0, 1.2])
    with r1c1:
        unified_query = st.text_input(
            "통합 검색 (이니셜/거래처/품목코드/제품명/R코드/Q코드)",
            value="",
            key="flt_unified_query",
            placeholder="예: PIA, 국내, P1234",
            help="콤마(,)로 여러 키워드를 입력하면 OR 조건으로 검색합니다.",
        ).strip()
    with r1c2:
        only_with_stock = st.checkbox("공정재고만", value=False, key="flt_only_stock")
        exclude_safe_initial = st.checkbox("안전 이니셜 제외", value=False, key="flt_exclude_safe_initial")
        only_same_rq_group = st.checkbox("동일 RQ그룹만", value=False, key="flt_only_same_rq_group")

    product_query = st.text_input(
        "R코드 제품명 검색",
        value="",
        key="flt_product_query",
        placeholder="예: 1-Day_58, Bella, Chai Cafe",
        help="콤마(,)로 여러 R코드 제품명을 입력하면 OR 조건으로 검색합니다.",
    ).strip()

    sheet_sum_map = (
        df.groupby("시트분류", as_index=True)["부족수량"].sum().sort_values(ascending=False).to_dict()
    )
    summary_sum_map = (
        df.groupby("분류별요약", as_index=True)["부족수량"].sum().sort_values(ascending=False).to_dict()
    )
    product_sum_map = (
        df.groupby("R코드 제품명", as_index=True)["부족수량"].sum().sort_values(ascending=False).to_dict()
    )

    sheet_options = ["전체"] + list(sheet_sum_map.keys())
    summary_options = ["전체"] + list(summary_sum_map.keys())
    product_top_n = 20
    product_options = ["전체"] + [p for p in list(product_sum_map.keys()) if str(p).strip() not in {"", "-", "nan", "None"}][:product_top_n]
    sheet_count_map = {"전체": float(df["부족수량"].sum()), **sheet_sum_map}
    summary_count_map = {"전체": float(df["부족수량"].sum()), **summary_sum_map}
    product_count_map = {"전체": float(df["부족수량"].sum())}
    for product_name in product_options[1:]:
        product_count_map[product_name] = float(product_sum_map.get(product_name, 0))

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
    selected_product_option = st.pills(
        "R코드 제품명 (상위)",
        options=product_options,
        default="전체",
        key="flt_product_pills",
        format_func=lambda x: format_pill_label(x, product_count_map),
    )

    rq_option_map: dict[str, tuple[str, str]] = {}
    if {"R코드", "Q코드", "품목코드", "부족수량"}.issubset(df.columns):
        rq_meta = (
            df.groupby(["R코드", "Q코드"], as_index=False)
            .agg({"품목코드": "nunique", "부족수량": "sum"})
            .rename(columns={"품목코드": "p_count", "부족수량": "sum_shortage"})
            .sort_values(["p_count", "sum_shortage"], ascending=[False, False])
        )
        rq_options = ["전체"]
        for _, row in rq_meta.iterrows():
            r_code = str(row["R코드"])
            q_code = str(row["Q코드"])
            p_count = int(row["p_count"])
            shortage = float(row["sum_shortage"])
            label = f"{r_code} | {q_code} (P:{p_count}, 부족:{shortage:,.0f})"
            rq_option_map[label] = (r_code, q_code)
            rq_options.append(label)
        selected_rq_option = st.selectbox("RQ 그룹 선택", options=rq_options, index=0, key="flt_rq_group")
    else:
        selected_rq_option = "전체"

    filtered = df.copy()
    search_cols = [c for c in ["이니셜", "거래처", "품목코드", "제품명", "R코드 제품명", "R코드", "Q코드"] if c in filtered.columns]
    filtered = filter_with_terms_any(filtered, search_cols, unified_query)
    filtered = filter_with_terms(filtered, "R코드 제품명", product_query)
    if exclude_safe_initial:
        filtered = filtered[~filtered["이니셜"].astype(str).str.contains("안전", na=False)]
    if selected_sheet_option and selected_sheet_option != "전체":
        filtered = filtered[filtered["시트분류"] == selected_sheet_option]
    if selected_summary_option and selected_summary_option != "전체":
        filtered = filtered[filtered["분류별요약"] == selected_summary_option]
    if selected_product_option and selected_product_option != "전체":
        filtered = filtered[filtered["R코드 제품명"] == selected_product_option]
    if selected_rq_option != "전체" and selected_rq_option in rq_option_map:
        r_code, q_code = rq_option_map[selected_rq_option]
        filtered = filtered[(filtered["R코드"].astype(str) == r_code) & (filtered["Q코드"].astype(str) == q_code)]
    if only_same_rq_group and {"R코드", "Q코드", "품목코드"}.issubset(filtered.columns):
        p_count_per_group = filtered.groupby(["R코드", "Q코드"])["품목코드"].transform("nunique")
        filtered = filtered[p_count_per_group >= 2]
    if only_with_stock:
        filtered = filtered[filtered["공정재고 합계"] > 0]

    return filtered


@st.cache_data(show_spinner=False)
def load_leadji_data(refresh_key: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    _ = refresh_key
    ref_path = find_product_name_reference_file(BASE_DIR)
    if ref_path is None:
        return pd.DataFrame(), pd.DataFrame()

    sheet_names = pd.ExcelFile(ref_path).sheet_names
    leadji_info_sheet = next((s for s in sheet_names if s.replace(" ", "") == "리드지정보"), None)
    leadji_stock_sheet = next((s for s in sheet_names if s.replace(" ", "") == "리드지재고"), None)

    leadji_info = pd.read_excel(ref_path, sheet_name=leadji_info_sheet) if leadji_info_sheet else pd.DataFrame()
    leadji_stock = pd.read_excel(ref_path, sheet_name=leadji_stock_sheet) if leadji_stock_sheet else pd.DataFrame()

    if not leadji_info.empty:
        leadji_info.columns = [str(c).strip() for c in leadji_info.columns]
        for col in leadji_info.columns:
            if "소요량" in col:
                leadji_info[col] = pd.to_numeric(leadji_info[col], errors="coerce").fillna(0)

    if not leadji_stock.empty:
        leadji_stock.columns = [str(c).strip() for c in leadji_stock.columns]
        for col in ["기초", "입고", "출고", "재고", "검사대기"]:
            if col in leadji_stock.columns:
                leadji_stock[col] = pd.to_numeric(leadji_stock[col], errors="coerce").fillna(0)

    return leadji_info, leadji_stock


def render_shortage_dashboard(df: pd.DataFrame, updated_at: str) -> None:
    enriched_df = add_rq_group_columns(df)
    filtered = apply_filters(enriched_df, updated_at)
    q_summary = build_qcode_summary(filtered)
    rq_summary = build_rq_group_summary(filtered)

    tab_p, tab_q, tab_rq = st.tabs(["P코드 기준 현황", "Q코드 기준 집계", "RQ코드 그룹 집계"])

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
                "R코드",
                "Q코드",
                "R코드 제품명",
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

    with tab_rq:
        r1, r2, r3 = st.columns(3)
        r1.metric("RQ 그룹 수", f"{len(rq_summary):,}")
        if rq_summary.empty:
            r2.metric("동일 RQ 그룹 수(P코드2+)", "0")
            r3.metric("사출 부족수량 합계", "0")
            st.info("표시할 RQ 그룹 데이터가 없습니다.")
        else:
            same_group_count = int((rq_summary["P코드 수"] >= 2).sum())
            r2.metric("동일 RQ 그룹 수(P코드2+)", f"{same_group_count:,}")
            r3.metric("사출 부족수량 합계", f"{rq_summary['사출 부족수량'].sum():,.0f}")

            rq_table_display = format_numeric_columns_for_display(rq_summary)
            st.dataframe(
                rq_table_display,
                use_container_width=True,
                height=700,
            )


def render_leadji_dashboard(updated_at: str, leadji_info: pd.DataFrame, leadji_stock: pd.DataFrame) -> None:
    st.subheader("리드지 현황")
    st.caption(f"업데이트: {updated_at}")

    info_tab, stock_tab = st.tabs(["리드지 정보", "리드지 재고"])

    with info_tab:
        if leadji_info.empty:
            st.warning("리드지정보 시트를 찾지 못했습니다.")
        else:
            qcol, _ = st.columns([3.0, 1.0])
            with qcol:
                info_query = st.text_input(
                    "통합 검색 (판매/생산/분리/B코드)",
                    value="",
                    key="leadji_info_query",
                    placeholder="예: T4061, P1089, BS0054",
                ).strip()

            filtered_info = leadji_info.copy()
            info_search_cols = [
                c
                for c in ["판매", "판매명", "생산", "생산명", "분리", "분리명", "B1코드", "B1코드명", "B2코드", "B2코드명", "B3코드", "B3코드명"]
                if c in filtered_info.columns
            ]
            if info_search_cols:
                filtered_info = filter_with_terms_any(filtered_info, info_search_cols, info_query)

            m1, m2, m3 = st.columns(3)
            for metric_col, metric_box, metric_name in [
                ("판매", m1, "판매 코드 수"),
                ("생산", m2, "생산 코드 수"),
                ("분리", m3, "분리 코드 수"),
            ]:
                if metric_col in filtered_info.columns:
                    valid = (
                        filtered_info[metric_col]
                        .astype(str)
                        .str.strip()
                        .replace({"nan": "", "None": ""})
                    )
                    metric_box.metric(metric_name, f"{(valid != '').sum():,}")
                else:
                    metric_box.metric(metric_name, "-")

            info_cols = [
                c
                for c in [
                    "신규분류요약",
                    "판매",
                    "판매명",
                    "생산",
                    "생산명",
                    "분리",
                    "분리명",
                    "B1코드",
                    "B1코드명",
                    "B1소요량",
                    "B2코드",
                    "B2코드명",
                    "B2소요량",
                    "B3코드",
                    "B3코드명",
                    "B3소요량",
                ]
                if c in filtered_info.columns
            ]
            info_table = filtered_info[info_cols] if info_cols else filtered_info
            st.dataframe(
                format_numeric_columns_for_display(info_table),
                use_container_width=True,
                height=700,
            )

    with stock_tab:
        if leadji_stock.empty:
            st.warning("리드지 재고 시트를 찾지 못했습니다.")
        else:
            qcol, optcol = st.columns([3.0, 1.0])
            with qcol:
                stock_query = st.text_input(
                    "통합 검색 (품목코드/품목명/중분류/창고)",
                    value="",
                    key="leadji_stock_query",
                    placeholder="예: BS0054, 리드지, 원료창고",
                ).strip()
            with optcol:
                only_positive_stock = st.checkbox("재고>0만", value=True, key="leadji_only_positive")

            filtered_stock = leadji_stock.copy()
            stock_search_cols = [c for c in ["품목코드", "품목명", "중분류", "창고", "구분"] if c in filtered_stock.columns]
            if stock_search_cols:
                filtered_stock = filter_with_terms_any(filtered_stock, stock_search_cols, stock_query)

            if only_positive_stock and "재고" in filtered_stock.columns:
                filtered_stock = filtered_stock[filtered_stock["재고"] > 0]

            c1, c2, c3 = st.columns(3)
            if "품목코드" in filtered_stock.columns:
                c1.metric("재고 품목 수", f"{filtered_stock['품목코드'].astype(str).nunique():,}")
            else:
                c1.metric("재고 품목 수", "-")
            c2.metric("재고 합계", f"{filtered_stock['재고'].sum():,.0f}" if "재고" in filtered_stock.columns else "-")
            c3.metric("검사대기 합계", f"{filtered_stock['검사대기'].sum():,.0f}" if "검사대기" in filtered_stock.columns else "-")

            stock_cols = [
                c
                for c in [
                    "창고",
                    "구분",
                    "품목코드",
                    "품목명",
                    "규격",
                    "대분류",
                    "중분류",
                    "Lot.No",
                    "기초",
                    "입고",
                    "출고",
                    "재고",
                    "검사대기",
                    "단위",
                ]
                if c in filtered_stock.columns
            ]
            stock_table = filtered_stock[stock_cols] if stock_cols else filtered_stock
            st.dataframe(
                format_numeric_columns_for_display(stock_table),
                use_container_width=True,
                height=700,
            )


def main() -> None:
    st.title("이니셜/거래처/품목코드 기준 제품 부족수량 현황")
    st.caption("기준: 누수규격검사 생산수량 = 부족수량, 품목코드는 P코드만 표시")

    try:
        refresh_key = build_data_refresh_key(BASE_DIR)
        df, _, _ = load_data(refresh_key)
        leadji_info, leadji_stock = load_leadji_data(refresh_key)
    except Exception as exc:
        st.error(f"데이터 로드 실패: {exc}")
        st.stop()

    updated_at = get_data_updated_at(BASE_DIR)
    top_shortage_tab, top_leadji_tab = st.tabs(["제품 부족수량 현황", "리드지 현황"])

    with top_shortage_tab:
        render_shortage_dashboard(df, updated_at)

    with top_leadji_tab:
        render_leadji_dashboard(updated_at, leadji_info, leadji_stock)


if __name__ == "__main__":
    main()
