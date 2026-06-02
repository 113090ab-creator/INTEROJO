import hashlib
import re
import shutil
import uuid
from datetime import datetime
from io import BytesIO
from pathlib import Path
from zoneinfo import ZoneInfo

import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(page_title="생산현황", layout="wide")

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_WORKSPACE_ROOT = BASE_DIR / ".uploaded_workspaces"
DISPLAY_TZ = ZoneInfo("Asia/Seoul")
LEADJI_REQUIRED_QTY_COL = "[45]하이드레이션/전면검사 필요수량"
LEADJI_REQUIRED_DUE_COL = "[45]하이드레이션/전면검사 납기일"
LEADJI_COMPLETED_STOCK_COL = "누수규격검사 창고"

WAREHOUSE_MAP = {
    "사출창고": "사출창고",
    "분리창고": "분리창고",
    "검사접착": "검사접착창고",
    "누수규격검사": "누수규격검사 창고",
}
TARGET_WAREHOUSES = list(WAREHOUSE_MAP.keys())

COLUMN_LABEL_ALIASES = {
    "사출창고": "사출 재고",
    "분리창고": "분리 재고",
    "검사접착창고": "검사접착 재고",
    "누수규격검사 창고": "누수규격 재고",
    "사출창고 합계": "사출 재고",
    "분리창고 합계": "분리 재고",
    "검사접착창고 합계": "검사접착 재고",
    "누수규격검사창고 합계": "누수규격 재고",
    "공정재고 합계": "공정재고",
    "사출 부족수량": "사출부족",
    "사출생산필요수량": "사출필요",
    "생산필요수량": "생산필요",
    "최소납기일": "생산 최소 납기일",
}


def inject_dashboard_theme() -> None:
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
        .stApp {
            background: #FFFFFF;
            color: #111827;
            font-family: Inter, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        }
        [data-testid="stAppViewContainer"] > .main {
            background: linear-gradient(180deg, #FFFFFF 0%, #F4F6F8 100%);
        }
        [data-testid="stHeader"] {
            background: rgba(255, 255, 255, 0.94);
            backdrop-filter: blur(8px);
        }
        .main .block-container,
        .block-container {
            max-width: 100%;
            padding-left: 2rem;
            padding-right: 2rem;
            padding-top: 24px;
            padding-bottom: 44px;
        }
        [data-testid="stSidebar"] {
            background: #FFFFFF;
            border-right: 1px solid #E5E7EB;
        }
        [data-testid="stSidebar"] * {
            font-family: Inter, sans-serif;
        }
        [data-testid="stSidebar"] label,
        [data-testid="stSidebar"] span,
        [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p {
            color: #334155;
        }
        .sidebar-brand {
            display: flex;
            align-items: center;
            gap: 10px;
            padding: 10px 0 18px;
            margin-bottom: 10px;
            border-bottom: 1px solid #E5E7EB;
        }
        .sidebar-brand-icon {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 30px;
            height: 30px;
            border-radius: 8px;
            background: #EEF2FF;
            color: #1A2B5E;
        }
        .sidebar-brand-title {
            color: #111827;
            font-size: 19px;
            font-weight: 850;
            letter-spacing: 0;
        }
        .sidebar-divider {
            height: 1px;
            background: #E5E7EB;
            margin: 16px 0;
        }
        [data-testid="stSidebar"] [role="radiogroup"] label {
            border-radius: 8px;
            padding: 6px 8px;
            margin-bottom: 3px;
        }
        h1, h2, h3 {
            color: #111827;
            letter-spacing: 0;
        }
        h1 {
            font-size: 28px;
            font-weight: 800;
            margin-bottom: 8px;
        }
        h2, h3 {
            font-weight: 800;
        }
        .dashboard-hero {
            border: 1px solid #E5E7EB;
            border-radius: 14px;
            background: linear-gradient(135deg, #FFFFFF 0%, #F4F6F8 100%);
            box-shadow: 0 14px 40px rgba(15, 23, 42, 0.08);
            padding: 22px 24px;
            margin-bottom: 18px;
            border-left: 5px solid #1A2B5E;
        }
        .dashboard-hero-title {
            color: #1A2B5E;
            font-size: 30px;
            font-weight: 850;
            line-height: 1.25;
            margin: 0 0 6px;
        }
        .dashboard-hero-subtitle {
            color: #64748B;
            font-size: 14px;
            margin: 0;
        }
        .sidebar-section-title {
            color: #1A2B5E;
            font-size: 14px;
            font-weight: 850;
            padding: 4px 0 9px;
            border-bottom: 1px solid #E5E7EB;
            margin-bottom: 12px;
        }
        [data-testid="stCaptionContainer"] {
            color: #6b7280;
        }
        [data-testid="stAlert"] {
            border-radius: 8px;
            border: 1px solid #D7DBE8;
            background: #F7F8FB;
            box-shadow: 0 4px 14px rgba(26, 43, 94, 0.06);
        }
        [data-testid="stDataFrame"] {
            border: 1px solid #E5E7EB;
            border-radius: 12px;
            overflow: hidden;
            box-shadow: 0 10px 30px rgba(15, 23, 42, 0.06);
            background: #ffffff;
        }
        [data-testid="stExpander"] {
            border: 1px solid #E5E7EB;
            border-radius: 12px;
            background: #ffffff;
            box-shadow: 0 8px 22px rgba(15, 23, 42, 0.05);
        }
        [data-testid="stTextInput"] input {
            border-radius: 8px;
            border-color: #D1D5DB;
            background: #ffffff;
            color: #111827;
        }
        [data-testid="stTextInput"] input:focus {
            border-color: #1A2B5E;
            box-shadow: 0 0 0 2px rgba(26, 43, 94, 0.10);
        }
        div[data-testid="stMetric"] {
            background: #FFFFFF;
            border: 1px solid #E5E7EB;
            border-radius: 14px;
            padding: 15px 16px;
            border-left: 4px solid #1A2B5E;
            box-shadow: 0 10px 28px rgba(15, 23, 42, 0.07);
        }
        div[data-testid="stMetric"] label {
            color: #64748B !important;
            font-weight: 700;
        }
        div[data-testid="stMetric"] [data-testid="stMetricValue"] {
            color: #374151;
            font-weight: 850;
            white-space: nowrap;
            overflow: visible;
            text-overflow: unset;
            font-size: clamp(22px, 1.7vw, 32px);
        }
        .ops-kpi-card {
            min-height: 108px;
            border-radius: 14px;
            background: #FFFFFF;
            border: 1px solid #E5E7EB;
            border-left: 5px solid #2563EB;
            box-shadow: 0 10px 28px rgba(15, 23, 42, 0.07);
            padding: 16px 18px;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            overflow: visible;
        }
        .ops-kpi-card.risk {
            border-left-color: #DC2626;
            background: linear-gradient(180deg, #FFFFFF 0%, #FFF7F7 100%);
        }
        .ops-kpi-card.stock {
            border-left-color: #2563EB;
            background: linear-gradient(180deg, #FFFFFF 0%, #F8FAFF 100%);
        }
        .kpi-label {
            color: #64748B;
            font-size: 13px;
            font-weight: 800;
            line-height: 1.25;
        }
        .kpi-value {
            white-space: nowrap;
            overflow: visible;
            text-overflow: unset;
            font-size: clamp(24px, 2vw, 36px);
            font-weight: 850;
            line-height: 1.1;
            letter-spacing: 0;
            color: #374151;
        }
        .ops-kpi-card.risk .kpi-value {
            color: #DC2626;
        }
        .ops-kpi-card.stock .kpi-value {
            color: #1A2B5E;
        }
        .stButton > button,
        [data-testid="stDownloadButton"] button {
            border-radius: 8px;
            border: 1px solid #1A2B5E;
            background: #1A2B5E;
            color: #FFFFFF;
            font-weight: 700;
            box-shadow: 0 6px 16px rgba(26, 43, 94, 0.16);
        }
        .stButton > button:hover,
        [data-testid="stDownloadButton"] button:hover {
            border-color: #233A7A;
            color: #FFFFFF;
            background: #233A7A;
        }
        [data-testid="stSegmentedControl"] {
            background: #F4F6F8;
            border: 1px solid #E5E7EB;
            border-radius: 10px;
            padding: 4px;
        }
        .dashboard-section-header {
            display: flex;
            align-items: center;
            gap: 10px;
            margin: 22px 0 10px;
        }
        .dashboard-section-header h3 {
            margin: 0;
            font-size: 20px;
            line-height: 1.2;
        }
        .dashboard-count-badge {
            display: inline-flex;
            align-items: center;
            border-radius: 999px;
            background: #EEF2FF;
            color: #1A2B5E;
            padding: 4px 9px;
            font-size: 12px;
            font-weight: 800;
        }
        .dashboard-section-subtle {
            color: #6b7280;
            font-size: 13px;
            margin-top: -4px;
            margin-bottom: 10px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


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


def find_demand_update_file(base_dir: Path) -> Path | None:
    xlsx_files = [p for p in base_dir.glob("*.xlsx") if not p.name.startswith("~$")]
    if not xlsx_files:
        return None

    exact = [p for p in xlsx_files if p.name == "수요정보(전공정).xlsx"]
    if exact:
        return max(exact, key=lambda p: p.stat().st_mtime)

    normalized = lambda s: str(s).replace(" ", "")
    full_process = [p for p in xlsx_files if "수요정보(전공정)" in normalized(p.stem)]
    if full_process:
        return max(full_process, key=lambda p: p.stat().st_mtime)

    demand_info = [p for p in xlsx_files if "수요정보" in normalized(p.stem)]
    if demand_info:
        return max(demand_info, key=lambda p: p.stat().st_mtime)

    demand_like = [p for p in xlsx_files if "수요" in normalized(p.stem)]
    if demand_like:
        return max(demand_like, key=lambda p: p.stat().st_mtime)

    return None


def find_leadji_order_status_file(base_dir: Path) -> Path | None:
    xlsx_files = [p for p in base_dir.glob("*.xlsx") if not p.name.startswith("~$")]
    if not xlsx_files:
        return None

    normalized = lambda s: str(s).replace(" ", "")
    candidates = [p for p in xlsx_files if normalized(p.name) == "리드지발주현황.xlsx"]
    if not candidates:
        candidates = [p for p in xlsx_files if "리드지" in normalized(p.stem) and "발주현황" in normalized(p.stem)]
    if not candidates:
        candidates = [p for p in xlsx_files if "발주현황" in normalized(p.stem)]
    if not candidates:
        return None
    return max(candidates, key=lambda p: p.stat().st_mtime)


def get_data_updated_at(base_dir: Path) -> str:
    dem_path = base_dir / "수요정보(전공정).xlsx"
    if not dem_path.exists():
        return "-"

    # 배포 환경에서는 파일시스템 mtime이 배포 시각으로 바뀔 수 있어
    # 엑셀 내부 문서 속성(modified)을 우선 사용한다.
    latest_dt: datetime | None = None
    try:
        wb = openpyxl.load_workbook(dem_path, read_only=True, data_only=True)
        modified = wb.properties.modified
        wb.close()
        if isinstance(modified, datetime):
            if modified.tzinfo is None:
                # Excel core property is commonly stored as UTC naive datetime.
                latest_dt = modified.replace(tzinfo=ZoneInfo("UTC")).astimezone(DISPLAY_TZ)
            else:
                latest_dt = modified.astimezone(DISPLAY_TZ)
    except Exception:
        latest_dt = None

    if latest_dt is None:
        latest_dt = datetime.fromtimestamp(dem_path.stat().st_mtime, tz=DISPLAY_TZ)

    return latest_dt.strftime("%Y-%m-%d %H:%M:%S")


def get_or_create_upload_session_id() -> str:
    key = "upload_session_id"
    if key not in st.session_state:
        st.session_state[key] = uuid.uuid4().hex
    return str(st.session_state[key])


def stage_uploaded_data_files(
    base_dir: Path,
    inventory_file,
    demand_file,
    reference_file=None,
) -> Path:
    session_id = get_or_create_upload_session_id()
    session_dir = UPLOAD_WORKSPACE_ROOT / session_id
    session_dir.mkdir(parents=True, exist_ok=True)

    inv_bytes = bytes(inventory_file.getbuffer())
    dem_bytes = bytes(demand_file.getbuffer())
    ref_bytes = bytes(reference_file.getbuffer()) if reference_file is not None else None

    local_ref = find_product_name_reference_file(base_dir) if reference_file is None else None
    local_ref_sig = ""
    if local_ref is not None and local_ref.exists():
        stat = local_ref.stat()
        local_ref_sig = f"{local_ref.name}:{stat.st_size}:{stat.st_mtime_ns}"

    upload_signature = "|".join(
        [
            hashlib.md5(inv_bytes).hexdigest(),
            hashlib.md5(dem_bytes).hexdigest(),
            hashlib.md5(ref_bytes).hexdigest() if ref_bytes is not None else "-",
            local_ref_sig,
        ]
    )
    signature_key = f"upload_workspace_signature_{session_id}"
    inv_staged = session_dir / "ODV_WIP_uploaded.xlsx"
    dem_staged = session_dir / "수요정보(전공정).xlsx"
    ref_staged = session_dir / "제품명 기준 정보.xlsx"
    if (
        st.session_state.get(signature_key) == upload_signature
        and inv_staged.exists()
        and dem_staged.exists()
        and (reference_file is not None or ref_staged.exists() or local_ref is None)
    ):
        return session_dir

    for old_xlsx in session_dir.glob("*.xlsx"):
        old_xlsx.unlink(missing_ok=True)

    inv_staged.write_bytes(inv_bytes)
    dem_staged.write_bytes(dem_bytes)
    ref_dst = ref_staged
    if reference_file is not None:
        ref_dst.write_bytes(ref_bytes if ref_bytes is not None else b"")
    else:
        if local_ref is not None and local_ref.exists():
            shutil.copy2(local_ref, ref_dst)

    st.session_state[signature_key] = upload_signature
    return session_dir


def select_data_source(base_dir: Path) -> tuple[Path, str, str]:
    st.subheader("데이터 소스")
    with st.expander("업로드 설정", expanded=False):
        use_uploaded = st.toggle("업로드 파일 사용", value=False, key="use_uploaded_data_mode")
        inv_file = st.file_uploader("재고 파일(.xlsx)", type=["xlsx"], key="upload_inventory_xlsx", disabled=not use_uploaded)
        dem_file = st.file_uploader("수요 파일(.xlsx)", type=["xlsx"], key="upload_demand_xlsx", disabled=not use_uploaded)
        ref_file = st.file_uploader(
            "기준정보 파일(.xlsx, 선택)",
            type=["xlsx"],
            key="upload_reference_xlsx",
            disabled=not use_uploaded,
            help="미업로드 시 로컬의 '제품명 기준 정보.xlsx'를 사용합니다.",
        )

        if not use_uploaded:
            st.caption("현재 설정: 로컬 폴더의 파일 사용")

    if not use_uploaded:
        return base_dir, "로컬 파일", get_data_updated_at(base_dir)

    if inv_file is None or dem_file is None:
        st.info("업로드 모드에서는 재고/수요 파일 2개 업로드가 필요합니다.")
        st.stop()

    workspace_dir = stage_uploaded_data_files(base_dir, inv_file, dem_file, ref_file)
    source_label = f"업로드 파일 ({inv_file.name}, {dem_file.name})"
    updated_at = get_data_updated_at(workspace_dir)
    return workspace_dir, source_label, updated_at


def pick_first_existing_column(columns: list[str], candidates: list[str]) -> str | None:
    for col in candidates:
        if col in columns:
            return col
    return None


def parse_mixed_excel_date(series: pd.Series) -> pd.Series:
    """Parse mixed date inputs safely, including Excel serial dates."""
    text = series.astype(str).str.strip()
    invalid_tokens = {"", "nan", "none", "nat"}
    cleaned = series.where(~text.str.lower().isin(invalid_tokens), pd.NA)

    # First pass: strings/datetime objects.
    parsed = pd.to_datetime(cleaned, errors="coerce")

    # Second pass: Excel serial numbers (days since 1899-12-30).
    # IMPORTANT: only treat true numeric-like cells as serials.
    # Datetime64 values can be converted to large integers (ns) by to_numeric,
    # which would overwrite valid dates with NaT if we don't filter first.
    obj = cleaned.astype("object")
    numeric_like = obj.map(lambda v: isinstance(v, (int, float)) and not isinstance(v, bool))
    numeric_text = text.str.fullmatch(r"[+-]?\d+(?:\.\d+)?").fillna(False)
    numeric_mask_source = numeric_like | numeric_text

    numeric = pd.to_numeric(obj.where(numeric_mask_source), errors="coerce")
    numeric_mask = numeric.notna() & (numeric > 0)
    if numeric_mask.any():
        parsed.loc[numeric_mask] = pd.to_datetime(
            numeric.loc[numeric_mask], unit="D", origin="1899-12-30", errors="coerce"
        )

    return parsed


def parse_mixed_numeric(series: pd.Series) -> pd.Series:
    """Parse mixed numeric inputs safely (number/string/comma text)."""
    text = series.astype(str).str.strip()
    invalid_tokens = {"", "nan", "none", "nat", "-"}
    normalized = text.where(~text.str.lower().isin(invalid_tokens), pd.NA)

    # Support accounting-style negatives like "(1,234)".
    normalized = normalized.str.replace(r"^\((.*)\)$", r"-\1", regex=True)
    normalized = normalized.str.replace(",", "", regex=False)
    normalized = normalized.str.replace("\u00a0", "", regex=False).str.replace(" ", "", regex=False)

    return pd.to_numeric(normalized, errors="coerce").fillna(0)


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
            "재고량": parse_mixed_numeric(inv[qty_col]),
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


def extract_demand_header_info(dem_path: Path) -> tuple[
    dict[str, str], dict[str, int], list[int], list[int], dict[str, int]
]:
    header_rows = pd.read_excel(dem_path, sheet_name=0, header=None, nrows=2)
    if header_rows.shape[0] < 2:
        return {}, {}, [], [], {}

    top_row = header_rows.iloc[0]
    second_row = header_rows.iloc[1]

    code_map: dict[str, str] = {}
    warehouse_qty_col_indices: dict[str, int] = {}
    qty_col_indices: list[int] = []
    total_qty_col_indices: list[int] = []
    process_qty_col_indices: dict[str, int] = {}

    for idx, column_name in second_row.items():
        if "생산 수량" not in str(column_name):
            continue

        idx = int(idx)
        qty_col_indices.append(idx)

        top_label = str(top_row.iloc[idx]).strip()
        if not top_label or top_label.lower() == "nan":
            continue
        process_qty_col_indices[top_label.replace(" ", "")] = idx
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

    return code_map, warehouse_qty_col_indices, qty_col_indices, total_qty_col_indices, process_qty_col_indices


def map_demand_code_to_process_code(demand_code: str, process_prefix: str) -> str:
    code = str(demand_code).strip()
    if not code or code.lower() == "nan":
        return code

    letter_pattern = re.match(r"^P(\d{4})([A-Z])(.*)$", code)
    if letter_pattern:
        return f"{process_prefix}{letter_pattern.group(1)}{letter_pattern.group(3)}"
    if code.startswith("P"):
        return f"{process_prefix}{code[1:]}"
    if code[0] in {"Q", "R"} and len(code) > 1:
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


def find_reference_sheet_with_columns(
    ref_path: Path, sheet_names: list[str], required_columns: set[str], preferred_name: str | None = None
) -> str | None:
    if preferred_name:
        normalized = preferred_name.replace(" ", "")
        by_name = next((s for s in sheet_names if str(s).replace(" ", "") == normalized), None)
        if by_name is not None:
            return by_name

    for sheet_name in sheet_names:
        try:
            preview = pd.read_excel(ref_path, sheet_name=sheet_name, nrows=0)
        except Exception:
            continue
        cols = {str(c).strip() for c in preview.columns}
        if required_columns.issubset(cols):
            return sheet_name
    return None


def load_bom_base_code_maps(base_dir: Path) -> tuple[dict[str, str], dict[str, str], dict[str, str], dict[str, str]]:
    ref_path = find_product_name_reference_file(base_dir)
    if ref_path is None:
        return {}, {}, {}, {}

    sheet_names = pd.ExcelFile(ref_path).sheet_names
    bom_sheet = find_reference_sheet_with_columns(
        ref_path, sheet_names, {"SALES_ITEM_CD", "FROM_ITEM_ID"}, preferred_name="BOM정보"
    )
    if bom_sheet is None:
        return {}, {}, {}, {}

    use_cols = ["SALES_ITEM_CD", "TO_ITEM_ID", "FROM_ITEM_ID", "SEQ"]
    bom = pd.read_excel(ref_path, sheet_name=bom_sheet, usecols=lambda c: str(c).strip() in set(use_cols))
    bom.columns = [str(c).strip() for c in bom.columns]
    if not {"SALES_ITEM_CD", "FROM_ITEM_ID"}.issubset(bom.columns):
        return {}, {}, {}, {}

    bom["SALES_ITEM_CD"] = bom["SALES_ITEM_CD"].astype(str).str.strip()
    bom["FROM_ITEM_ID"] = bom["FROM_ITEM_ID"].astype(str).str.strip()
    if "SEQ" in bom.columns:
        bom["SEQ"] = pd.to_numeric(bom["SEQ"], errors="coerce").fillna(9999)
    else:
        bom["SEQ"] = 9999

    bom = bom[
        (bom["SALES_ITEM_CD"] != "")
        & (bom["SALES_ITEM_CD"].str.lower() != "nan")
        & (bom["FROM_ITEM_ID"] != "")
        & (bom["FROM_ITEM_ID"].str.lower() != "nan")
    ].copy()
    if bom.empty:
        return {}, {}, {}, {}

    # Exact TO_ITEM_ID mapping (authoritative when available).
    if "TO_ITEM_ID" in bom.columns:
        bom["TO_ITEM_ID"] = bom["TO_ITEM_ID"].astype(str).str.strip()
        exact = bom[
            (bom["TO_ITEM_ID"] != "")
            & (bom["TO_ITEM_ID"].str.lower() != "nan")
            & bom["FROM_ITEM_ID"].str.match(r"^[QR].+", na=False)
        ].copy()
        exact = exact.sort_values(["TO_ITEM_ID", "SEQ"], ascending=[True, True]).drop_duplicates(
            subset=["TO_ITEM_ID"], keep="first"
        )
    else:
        exact = pd.DataFrame(columns=["TO_ITEM_ID", "FROM_ITEM_ID"])

    q_exact_map: dict[str, str] = {}
    r_exact_map: dict[str, str] = {}
    if not exact.empty:
        for to_code, from_code in exact[["TO_ITEM_ID", "FROM_ITEM_ID"]].itertuples(index=False):
            from_code = str(from_code).strip()
            if from_code.startswith("Q"):
                q_exact_map[to_code] = from_code
                if len(from_code) > 1:
                    r_exact_map[to_code] = "R" + from_code[1:]
            elif from_code.startswith("R"):
                r_exact_map[to_code] = from_code
                if len(from_code) > 1:
                    q_exact_map[to_code] = "Q" + from_code[1:]

    bom["SALES_CODE5"] = bom["SALES_ITEM_CD"].str[:5]
    bom["FROM_CODE5"] = bom["FROM_ITEM_ID"].str[:5]
    bom = bom[bom["SALES_CODE5"].str.match(r"^[PQRSTU]\d{4}$", na=False)]
    bom = bom[bom["FROM_CODE5"].str.match(r"^[PQRSTU]\d{4}$", na=False)]
    if bom.empty:
        return {}, {}, r_exact_map, q_exact_map

    bom = bom.sort_values(["SALES_CODE5", "SEQ"], ascending=[True, True])

    q_df = bom[bom["FROM_CODE5"].str.startswith("Q")].drop_duplicates(subset=["SALES_CODE5"], keep="first")
    r_df = bom[bom["FROM_CODE5"].str.startswith("R")].drop_duplicates(subset=["SALES_CODE5"], keep="first")
    q_base_map = q_df.set_index("SALES_CODE5")["FROM_CODE5"].to_dict()
    r_base_map = r_df.set_index("SALES_CODE5")["FROM_CODE5"].to_dict()

    # If BOM has only Q mapping for a sales code, derive R base from the same numeric part.
    for sales_code5, q_code5 in q_base_map.items():
        if sales_code5 not in r_base_map and str(q_code5).startswith("Q") and len(str(q_code5)) >= 5:
            r_base_map[sales_code5] = "R" + str(q_code5)[1:5]

    return r_base_map, q_base_map, r_exact_map, q_exact_map


def load_sheet2_group_map(base_dir: Path) -> dict[str, str]:
    ref_path = find_product_name_reference_file(base_dir)
    if ref_path is None:
        return {}

    sheet_names = pd.ExcelFile(ref_path).sheet_names
    if len(sheet_names) < 2:
        return {}

    sheet_name = find_reference_sheet_with_columns(
        ref_path, sheet_names, {"코드", "시트이름"}, preferred_name="분류정보"
    )
    if sheet_name is None:
        return {}

    sheet2 = pd.read_excel(ref_path, sheet_name=sheet_name)
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

    sheet_name = find_reference_sheet_with_columns(
        ref_path, sheet_names, {"코드", "Q코드", "R코드"}, preferred_name="분류정보"
    )
    if sheet_name is None:
        return {}, {}, {}

    sheet2 = pd.read_excel(ref_path, sheet_name=sheet_name)
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
    code5_df = df.drop_duplicates(subset=["코드5"], keep="first")

    q_map = code5_df.set_index("코드5")["Q코드"].to_dict()
    r_map = code5_df.set_index("코드5")["R코드"].to_dict()
    if name_col:
        r_name_df = df[(df[name_col] != "") & (df[name_col].str.lower() != "nan")].copy()
        r_name_df["R코드5"] = r_name_df["R코드"].str[:5]
        r_name_df = r_name_df[(r_name_df["R코드5"] != "") & (r_name_df["R코드5"].str.lower() != "nan")]
        r_name_df = r_name_df.drop_duplicates(subset=["R코드5", name_col], keep="first")
        r_name_df = r_name_df.drop_duplicates(subset=["R코드5"], keep="first")
        r_name_map = r_name_df.set_index("R코드5")[name_col].to_dict()
    else:
        r_name_map = {}

    # Fallback: enrich R코드5 -> 제품명 from sheet0(제품명정보).
    # 제품명코드가 P/Q/R/T/U + 4자리인 경우 모두 R + 4자리 키로 정규화한다.
    try:
        sheet1 = pd.read_excel(ref_path, sheet_name=sheet_names[0])
        sheet1.columns = [str(c).strip() for c in sheet1.columns]
        if len(sheet1.columns) >= 2:
            code_col = "제품명코드" if "제품명코드" in sheet1.columns else sheet1.columns[0]
            name1_col = "제품명" if "제품명" in sheet1.columns else sheet1.columns[1]
            fb = sheet1[[code_col, name1_col]].copy()
            fb[code_col] = fb[code_col].astype(str).str.strip()
            fb[name1_col] = fb[name1_col].astype(str).str.strip()
            fb = fb[
                (fb[code_col] != "")
                & (fb[code_col].str.lower() != "nan")
                & (fb[name1_col] != "")
                & (fb[name1_col].str.lower() != "nan")
            ]
            fb["코드5"] = fb[code_col].str[:5]
            fb = fb[fb["코드5"].str.match(r"^[PQRSTU]\d{4}$", na=False)]
            fb["R코드5"] = "R" + fb["코드5"].str[-4:]
            fb = fb.drop_duplicates(subset=["R코드5", name1_col], keep="first")
            fb = fb.drop_duplicates(subset=["R코드5"], keep="first")
            for k, v in fb.set_index("R코드5")[name1_col].to_dict().items():
                if k not in r_name_map:
                    r_name_map[k] = v
    except Exception:
        pass
    return r_map, q_map, r_name_map


def load_leadji_process_maps(base_dir: Path) -> tuple[dict[str, str], dict[str, str]]:
    ref_path = find_product_name_reference_file(base_dir)
    if ref_path is None:
        return {}, {}

    sheet_names = pd.ExcelFile(ref_path).sheet_names
    if len(sheet_names) < 3:
        return {}, {}

    leadji_sheet = next((s for s in sheet_names if s.replace(" ", "") == "리드지정보"), sheet_names[2])
    leadji = pd.read_excel(ref_path, sheet_name=leadji_sheet)
    leadji.columns = [str(c).strip() for c in leadji.columns]

    prod_col = "생산" if "생산" in leadji.columns else (leadji.columns[3] if len(leadji.columns) > 3 else None)
    q_col = "분리" if "분리" in leadji.columns else (leadji.columns[9] if len(leadji.columns) > 9 else None)
    r_col = "사출" if "사출" in leadji.columns else (leadji.columns[21] if len(leadji.columns) > 21 else None)
    if prod_col is None or q_col is None or r_col is None:
        return {}, {}

    df = leadji[[prod_col, q_col, r_col]].copy()
    for col in [prod_col, q_col, r_col]:
        df[col] = df[col].astype(str).str.strip()
        df.loc[df[col].str.lower() == "nan", col] = ""

    df = df[df[prod_col].str.startswith("P")]
    if df.empty:
        return {}, {}

    df["코드5"] = df[prod_col].str[:5]
    df = df[(df["코드5"] != "") & (df["코드5"].str.lower() != "nan")]

    def normalize_to_code(code: str, prefix: str) -> str:
        v = str(code).strip()
        if not v or v.lower() == "nan":
            return ""
        if v.startswith(prefix):
            return v
        if v.startswith("P"):
            return f"{prefix}{v[1:]}"
        return v

    df["Q정규"] = df[q_col].map(lambda x: normalize_to_code(x, "Q"))
    df["R정규"] = df[r_col].map(lambda x: normalize_to_code(x, "R"))

    q_df = df[df["Q정규"] != ""].drop_duplicates(subset=["코드5"], keep="first")
    r_df = df[df["R정규"] != ""].drop_duplicates(subset=["코드5"], keep="first")
    q_map = q_df.set_index("코드5")["Q정규"].to_dict()
    r_map = r_df.set_index("코드5")["R정규"].to_dict()
    return r_map, q_map


def merge_mapped_base_code(inferred_code: str, mapped_base_code: str, prefix: str) -> str:
    inferred = str(inferred_code).strip()
    mapped = str(mapped_base_code).strip()

    if not mapped or mapped.lower() == "nan":
        return inferred
    if not inferred or inferred.lower() == "nan":
        return mapped

    if inferred.startswith(prefix) and mapped.startswith(prefix) and len(inferred) >= 5 and len(mapped) >= 5:
        return mapped[:5] + inferred[5:]
    return mapped


def iter_inventory_code_candidates(process_code: str) -> list[str]:
    code = str(process_code).strip()
    if not code or code.lower() == "nan":
        return []

    candidates = [code]
    bul_match = re.match(r"^(.*BUL)\d+$", code, flags=re.IGNORECASE)
    if bul_match:
        candidates.append(bul_match.group(1))

    # 순서를 유지한 unique
    return list(dict.fromkeys(candidates))


def lookup_stock_qty(stock_map: dict[str, float], process_code: str) -> float:
    for candidate in iter_inventory_code_candidates(process_code):
        qty = stock_map.get(candidate)
        if qty is not None:
            return float(qty)
    return 0.0


def resolve_process_code_for_stock(stock_map: dict[str, float], process_code: str) -> str:
    code = str(process_code).strip()
    if not code or code.lower() == "nan":
        return code
    for candidate in iter_inventory_code_candidates(code):
        if candidate in stock_map:
            return candidate
    return code


def normalize_leadji_code_key(value: object) -> str:
    code = str(value).strip().upper()
    if code in {"", "NAN", "NONE", "-", "NULL"}:
        return ""
    code = re.sub(r"\s+", "", code)
    matched = re.match(r"^([A-Z]{2}\d{4})", code)
    return matched.group(1) if matched else code


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


def build_reference_refresh_key(base_dir: Path) -> str:
    ref_path = find_product_name_reference_file(base_dir)
    if ref_path is None:
        return "-"
    stat = ref_path.stat()
    return f"{ref_path.name}:{stat.st_size}:{stat.st_mtime_ns}"


def build_leadji_order_refresh_key(base_dir: Path) -> str:
    order_path = find_leadji_order_status_file(base_dir)
    if order_path is None:
        return "-"
    stat = order_path.stat()
    return f"{order_path.name}:{stat.st_size}:{stat.st_mtime_ns}"


@st.cache_data(show_spinner=False, persist="disk")
def load_reference_maps_bundle(
    base_dir: Path, reference_refresh_key: str
) -> tuple[
    dict[str, str],
    dict[str, str],
    dict[str, str],
    dict[str, str],
    dict[str, str],
    dict[str, str],
    dict[str, str],
    dict[str, str],
    dict[str, str],
    dict[str, str],
    dict[str, str],
    dict[str, str],
]:
    _ = reference_refresh_key
    ref_path = find_product_name_reference_file(base_dir)
    empty_bundle = ({}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {})
    if ref_path is None:
        return empty_bundle

    try:
        xls = pd.ExcelFile(ref_path)
    except Exception:
        return empty_bundle

    sheet_names = xls.sheet_names
    if not sheet_names:
        return empty_bundle

    def parse_sheet(sheet_name: str, usecols=None) -> pd.DataFrame:
        try:
            df = xls.parse(sheet_name=sheet_name, usecols=usecols)
        except Exception:
            return pd.DataFrame()
        df.columns = [str(c).strip() for c in df.columns]
        return df

    def find_sheet(required_columns: set[str], preferred_name: str | None = None) -> str | None:
        if preferred_name:
            normalized = preferred_name.replace(" ", "")
            by_name = next((s for s in sheet_names if str(s).replace(" ", "") == normalized), None)
            if by_name is not None:
                return by_name

        for sheet_name in sheet_names:
            try:
                preview = xls.parse(sheet_name=sheet_name, nrows=0)
            except Exception:
                continue
            cols = {str(c).strip() for c in preview.columns}
            if required_columns.issubset(cols):
                return sheet_name
        return None

    product_name_map: dict[str, str] = {}
    product_group_map: dict[str, str] = {}
    sheet2_group_map: dict[str, str] = {}
    r_ref_map: dict[str, str] = {}
    q_ref_map: dict[str, str] = {}
    r_name_map: dict[str, str] = {}
    bom_r_base_map: dict[str, str] = {}
    bom_q_base_map: dict[str, str] = {}
    bom_r_exact_map: dict[str, str] = {}
    bom_q_exact_map: dict[str, str] = {}
    leadji_r_map: dict[str, str] = {}
    leadji_q_map: dict[str, str] = {}

    # 1) 제품명/분류 맵 + R코드명 fallback 기반 시트
    sheet0 = parse_sheet(sheet_names[0])
    if not sheet0.empty and len(sheet0.columns) >= 2:
        code_col = "제품명코드" if "제품명코드" in sheet0.columns else sheet0.columns[0]
        name_col = "제품명" if "제품명" in sheet0.columns else sheet0.columns[1]
        group_col = "분류요약" if "분류요약" in sheet0.columns else None
        if group_col is None and "판매제품군" in sheet0.columns:
            group_col = "판매제품군"
        if group_col is None and "생산제품군" in sheet0.columns:
            group_col = "생산제품군"

        selected_cols = [code_col, name_col] + ([group_col] if group_col is not None else [])
        ref_df = sheet0[selected_cols].copy()
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
        product_name_map = ref_df.set_index("코드5")[name_col].to_dict()

        if group_col is not None:
            group_df = ref_df[(ref_df[group_col] != "") & (ref_df[group_col].str.lower() != "nan")]
            product_group_map = group_df.set_index("코드5")[group_col].to_dict()

    # 2) 분류정보 시트 기반 (시트분류 + R/Q 맵 + R코드명 우선)
    group_sheet = find_sheet({"코드", "시트이름"}, preferred_name="분류정보")
    rq_sheet = find_sheet({"코드", "Q코드", "R코드"}, preferred_name="분류정보")
    group_df_source = parse_sheet(group_sheet) if group_sheet else pd.DataFrame()
    rq_df_source = group_df_source if (rq_sheet and group_sheet and rq_sheet == group_sheet) else (
        parse_sheet(rq_sheet) if rq_sheet else pd.DataFrame()
    )

    if not group_df_source.empty and {"코드", "시트이름"}.issubset(group_df_source.columns):
        s2 = group_df_source[["코드", "시트이름"]].copy()
        s2["코드"] = s2["코드"].astype(str).str.strip()
        s2["시트이름"] = s2["시트이름"].astype(str).str.strip()
        s2 = s2[
            s2["코드"].str.startswith("P")
            & (s2["코드"].str.lower() != "nan")
            & (s2["시트이름"] != "")
            & (s2["시트이름"].str.lower() != "nan")
        ]
        s2["코드5"] = s2["코드"].str[:5]
        s2 = s2.drop_duplicates(subset=["코드5"], keep="first")
        sheet2_group_map = s2.set_index("코드5")["시트이름"].to_dict()

    if not rq_df_source.empty and {"코드", "Q코드", "R코드"}.issubset(rq_df_source.columns):
        rq = rq_df_source.copy()
        for col in ["코드", "Q코드", "R코드"]:
            rq[col] = rq[col].astype(str).str.strip()
        name_col = "제품명" if "제품명" in rq.columns else None
        if name_col:
            rq[name_col] = rq[name_col].astype(str).str.strip()

        rq = rq[
            rq["코드"].str.startswith("P")
            & (rq["코드"].str.lower() != "nan")
            & (rq["Q코드"].str.lower() != "nan")
            & (rq["R코드"].str.lower() != "nan")
            & (rq["Q코드"] != "")
            & (rq["R코드"] != "")
        ]
        rq["코드5"] = rq["코드"].str[:5]
        code5_df = rq.drop_duplicates(subset=["코드5"], keep="first")
        q_ref_map = code5_df.set_index("코드5")["Q코드"].to_dict()
        r_ref_map = code5_df.set_index("코드5")["R코드"].to_dict()

        if name_col:
            r_name_df = rq[(rq[name_col] != "") & (rq[name_col].str.lower() != "nan")].copy()
            r_name_df["R코드5"] = r_name_df["R코드"].str[:5]
            r_name_df = r_name_df[(r_name_df["R코드5"] != "") & (r_name_df["R코드5"].str.lower() != "nan")]
            r_name_df = r_name_df.drop_duplicates(subset=["R코드5", name_col], keep="first")
            r_name_df = r_name_df.drop_duplicates(subset=["R코드5"], keep="first")
            r_name_map = r_name_df.set_index("R코드5")[name_col].to_dict()

    # 2-b) sheet0 기반 R코드명 fallback
    if not sheet0.empty and len(sheet0.columns) >= 2:
        code_col = "제품명코드" if "제품명코드" in sheet0.columns else sheet0.columns[0]
        name_col = "제품명" if "제품명" in sheet0.columns else sheet0.columns[1]
        fb = sheet0[[code_col, name_col]].copy()
        fb[code_col] = fb[code_col].astype(str).str.strip()
        fb[name_col] = fb[name_col].astype(str).str.strip()
        fb = fb[
            (fb[code_col] != "")
            & (fb[code_col].str.lower() != "nan")
            & (fb[name_col] != "")
            & (fb[name_col].str.lower() != "nan")
        ]
        fb["코드5"] = fb[code_col].str[:5]
        fb = fb[fb["코드5"].str.match(r"^[PQRSTU]\d{4}$", na=False)]
        fb["R코드5"] = "R" + fb["코드5"].str[-4:]
        fb = fb.drop_duplicates(subset=["R코드5", name_col], keep="first")
        fb = fb.drop_duplicates(subset=["R코드5"], keep="first")
        for key, value in fb.set_index("R코드5")[name_col].to_dict().items():
            if key not in r_name_map:
                r_name_map[key] = value

    # 3) BOM 기반 매핑
    bom_sheet = find_sheet({"SALES_ITEM_CD", "FROM_ITEM_ID"}, preferred_name="BOM정보")
    if bom_sheet is not None:
        use_cols = {"SALES_ITEM_CD", "TO_ITEM_ID", "FROM_ITEM_ID", "SEQ"}
        bom = parse_sheet(bom_sheet, usecols=lambda c: str(c).strip() in use_cols)
        if not bom.empty and {"SALES_ITEM_CD", "FROM_ITEM_ID"}.issubset(bom.columns):
            bom["SALES_ITEM_CD"] = bom["SALES_ITEM_CD"].astype(str).str.strip()
            bom["FROM_ITEM_ID"] = bom["FROM_ITEM_ID"].astype(str).str.strip()
            bom["SEQ"] = pd.to_numeric(bom["SEQ"], errors="coerce").fillna(9999) if "SEQ" in bom.columns else 9999
            bom = bom[
                (bom["SALES_ITEM_CD"] != "")
                & (bom["SALES_ITEM_CD"].str.lower() != "nan")
                & (bom["FROM_ITEM_ID"] != "")
                & (bom["FROM_ITEM_ID"].str.lower() != "nan")
            ].copy()

            if not bom.empty:
                if "TO_ITEM_ID" in bom.columns:
                    bom["TO_ITEM_ID"] = bom["TO_ITEM_ID"].astype(str).str.strip()
                    exact = bom[
                        (bom["TO_ITEM_ID"] != "")
                        & (bom["TO_ITEM_ID"].str.lower() != "nan")
                        & bom["FROM_ITEM_ID"].str.match(r"^[QR].+", na=False)
                    ].copy()
                    exact = exact.sort_values(["TO_ITEM_ID", "SEQ"], ascending=[True, True]).drop_duplicates(
                        subset=["TO_ITEM_ID"], keep="first"
                    )
                else:
                    exact = pd.DataFrame(columns=["TO_ITEM_ID", "FROM_ITEM_ID"])

                if not exact.empty:
                    for to_code, from_code in exact[["TO_ITEM_ID", "FROM_ITEM_ID"]].itertuples(index=False):
                        from_code = str(from_code).strip()
                        if from_code.startswith("Q"):
                            bom_q_exact_map[to_code] = from_code
                            if len(from_code) > 1:
                                bom_r_exact_map[to_code] = "R" + from_code[1:]
                        elif from_code.startswith("R"):
                            bom_r_exact_map[to_code] = from_code
                            if len(from_code) > 1:
                                bom_q_exact_map[to_code] = "Q" + from_code[1:]

                bom["SALES_CODE5"] = bom["SALES_ITEM_CD"].str[:5]
                bom["FROM_CODE5"] = bom["FROM_ITEM_ID"].str[:5]
                bom = bom[bom["SALES_CODE5"].str.match(r"^[PQRSTU]\d{4}$", na=False)]
                bom = bom[bom["FROM_CODE5"].str.match(r"^[PQRSTU]\d{4}$", na=False)]
                if not bom.empty:
                    bom = bom.sort_values(["SALES_CODE5", "SEQ"], ascending=[True, True])
                    q_df = bom[bom["FROM_CODE5"].str.startswith("Q")].drop_duplicates(
                        subset=["SALES_CODE5"], keep="first"
                    )
                    r_df = bom[bom["FROM_CODE5"].str.startswith("R")].drop_duplicates(
                        subset=["SALES_CODE5"], keep="first"
                    )
                    bom_q_base_map = q_df.set_index("SALES_CODE5")["FROM_CODE5"].to_dict()
                    bom_r_base_map = r_df.set_index("SALES_CODE5")["FROM_CODE5"].to_dict()
                    for sales_code5, q_code5 in bom_q_base_map.items():
                        q_code5 = str(q_code5)
                        if sales_code5 not in bom_r_base_map and q_code5.startswith("Q") and len(q_code5) >= 5:
                            bom_r_base_map[sales_code5] = "R" + q_code5[1:5]

    # 4) 리드지 공정 맵
    if len(sheet_names) >= 3:
        leadji_sheet = next((s for s in sheet_names if s.replace(" ", "") == "리드지정보"), sheet_names[2])
        leadji = parse_sheet(leadji_sheet)
        if not leadji.empty:
            prod_col = "생산" if "생산" in leadji.columns else (leadji.columns[3] if len(leadji.columns) > 3 else None)
            q_col = "분리" if "분리" in leadji.columns else (leadji.columns[9] if len(leadji.columns) > 9 else None)
            r_col = "사출" if "사출" in leadji.columns else (leadji.columns[21] if len(leadji.columns) > 21 else None)
            if prod_col is not None and q_col is not None and r_col is not None:
                ldf = leadji[[prod_col, q_col, r_col]].copy()
                for col in [prod_col, q_col, r_col]:
                    ldf[col] = ldf[col].astype(str).str.strip()
                    ldf.loc[ldf[col].str.lower() == "nan", col] = ""
                ldf = ldf[ldf[prod_col].str.startswith("P")]
                if not ldf.empty:
                    ldf["코드5"] = ldf[prod_col].str[:5]
                    ldf = ldf[(ldf["코드5"] != "") & (ldf["코드5"].str.lower() != "nan")]

                    def normalize_to_code(code: str, prefix: str) -> str:
                        value = str(code).strip()
                        if not value or value.lower() == "nan":
                            return ""
                        if value.startswith(prefix):
                            return value
                        if value.startswith("P"):
                            return f"{prefix}{value[1:]}"
                        return value

                    ldf["Q정규"] = ldf[q_col].map(lambda x: normalize_to_code(x, "Q"))
                    ldf["R정규"] = ldf[r_col].map(lambda x: normalize_to_code(x, "R"))
                    q_df = ldf[ldf["Q정규"] != ""].drop_duplicates(subset=["코드5"], keep="first")
                    r_df = ldf[ldf["R정규"] != ""].drop_duplicates(subset=["코드5"], keep="first")
                    leadji_q_map = q_df.set_index("코드5")["Q정규"].to_dict()
                    leadji_r_map = r_df.set_index("코드5")["R정규"].to_dict()

    return (
        product_name_map,
        product_group_map,
        sheet2_group_map,
        r_ref_map,
        q_ref_map,
        r_name_map,
        bom_r_base_map,
        bom_q_base_map,
        bom_r_exact_map,
        bom_q_exact_map,
        leadji_r_map,
        leadji_q_map,
    )


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


def infer_numeric_like_series(series: pd.Series) -> bool:
    sample = series.astype(str).str.replace(",", "", regex=False).str.strip()
    sample = sample[~sample.str.lower().isin({"", "nan", "none"})].head(200)
    if sample.empty:
        return False
    numeric_mask = sample.str.fullmatch(r"[+-]?\d+(?:\.\d+)?").fillna(False)
    return bool(float(numeric_mask.mean()) >= 0.85)


def pick_fixed_column_width_px(column_name: str, max_length: int, numeric_like: bool) -> int:
    if numeric_like:
        return int(max(90, min(145, 24 + max_length * 7)))

    long_text_columns = {"제품명", "R코드 제품명", "리드지명", "제품명 예시"}
    medium_text_columns = {"품목코드", "R코드", "Q코드", "생산코드", "리드지코드", "P코드 예시"}
    status_columns = {"상태"}
    date_columns = {"납기일", "입고예상일자", "생산 최소 납기일", "최소납기일"}

    if column_name in long_text_columns:
        return int(max(240, min(380, 28 + max_length * 7)))
    if column_name in medium_text_columns:
        return int(max(120, min(170, 24 + max_length * 7)))
    if column_name in status_columns:
        return 96
    if column_name in date_columns:
        return 118
    return int(max(92, min(145, 24 + max_length * 7)))


def build_auto_column_config(
    df: pd.DataFrame, columns: list[str], source_df: pd.DataFrame | None = None
) -> dict[str, st.column_config.Column]:
    config: dict[str, st.column_config.Column] = {}
    for col in columns:
        if col not in df.columns:
            continue
        col_series = df[col].astype(str)
        length_series = col_series.map(len)
        p90_len = int(length_series.quantile(0.90)) if not length_series.empty else 0
        max_len = max(len(str(col)), p90_len)

        numeric_like = False
        if source_df is not None and col in source_df.columns:
            numeric_like = pd.api.types.is_numeric_dtype(source_df[col]) or infer_numeric_like_series(col_series)
        else:
            numeric_like = infer_numeric_like_series(col_series)

        width_px = pick_fixed_column_width_px(col, max_len, numeric_like)
        config[col] = st.column_config.Column(
            label=COLUMN_LABEL_ALIASES.get(col, col),
            width=width_px,
        )
    return config


def render_dashboard_kpi(label: str, value: str, variant: str = "stock") -> None:
    st.markdown(
        f"""
        <div class="ops-kpi-card {variant}">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def style_operational_table(display_df: pd.DataFrame, source_df: pd.DataFrame | None = None):
    if display_df.empty:
        return display_df.style

    source = source_df if source_df is not None else display_df
    styler = display_df.style
    numeric_cols: list[str] = []
    for col in display_df.columns:
        source_col = source[col] if col in source.columns else display_df[col]
        if pd.api.types.is_numeric_dtype(source_col) or infer_numeric_like_series(display_df[col]):
            numeric_cols.append(col)

    if numeric_cols:
        styler = styler.set_properties(subset=numeric_cols, **{"text-align": "right"})

    text_cols = [c for c in ["제품명", "R코드 제품명", "리드지명", "제품명 예시"] if c in display_df.columns]
    if text_cols:
        styler = styler.set_properties(subset=text_cols, **{"text-align": "left"})

    shortage_cols = [c for c in display_df.columns if "부족" in c]
    for col in shortage_cols:
        source_col = source[col] if col in source.columns else display_df[col]
        shortage_numeric = parse_mixed_numeric(source_col)
        shortage_style = shortage_numeric.map(
            lambda v: "color: #DC2626; font-weight: 850;" if pd.notna(v) and abs(float(v)) > 0 else "color: #6B7280;"
        )
        styler = styler.apply(lambda _: shortage_style, axis=0, subset=[col])

    if "상태" in display_df.columns:
        styler = styler.set_properties(subset=["상태"], **{"text-align": "center"})
        styler = styler.map(
            lambda v: (
                "background-color: #FEE2E2; color: #B91C1C; font-weight: 850;"
                if str(v).strip() == "부족"
                else "background-color: #FEF3C7; color: #92400E; font-weight: 850;"
                if str(v).strip() == "확인필요"
                else "background-color: #DCFCE7; color: #166534; font-weight: 850;"
                if str(v).strip() == "정상"
                else "background-color: #EEF2FF; color: #1A2B5E; font-weight: 800;"
            ),
            subset=["상태"],
        )

    return styler


def dataframe_to_excel_bytes(df: pd.DataFrame, sheet_name: str = "data") -> bytes:
    safe_sheet = re.sub(r"[:\\\\/?*\\[\\]]", "_", str(sheet_name)).strip()
    if not safe_sheet:
        safe_sheet = "data"
    safe_sheet = safe_sheet[:31]

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=safe_sheet)
    output.seek(0)
    return output.getvalue()


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


def normalize_warehouse_name(value: str) -> str:
    return re.sub(r"\s+", "", str(value)).strip().lower()


def find_warehouse_column(columns: list[str], candidates: list[str]) -> str | None:
    normalized_map = {normalize_warehouse_name(col): col for col in columns}
    for candidate in candidates:
        matched = normalized_map.get(normalize_warehouse_name(candidate))
        if matched is not None:
            return matched
    return None


def style_leadji_shortage_table(display_df: pd.DataFrame, source_df: pd.DataFrame):
    if display_df.empty:
        return display_df.style

    styler = display_df.style
    if "우선순위" in display_df.columns:
        styler = styler.set_properties(subset=["우선순위"], **{"text-align": "center"})
        styler = styler.map(
            lambda v: (
                "background-color: #FEE2E2; color: #B91C1C; font-weight: 800;"
                if str(v).strip() == "긴급"
                else "background-color: #FEF3C7; color: #92400E; font-weight: 800;"
                if str(v).strip() == "확인필요"
                else "background-color: #F1F5F9; color: #475569; font-weight: 700;"
            ),
            subset=["우선순위"],
        )
    if "리드지부족" in display_df.columns:
        styler = styler.set_properties(subset=["리드지부족"], **{"text-align": "center"})
        styler = styler.map(
            lambda v: "color: #d00000; font-weight: 700;" if str(v).strip() not in {"", "-", "nan", "None"} else "",
            subset=["리드지부족"],
        )
    if "리드지부족수량" in display_df.columns and "리드지부족수량" in source_df.columns:
        shortage_numeric = parse_mixed_numeric(source_df["리드지부족수량"])
        shortage_style = shortage_numeric.map(
            lambda v: "color: #DC2626; font-weight: 850;" if pd.notna(v) and v < 0 else "color: #6B7280;"
        )
        styler = styler.apply(lambda _: shortage_style, axis=0, subset=["리드지부족수량"])
    if "상태" in display_df.columns:
        styler = styler.set_properties(subset=["상태"], **{"text-align": "center"})
        styler = styler.map(
            lambda v: (
                "background-color: #FEE2E2; color: #B91C1C; font-weight: 800;"
                if str(v).strip() == "입고일 미확인"
                else "background-color: #FEF3C7; color: #92400E; font-weight: 800;"
                if str(v).strip() == "발주부족"
                else "background-color: #EEF2FF; color: #1A2B5E; font-weight: 800;"
                if str(v).strip() in {"입고 예정", "입고 예정+의뢰"}
                else "background-color: #D1FAE5; color: #047857; font-weight: 800;"
                if str(v).strip() == "구매의뢰"
                else "background-color: #F1F5F9; color: #475569; font-weight: 700;"
            ),
            subset=["상태"],
        )
    return styler


def add_rq_group_columns(df: pd.DataFrame) -> pd.DataFrame:
    enriched = df.copy()
    if "R코드" not in enriched.columns:
        enriched["R코드"] = enriched["품목코드"].map(lambda x: map_demand_code_to_process_code(x, "R"))
    if "Q코드" not in enriched.columns:
        enriched["Q코드"] = enriched["품목코드"].map(lambda x: map_demand_code_to_process_code(x, "Q"))
    if "R코드 제품명" not in enriched.columns:
        enriched["R코드 제품명"] = enriched.get("제품명", "-")
    enriched["R코드5"] = enriched["R코드"].astype(str).str[:5]
    enriched["Q코드5"] = enriched["Q코드"].astype(str).str[:5]
    enriched["P코드5"] = enriched["품목코드"].astype(str).str[:5]
    enriched["RQ그룹"] = enriched["R코드"].astype(str) + " | " + enriched["Q코드"].astype(str)
    return enriched


def build_rcode_summary(df: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "R코드",
        "R코드 제품명",
        "사출 납기일",
        "사출 생산 필요수량 합계",
        "사출창고 합계",
        "분리창고 합계",
        "공정재고 합계",
    ]
    if df.empty:
        return pd.DataFrame(columns=columns)

    r_df = df.copy()
    r_df["R코드"] = r_df["R코드"].astype(str).str.strip()
    r_df = r_df[(r_df["R코드"] != "") & (r_df["R코드"].str.lower() != "nan")]
    if "사출생산필요수량" in r_df.columns:
        r_df["사출생산필요수량"] = parse_mixed_numeric(r_df["사출생산필요수량"])
        r_df = r_df[r_df["사출생산필요수량"] > 0]
    if r_df.empty:
        return pd.DataFrame(columns=columns)

    grouped = (
        r_df.groupby("R코드", as_index=False)
        .agg(
            {
                "R코드 제품명": lambda s: summarize_unique(s, head_count=1),
                "사출납기일": "min",
                "사출생산필요수량": "sum",
                "사출창고": "sum",
                "분리창고": "sum",
                "공정재고 합계": "sum",
            }
        )
        .rename(
            columns={
                "사출납기일": "사출 납기일",
                "사출생산필요수량": "사출 생산 필요수량 합계",
                "사출창고": "사출창고 합계",
                "분리창고": "분리창고 합계",
            }
        )
    )

    grouped = grouped.sort_values(["사출 생산 필요수량 합계", "R코드"], ascending=[False, True])
    return grouped[columns]


def build_rq_group_summary(df: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "R코드",
        "Q코드",
        "R코드 제품명",
        "P코드5 수",
        "제품명 예시",
        "P코드 예시",
        "부족수량 합계",
        "사출 생산 필요수량 합계",
        "사출창고 합계",
        "분리창고 합계",
        "공정재고 합계",
        "사출 부족수량",
    ]
    if df.empty:
        return pd.DataFrame(columns=columns)
    if "부족수량" not in df.columns:
        return pd.DataFrame(columns=columns)

    df = df.copy()
    df["부족수량"] = parse_mixed_numeric(df["부족수량"])
    if "사출생산필요수량" in df.columns:
        df["사출생산필요수량"] = parse_mixed_numeric(df["사출생산필요수량"])
    else:
        df["사출생산필요수량"] = 0
    df = df[(df["부족수량"] > 0) | (df["사출생산필요수량"] > 0)]
    if df.empty:
        return pd.DataFrame(columns=columns)

    grouped = (
        df.groupby(["R코드5", "Q코드5"], as_index=False)
        .agg(
            {
                "R코드 제품명": lambda s: summarize_unique(s, head_count=1),
                "제품명": lambda s: summarize_unique(s, head_count=3),
                "품목코드": lambda s: summarize_unique(s, head_count=5),
                "부족수량": "sum",
                "사출생산필요수량": "sum",
                "사출창고": "sum",
                "분리창고": "sum",
                "공정재고 합계": "sum",
            }
        )
        .rename(
            columns={
                "R코드5": "R코드",
                "Q코드5": "Q코드",
                "제품명": "제품명 예시",
                "품목코드": "P코드 예시",
                "부족수량": "부족수량 합계",
                "사출생산필요수량": "사출 생산 필요수량 합계",
                "사출창고": "사출창고 합계",
                "분리창고": "분리창고 합계",
            }
        )
    )
    p_count = df.groupby(["R코드5", "Q코드5"])["P코드5"].nunique().rename("P코드5 수").reset_index()
    p_count = p_count.rename(columns={"R코드5": "R코드", "Q코드5": "Q코드"})
    grouped = grouped.merge(p_count, on=["R코드", "Q코드"], how="left")
    grouped["사출 부족수량"] = (grouped["사출 생산 필요수량 합계"] - grouped["사출창고 합계"]).clip(lower=0)
    grouped = grouped.sort_values(["부족수량 합계", "P코드5 수"], ascending=[False, False])
    return grouped[columns]


def build_initial_injection_summary(df: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "이니셜",
        "거래처 수",
        "품목코드 수",
        "부족수량 합계",
        "사출 생산 필요수량 합계",
        "사출창고 합계",
        "사출 부족수량",
    ]
    if df.empty or "이니셜" not in df.columns or "품목코드" not in df.columns:
        return pd.DataFrame(columns=columns)

    base = df.copy()
    if "거래처" not in base.columns:
        base["거래처"] = "-"
    if "부족수량" not in base.columns:
        base["부족수량"] = 0
    if "사출생산필요수량" not in base.columns:
        base["사출생산필요수량"] = 0
    if "사출창고" not in base.columns:
        base["사출창고"] = 0

    base["이니셜"] = base["이니셜"].astype(str).str.strip()
    base["이니셜"] = base["이니셜"].replace({"": "(미지정)", "nan": "(미지정)", "None": "(미지정)"})
    base["부족수량"] = parse_mixed_numeric(base["부족수량"])
    base["사출생산필요수량"] = parse_mixed_numeric(base["사출생산필요수량"])
    base["사출창고"] = parse_mixed_numeric(base["사출창고"])
    base = base[(base["부족수량"] > 0) | (base["사출생산필요수량"] > 0)]
    if base.empty:
        return pd.DataFrame(columns=columns)

    # 사출창고는 품목별 고정 재고 성격이라, 이니셜+품목 기준 최대값으로 중복 집계를 완화한다.
    item_level = (
        base.groupby(["이니셜", "품목코드"], as_index=False)
        .agg(
            {
                "부족수량": "sum",
                "사출생산필요수량": "sum",
                "사출창고": "max",
            }
        )
    )

    summary = (
        item_level.groupby("이니셜", as_index=False)
        .agg(
            {
                "품목코드": "nunique",
                "부족수량": "sum",
                "사출생산필요수량": "sum",
                "사출창고": "sum",
            }
        )
        .rename(
            columns={
                "품목코드": "품목코드 수",
                "부족수량": "부족수량 합계",
                "사출생산필요수량": "사출 생산 필요수량 합계",
                "사출창고": "사출창고 합계",
            }
        )
    )

    customer_count = (
        base.groupby("이니셜", as_index=False)["거래처"]
        .nunique()
        .rename(columns={"거래처": "거래처 수"})
    )
    summary = summary.merge(customer_count, on="이니셜", how="left")
    summary["사출 부족수량"] = (summary["사출 생산 필요수량 합계"] - summary["사출창고 합계"]).clip(lower=0)
    summary = summary[columns].sort_values(
        ["사출 부족수량", "사출 생산 필요수량 합계", "부족수량 합계", "이니셜"],
        ascending=[False, False, False, True],
    )
    return summary


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
    if "부족수량" not in df.columns:
        return pd.DataFrame(columns=columns)

    q_df = df.copy()
    q_df["부족수량"] = parse_mixed_numeric(q_df["부족수량"])
    q_df = q_df[q_df["부족수량"] > 0]
    if q_df.empty:
        return pd.DataFrame(columns=columns)

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


def build_summary_group_totals_with_safe_split(df: pd.DataFrame) -> pd.DataFrame:
    columns = ["분류별요약", "오더 부족수량", "안전재고 부족수량", "총수량"]
    if df.empty or "부족수량" not in df.columns:
        return pd.DataFrame(columns=columns)

    base = df.copy()
    if "분류별요약" not in base.columns:
        base["분류별요약"] = "(미분류)"
    if "이니셜" not in base.columns:
        base["이니셜"] = ""

    group_label = base["분류별요약"].astype(str).str.strip()
    base["분류별요약"] = group_label.replace({"": "(미분류)", "nan": "(미분류)", "None": "(미분류)"})
    include_qty = base.groupby("분류별요약", as_index=False)["부족수량"].sum().rename(columns={"부족수량": "안전 포함"})
    exclude_qty = (
        base[~base["이니셜"].astype(str).str.contains("안전", na=False)]
        .groupby("분류별요약", as_index=False)["부족수량"]
        .sum()
        .rename(columns={"부족수량": "안전 미포함"})
    )

    grouped = include_qty.merge(exclude_qty, on="분류별요약", how="left").fillna(0)
    grouped["오더 부족수량"] = grouped["안전 미포함"]
    grouped["안전재고 부족수량"] = grouped["안전 포함"] - grouped["안전 미포함"]
    grouped["총수량"] = grouped["오더 부족수량"] + grouped["안전재고 부족수량"]
    grouped = grouped[columns].sort_values("총수량", ascending=False)

    total_row = pd.DataFrame(
        [
            {
                "분류별요약": "전체",
                "오더 부족수량": grouped["오더 부족수량"].sum(),
                "안전재고 부족수량": grouped["안전재고 부족수량"].sum(),
                "총수량": grouped["총수량"].sum(),
            }
        ]
    )
    return pd.concat([total_row, grouped], ignore_index=True)


@st.cache_data(show_spinner=False, persist="disk")
def load_data(refresh_key: str, base_dir_str: str | None = None) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    _ = refresh_key
    data_base_dir = Path(base_dir_str) if base_dir_str else BASE_DIR
    inv_path, dem_path = find_excel_files(data_base_dir)
    reference_refresh_key = build_reference_refresh_key(data_base_dir)
    (
        product_name_map,
        product_group_map,
        sheet2_group_map,
        r_ref_map,
        q_ref_map,
        r_name_map,
        bom_r_base_map,
        bom_q_base_map,
        bom_r_exact_map,
        bom_q_exact_map,
        leadji_r_map,
        leadji_q_map,
    ) = load_reference_maps_bundle(data_base_dir, reference_refresh_key)
    process_code_map, warehouse_qty_col_indices, qty_col_indices, total_qty_col_indices, process_qty_col_indices = (
        extract_demand_header_info(dem_path)
    )

    inv = pd.read_excel(inv_path, sheet_name=0)
    dem = pd.read_excel(dem_path, sheet_name=0, header=1)

    dem.columns = [str(c).strip() for c in dem.columns]

    site_col = pick_first_existing_column(
        dem.columns.tolist(),
        ["설비 사이트 코드", "설비사이트코드", "사이트코드"],
    )
    customer_col = pick_first_existing_column(
        dem.columns.tolist(),
        ["고객 이름", "고객이름", "거래처"],
    )
    initial_col = pick_first_existing_column(
        dem.columns.tolist(),
        ["이니셜"],
    )
    demand_item_col = pick_first_existing_column(
        dem.columns.tolist(),
        ["제품 코드", "제품코드", "품목코드", "ITEM_ID"],
    )
    demand_name_col = pick_first_existing_column(
        dem.columns.tolist(),
        ["수요 제품 이름", "수요제품이름", "제품명"],
    )

    site_series = (
        dem[site_col].astype(str).str.strip()
        if site_col is not None
        else (dem.iloc[:, 0].astype(str).str.strip() if dem.shape[1] > 0 else pd.Series("", index=dem.index))
    )
    customer_series = (
        dem[customer_col].astype(str).str.strip()
        if customer_col is not None
        else (dem.iloc[:, 1].astype(str).str.strip() if dem.shape[1] > 1 else pd.Series("", index=dem.index))
    )
    initial_series = (
        dem[initial_col].astype(str).str.strip()
        if initial_col is not None
        else (dem.iloc[:, 2].astype(str).str.strip() if dem.shape[1] > 2 else pd.Series("", index=dem.index))
    )
    item_series = (
        dem[demand_item_col].astype(str).str.strip()
        if demand_item_col is not None
        else (dem.iloc[:, 3].astype(str).str.strip() if dem.shape[1] > 3 else pd.Series("", index=dem.index))
    )
    name_series = (
        dem[demand_name_col].astype(str).str.strip()
        if demand_name_col is not None
        else (dem.iloc[:, 4].astype(str).str.strip() if dem.shape[1] > 4 else pd.Series("", index=dem.index))
    )

    # 기준1) 생산 현황: 누수/규격검사 생산수량 + 납기일
    leak_qty_idx = warehouse_qty_col_indices.get("누수규격검사 창고")
    leak_due_idx = leak_qty_idx + 1 if leak_qty_idx is not None and (leak_qty_idx + 1) < dem.shape[1] else None
    if leak_qty_idx is not None:
        shortage_qty = parse_mixed_numeric(dem.iloc[:, leak_qty_idx])
    elif total_qty_col_indices:
        total_qty_col = dem.columns[total_qty_col_indices[-1]]
        shortage_qty = parse_mixed_numeric(dem[total_qty_col])
    else:
        qty_cols = [dem.columns[i] for i in qty_col_indices]
        if not qty_cols:
            raise ValueError("수요 파일에서 '생산 수량' 컬럼을 찾지 못했습니다.")
        shortage_qty = dem[qty_cols].apply(parse_mixed_numeric).fillna(0).sum(axis=1)

    if leak_due_idx is not None:
        leak_due_date = parse_mixed_excel_date(dem.iloc[:, leak_due_idx])
    else:
        leak_due_date = pd.Series(pd.NaT, index=dem.index, dtype="datetime64[ns]")

    leadji_qty_idx: int | None = None
    for process_label, idx in process_qty_col_indices.items():
        if "[45]" in process_label and ("하이드레이션" in process_label or "전면검사" in process_label):
            leadji_qty_idx = idx
            break
    if leadji_qty_idx is None:
        for process_label, idx in process_qty_col_indices.items():
            if "하이드레이션" in process_label or "전면검사" in process_label:
                leadji_qty_idx = idx
                break

    if leadji_qty_idx is not None and 0 <= leadji_qty_idx < dem.shape[1]:
        leadji_required_qty = parse_mixed_numeric(dem.iloc[:, leadji_qty_idx])
        leadji_due_idx = leadji_qty_idx + 1 if (leadji_qty_idx + 1) < dem.shape[1] else None
        if leadji_due_idx is not None:
            leadji_due_date = parse_mixed_excel_date(dem.iloc[:, leadji_due_idx])
        else:
            leadji_due_date = pd.Series(pd.NaT, index=dem.index, dtype="datetime64[ns]")
    else:
        leadji_required_qty = shortage_qty
        leadji_due_date = leak_due_date

    # 기준2) 사출 생산 현황: [10]사출조립 생산수량 + 해당 납기일
    # 파일 구조가 바뀌어도 공정 헤더 기준으로 우선 선택한다.
    selected_qty_idx: int | None = warehouse_qty_col_indices.get("사출창고")
    if selected_qty_idx is not None and 0 <= selected_qty_idx < dem.shape[1]:
        inj_qty = parse_mixed_numeric(dem.iloc[:, selected_qty_idx])
    else:
        inj_qty_col = pick_first_existing_column(
            dem.columns.tolist(),
            ["사출조립 생산수량", "사출조립생산수량", "사출조립 생산 수량"],
        )
        if inj_qty_col is not None:
            inj_qty = parse_mixed_numeric(dem[inj_qty_col])
            selected_qty_idx = dem.columns.get_loc(inj_qty_col)
        elif dem.shape[1] > 5:
            inj_qty = parse_mixed_numeric(dem.iloc[:, 5])
            selected_qty_idx = 5
        else:
            raise ValueError("수요 파일 사출조립 생산수량 컬럼을 찾지 못했습니다.")

    inj_due_idx = selected_qty_idx + 1 if selected_qty_idx is not None and (selected_qty_idx + 1) < dem.shape[1] else None
    if inj_due_idx is not None:
        inj_due_date = parse_mixed_excel_date(dem.iloc[:, inj_due_idx])
    else:
        inj_due_date = pd.Series(pd.NaT, index=dem.index, dtype="datetime64[ns]")

    inv_df = build_inventory_df(inv)

    dem_df = pd.DataFrame(
        {
            "사이트코드": site_series,
            "거래처": customer_series,
            "이니셜": initial_series,
            "품목코드": item_series,
            "제품명": name_series,
            "납기일": leak_due_date,
            "사출납기일": inj_due_date,
            LEADJI_REQUIRED_DUE_COL: leadji_due_date,
            "생산수량": shortage_qty,
            "사출생산필요수량": inj_qty,
            LEADJI_REQUIRED_QTY_COL: leadji_required_qty,
        }
    )

    is_summary = (
        (dem_df["사이트코드"] == "총합계")
        | (dem_df["거래처"] == "총합계")
        | (dem_df["이니셜"] == "총합계")
        | (dem_df["품목코드"] == "총합계")
    )
    dem_df = dem_df[~is_summary]
    dem_df = dem_df[(dem_df["사이트코드"] != "") & (dem_df["사이트코드"].str.lower() != "nan")]
    dem_df = dem_df[(dem_df["품목코드"] != "") & (dem_df["품목코드"].str.lower() != "nan")]
    dem_df = dem_df[dem_df["품목코드"].astype(str).str.upper().str.startswith(("P", "Q", "R"))]
    dem_df = dem_df[
        (dem_df["생산수량"] > 0) | (dem_df["사출생산필요수량"] > 0) | (dem_df[LEADJI_REQUIRED_QTY_COL] > 0)
    ]
    dem_df["제품명"] = dem_df["제품명"].replace({"nan": "", "None": ""})

    grouped_demand = (
        dem_df.groupby(
            ["사이트코드", "이니셜", "거래처", "품목코드", "납기일", "사출납기일", LEADJI_REQUIRED_DUE_COL],
            as_index=False,
            dropna=False,
        )
        .agg(
            {
                "생산수량": "sum",
                "사출생산필요수량": "sum",
                LEADJI_REQUIRED_QTY_COL: "sum",
                "제품명": lambda s: next((v for v in s if str(v).strip() and str(v).strip().lower() != "nan"), "-"),
            }
        )
        .rename(columns={"생산수량": "부족수량"})
    )
    grouped_demand["코드5"] = grouped_demand["품목코드"].str[:5]
    grouped_demand["제품명"] = grouped_demand["코드5"].map(product_name_map).fillna(grouped_demand["제품명"])
    grouped_demand["제품명"] = grouped_demand["제품명"].replace({"": "-", "nan": "-", "None": "-"}).fillna("-")

    inferred_r = grouped_demand["품목코드"].map(lambda x: map_demand_code_to_process_code(x, "R"))
    inferred_q = grouped_demand["품목코드"].map(lambda x: map_demand_code_to_process_code(x, "Q"))
    mapped_r_base = grouped_demand["코드5"].map(leadji_r_map).fillna(grouped_demand["코드5"].map(r_ref_map))
    mapped_q_base = grouped_demand["코드5"].map(leadji_q_map).fillna(grouped_demand["코드5"].map(q_ref_map))

    merged_r = pd.Series(
        [merge_mapped_base_code(inferred, mapped, "R") for inferred, mapped in zip(inferred_r, mapped_r_base)],
        index=grouped_demand.index,
    )
    merged_q = pd.Series(
        [merge_mapped_base_code(inferred, mapped, "Q") for inferred, mapped in zip(inferred_q, mapped_q_base)],
        index=grouped_demand.index,
    )

    # P코드는 수요정보 품목코드 기준(P->R/Q)을 우선 유지한다.
    # 리드지/분류 매핑 코드는 비-P 항목에서만 적용한다.
    item_prefix = grouped_demand["품목코드"].astype(str).str.upper().str[:1]
    use_mapped_mask = item_prefix != "P"

    grouped_demand["R코드"] = inferred_r
    grouped_demand["Q코드"] = inferred_q
    grouped_demand.loc[use_mapped_mask, "R코드"] = merged_r.loc[use_mapped_mask]
    grouped_demand.loc[use_mapped_mask, "Q코드"] = merged_q.loc[use_mapped_mask]

    # BOM exact mapping has the highest priority when TO_ITEM_ID matches the demand item code.
    if bom_r_exact_map or bom_q_exact_map:
        bom_exact_r = grouped_demand["품목코드"].map(bom_r_exact_map)
        bom_exact_q = grouped_demand["품목코드"].map(bom_q_exact_map)
        exact_r_mask = bom_exact_r.notna() & (bom_exact_r.astype(str).str.strip() != "")
        exact_q_mask = bom_exact_q.notna() & (bom_exact_q.astype(str).str.strip() != "")
        grouped_demand.loc[exact_r_mask, "R코드"] = bom_exact_r.loc[exact_r_mask]
        grouped_demand.loc[exact_q_mask, "Q코드"] = bom_exact_q.loc[exact_q_mask]

    # BOM fallback: only for rows still not mappable as valid R/Q codes.
    if bom_r_base_map or bom_q_base_map:
        bom_r_base = grouped_demand["코드5"].map(bom_r_base_map)
        bom_q_base = grouped_demand["코드5"].map(bom_q_base_map)

        bom_merged_r = pd.Series(
            [merge_mapped_base_code(inferred, mapped, "R") for inferred, mapped in zip(inferred_r, bom_r_base)],
            index=grouped_demand.index,
        )
        bom_merged_q = pd.Series(
            [merge_mapped_base_code(inferred, mapped, "Q") for inferred, mapped in zip(inferred_q, bom_q_base)],
            index=grouped_demand.index,
        )

        r_norm = grouped_demand["R코드"].astype(str).str.strip()
        q_norm = grouped_demand["Q코드"].astype(str).str.strip()
        invalid_r_mask = (r_norm == "") | (r_norm.str.lower() == "nan") | (~r_norm.str.startswith("R"))
        invalid_q_mask = (q_norm == "") | (q_norm.str.lower() == "nan") | (~q_norm.str.startswith("Q"))

        grouped_demand.loc[invalid_r_mask, "R코드"] = bom_merged_r.loc[invalid_r_mask]
        grouped_demand.loc[invalid_q_mask, "Q코드"] = bom_merged_q.loc[invalid_q_mask]

    grouped_demand["R코드5"] = grouped_demand["R코드"].astype(str).str[:5]
    grouped_demand["R코드 제품명"] = grouped_demand["R코드5"].map(r_name_map)
    grouped_demand["R코드 제품명"] = grouped_demand["R코드 제품명"].fillna(grouped_demand["제품명"])
    grouped_demand["R코드 제품명"] = grouped_demand["R코드 제품명"].fillna(grouped_demand["R코드5"])
    grouped_demand["R코드 제품명"] = grouped_demand["R코드 제품명"].replace({"": "-", "nan": "-", "None": "-"}).fillna("-")
    grouped_demand["분류별요약"] = grouped_demand["코드5"].map(product_group_map).fillna("기타")
    grouped_demand["시트분류"] = grouped_demand["코드5"].map(sheet2_group_map).fillna("기타 해외")
    grouped_demand = grouped_demand.drop(columns=["코드5", "R코드5"])

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
    rq_by_p = grouped_demand.drop_duplicates(subset=["품목코드"], keep="first").set_index("품목코드")[["R코드", "Q코드"]]
    r_by_p = {
        item_code: resolve_process_code_for_stock(stock_lookup["사출창고"], process_code)
        for item_code, process_code in rq_by_p["R코드"].to_dict().items()
    }
    q_by_p = {
        item_code: resolve_process_code_for_stock(stock_lookup["분리창고"], process_code)
        for item_code, process_code in rq_by_p["Q코드"].to_dict().items()
    }

    grouped_demand["R코드"] = grouped_demand["품목코드"].map(
        lambda x: r_by_p.get(x, map_demand_code_to_process_code(x, "R"))
    )
    grouped_demand["Q코드"] = grouped_demand["품목코드"].map(
        lambda x: q_by_p.get(x, map_demand_code_to_process_code(x, "Q"))
    )
    grouped_demand["R코드5"] = grouped_demand["R코드"].astype(str).str[:5]
    grouped_demand["R코드 제품명"] = grouped_demand["R코드5"].map(r_name_map)
    grouped_demand["R코드 제품명"] = grouped_demand["R코드 제품명"].fillna(grouped_demand["제품명"])
    grouped_demand["R코드 제품명"] = grouped_demand["R코드 제품명"].fillna(grouped_demand["R코드5"])
    grouped_demand["R코드 제품명"] = grouped_demand["R코드 제품명"].replace({"": "-", "nan": "-", "None": "-"}).fillna("-")

    code_stock["사출창고"] = code_stock["품목코드"].map(
        lambda x: lookup_stock_qty(stock_lookup["사출창고"], r_by_p.get(x, map_demand_code_to_process_code(x, "R")))
    )
    code_stock["분리창고"] = code_stock["품목코드"].map(
        lambda x: lookup_stock_qty(stock_lookup["분리창고"], q_by_p.get(x, map_demand_code_to_process_code(x, "Q")))
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

    # 분류 필터 정합성 보정:
    # P코드는 코드5(Pxxxx) 기준 매핑을 그대로 사용하고,
    # R/Q/U 등 비-P코드는 같은 R코드5를 공유하는 P코드의 분류를 이어받는다.
    result["코드5"] = result["품목코드"].astype(str).str[:5]
    result["분류별요약"] = result["코드5"].map(product_group_map)
    result["시트분류"] = result["코드5"].map(sheet2_group_map)
    result["R코드5"] = result["R코드"].astype(str).str[:5]

    item_prefix = result["품목코드"].astype(str).str.upper().str[:1]
    p_scope = result[(item_prefix == "P") & result["R코드5"].str.startswith("R", na=False)].copy()
    if not p_scope.empty:
        p_scope["부족수량_num"] = parse_mixed_numeric(p_scope["부족수량"])
        p_scope = p_scope.sort_values(["부족수량_num", "품목코드"], ascending=[False, True])

        p_sheet_scope = p_scope[p_scope["시트분류"].notna()].copy()
        p_sheet_scope["시트분류"] = p_sheet_scope["시트분류"].astype(str).str.strip()
        p_sheet_scope = p_sheet_scope[
            (p_sheet_scope["시트분류"] != "")
            & (p_sheet_scope["시트분류"].str.lower() != "nan")
            & (p_sheet_scope["시트분류"].str.lower() != "none")
        ]
        r_to_sheet = p_sheet_scope.drop_duplicates(subset=["R코드5"], keep="first").set_index("R코드5")["시트분류"].to_dict()

        p_group_scope = p_scope[p_scope["분류별요약"].notna()].copy()
        p_group_scope["분류별요약"] = p_group_scope["분류별요약"].astype(str).str.strip()
        p_group_scope = p_group_scope[
            (p_group_scope["분류별요약"] != "")
            & (p_group_scope["분류별요약"].str.lower() != "nan")
            & (p_group_scope["분류별요약"].str.lower() != "none")
        ]
        r_to_group = (
            p_group_scope.drop_duplicates(subset=["R코드5"], keep="first").set_index("R코드5")["분류별요약"].to_dict()
        )
    else:
        r_to_sheet = {}
        r_to_group = {}

    non_p_mask = item_prefix != "P"
    result.loc[non_p_mask, "시트분류"] = result.loc[non_p_mask, "시트분류"].fillna(
        result.loc[non_p_mask, "R코드5"].map(r_to_sheet)
    )
    result.loc[non_p_mask, "분류별요약"] = result.loc[non_p_mask, "분류별요약"].fillna(
        result.loc[non_p_mask, "R코드5"].map(r_to_group)
    )

    result["시트분류"] = result["시트분류"].astype(str).str.strip()
    result["분류별요약"] = result["분류별요약"].astype(str).str.strip()
    result.loc[result["시트분류"].str.lower().isin({"", "nan", "none"}), "시트분류"] = "기타 해외"
    result.loc[result["분류별요약"].str.lower().isin({"", "nan", "none"}), "분류별요약"] = "기타"
    result = result.drop(columns=["코드5"], errors="ignore")

    result["파워"] = result["품목코드"].map(extract_power_from_code)
    result["납기일"] = pd.to_datetime(result["납기일"], errors="coerce").dt.strftime("%Y-%m-%d")
    result["납기일"] = result["납기일"].fillna("-")
    if "사출납기일" in result.columns:
        result["사출납기일"] = pd.to_datetime(result["사출납기일"], errors="coerce").dt.strftime("%Y-%m-%d")
        result["사출납기일"] = result["사출납기일"].fillna("-")

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
                "리드지정보 우선, 없으면 분류정보, 그래도 없으면 P코드->R코드 유추 (BUL1/BUL2는 BUL로 보정)",
                "리드지정보 우선, 없으면 분류정보, 그래도 없으면 P코드->Q코드 유추 (BUL1/BUL2는 BUL로 보정)",
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
    with st.sidebar:
        st.markdown('<div class="sidebar-section-title">필터</div>', unsafe_allow_html=True)
        st.caption(f"업데이트: {updated_at}")
        st.caption("기본 적용: 전체 수요")

        scope_df = df.copy()
        if "사이트코드" not in scope_df.columns:
            scope_df["사이트코드"] = "(미지정)"
        site_label = scope_df["사이트코드"].astype(str).str.strip()
        scope_df["사이트코드"] = site_label.replace({"": "(미지정)", "nan": "(미지정)", "None": "(미지정)"})

        site_sum_map = (
            scope_df.groupby("사이트코드", as_index=True)["부족수량"].sum().sort_values(ascending=False).to_dict()
        )
        site_options = ["전체"] + list(site_sum_map.keys())
        site_count_map = {"전체": float(scope_df["부족수량"].sum()), **site_sum_map}
        selected_site_option = st.pills(
            "사이트코드",
            options=site_options,
            default="전체",
            key="flt_site_pills",
            format_func=lambda x: format_pill_label(x, site_count_map),
        )
        if selected_site_option and selected_site_option != "전체":
            scope_df = scope_df[scope_df["사이트코드"] == selected_site_option]

        st.divider()
        unified_query = st.text_input(
            "통합 검색",
            value="",
            key="flt_unified_query",
            placeholder="사이트/거래처/품목/RQ 코드",
            help="콤마(,)로 여러 키워드를 입력하면 OR 조건으로 검색합니다.",
        ).strip()

        only_with_stock = st.checkbox("공정재고만", value=False, key="flt_only_stock")
        exclude_safe_initial = st.checkbox("안전 이니셜 제외", value=False, key="flt_exclude_safe_initial")
        only_same_rq_group = st.checkbox("동일 RQ그룹만(R5/Q5, P5종류2+)", value=False, key="flt_only_same_rq_group")

        sheet_sum_map = (
            scope_df.groupby("시트분류", as_index=True)["부족수량"].sum().sort_values(ascending=False).to_dict()
        )
        summary_sum_map = (
            scope_df.groupby("분류별요약", as_index=True)["부족수량"].sum().sort_values(ascending=False).to_dict()
        )

        sheet_options = ["전체"] + list(sheet_sum_map.keys())
        summary_options = ["전체"] + list(summary_sum_map.keys())
        sheet_count_map = {"전체": float(scope_df["부족수량"].sum()), **sheet_sum_map}
        summary_count_map = {"전체": float(scope_df["부족수량"].sum()), **summary_sum_map}

        st.divider()
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

    base_filtered = scope_df.copy()
    search_cols = [c for c in ["사이트코드", "이니셜", "거래처", "품목코드", "제품명", "R코드 제품명", "R코드", "Q코드"] if c in base_filtered.columns]
    base_filtered = filter_with_terms_any(base_filtered, search_cols, unified_query)
    if exclude_safe_initial:
        base_filtered = base_filtered[~base_filtered["이니셜"].astype(str).str.contains("안전", na=False)]
    if selected_sheet_option and selected_sheet_option != "전체":
        base_filtered = base_filtered[base_filtered["시트분류"] == selected_sheet_option]
    if selected_summary_option and selected_summary_option != "전체":
        base_filtered = base_filtered[base_filtered["분류별요약"] == selected_summary_option]
    if only_same_rq_group and {"R코드5", "Q코드5", "P코드5"}.issubset(base_filtered.columns):
        p_count_per_group = base_filtered.groupby(["R코드5", "Q코드5"])["P코드5"].transform("nunique")
        base_filtered = base_filtered[p_count_per_group >= 2]
    if only_with_stock:
        base_filtered = base_filtered[base_filtered["공정재고 합계"] > 0]

    filtered = base_filtered.copy()
    return filtered


@st.cache_data(show_spinner=False, persist="disk")
def load_leadji_data(refresh_key: str, base_dir_str: str | None = None) -> tuple[pd.DataFrame, pd.DataFrame]:
    _ = refresh_key
    data_base_dir = Path(base_dir_str) if base_dir_str else BASE_DIR
    ref_path = find_product_name_reference_file(data_base_dir)
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
                leadji_info[col] = parse_mixed_numeric(leadji_info[col])

    if not leadji_stock.empty:
        leadji_stock.columns = [str(c).strip() for c in leadji_stock.columns]
        for col in ["기초", "입고", "출고", "재고", "검사대기"]:
            if col in leadji_stock.columns:
                leadji_stock[col] = parse_mixed_numeric(leadji_stock[col])

    return leadji_info, leadji_stock


@st.cache_data(show_spinner=False, persist="disk")
def load_leadji_order_data(refresh_key: str, base_dir_str: str | None = None) -> pd.DataFrame:
    _ = refresh_key
    data_base_dir = Path(base_dir_str) if base_dir_str else BASE_DIR
    order_path = find_leadji_order_status_file(data_base_dir)

    empty = pd.DataFrame(
        columns=[
            "리드지코드",
            "리드지명",
            "발주수량",
            "입고예상일자",
            "입고예상일자_dt",
            "구매발주수량",
            "구매의뢰수량",
        ]
    )
    if order_path is None:
        return empty

    try:
        sheet_names = pd.ExcelFile(order_path).sheet_names
    except Exception:
        return empty

    normalized_sheet_names = {str(name).replace(" ", ""): name for name in sheet_names}
    purchase_order_sheet = normalized_sheet_names.get("구매발주현황", sheet_names[0] if sheet_names else 0)
    purchase_request_sheet = normalized_sheet_names.get("구매의뢰현황")

    summaries: list[pd.DataFrame] = []

    def first_nonempty_text(series: pd.Series) -> str:
        text = series.astype(str).str.strip()
        text = text[(text != "") & (text.str.lower() != "nan") & (text.str.lower() != "none")]
        return text.iloc[0] if not text.empty else "-"

    try:
        # 구매발주현황 기준: J열(품목코드), O열(미입고수량), X열(납기일자).
        # 화면의 "발주수량"은 현재 남아 있는 입고 예정 수량이므로 미입고수량을 사용한다.
        raw_order = pd.read_excel(order_path, sheet_name=purchase_order_sheet, header=0, usecols=[9, 11, 14, 23])
    except Exception:
        raw_order = pd.DataFrame()

    if not raw_order.empty:
        raw_order.columns = ["리드지코드_raw", "리드지명", "구매발주수량", "입고예상일자_raw"]
        raw_order["리드지코드"] = raw_order["리드지코드_raw"].map(normalize_leadji_code_key)
        raw_order = raw_order[raw_order["리드지코드"].str.fullmatch(r"[A-Z]{2}\d{4}", na=False)]
        raw_order["구매발주수량"] = parse_mixed_numeric(raw_order["구매발주수량"])
        raw_order["입고예상일자_dt"] = parse_mixed_excel_date(raw_order["입고예상일자_raw"])
        raw_order = raw_order[raw_order["구매발주수량"] > 0]
        if not raw_order.empty:
            order_summary = raw_order.groupby("리드지코드", as_index=False).agg(
                {"리드지명": first_nonempty_text, "구매발주수량": "sum", "입고예상일자_dt": "min"}
            )
            summaries.append(order_summary)

    if purchase_request_sheet is not None:
        try:
            # 구매의뢰현황 기준: G열(품목코드), U열(발주잔량), Y열(요청일).
            raw_request = pd.read_excel(
                order_path, sheet_name=purchase_request_sheet, header=0, usecols=[6, 7, 20, 24]
            )
        except Exception:
            raw_request = pd.DataFrame()
    else:
        raw_request = pd.DataFrame()

    if not raw_request.empty:
        raw_request.columns = ["리드지코드_raw", "리드지명", "구매의뢰수량", "입고예상일자_raw"]
        raw_request["리드지코드"] = raw_request["리드지코드_raw"].map(normalize_leadji_code_key)
        raw_request = raw_request[raw_request["리드지코드"].str.fullmatch(r"[A-Z]{2}\d{4}", na=False)]
        raw_request["구매의뢰수량"] = parse_mixed_numeric(raw_request["구매의뢰수량"])
        raw_request["입고예상일자_dt"] = parse_mixed_excel_date(raw_request["입고예상일자_raw"])
        raw_request = raw_request[raw_request["구매의뢰수량"] > 0]
        if not raw_request.empty:
            request_summary = raw_request.groupby("리드지코드", as_index=False).agg(
                {"리드지명": first_nonempty_text, "구매의뢰수량": "sum", "입고예상일자_dt": "min"}
            )
            summaries.append(request_summary)

    if not summaries:
        return empty

    summary = pd.concat(summaries, ignore_index=True, sort=False)
    for qty_col in ["구매발주수량", "구매의뢰수량"]:
        if qty_col not in summary.columns:
            summary[qty_col] = 0.0
        summary[qty_col] = parse_mixed_numeric(summary[qty_col])
    summary["입고예상일자_dt"] = pd.to_datetime(summary["입고예상일자_dt"], errors="coerce")
    summary = summary.groupby("리드지코드", as_index=False).agg(
        {
            "리드지명": first_nonempty_text,
            "구매발주수량": "sum",
            "구매의뢰수량": "sum",
            "입고예상일자_dt": "min",
        }
    )
    summary["발주수량"] = summary["구매발주수량"] + summary["구매의뢰수량"]
    summary["입고예상일자"] = summary["입고예상일자_dt"].dt.strftime("%Y-%m-%d").fillna("미확인")
    return summary[
        ["리드지코드", "리드지명", "발주수량", "입고예상일자", "입고예상일자_dt", "구매발주수량", "구매의뢰수량"]
    ]


def render_shortage_dashboard(df: pd.DataFrame, updated_at: str) -> None:
    enriched_df = add_rq_group_columns(df)
    filtered = apply_filters(enriched_df, updated_at)
    download_stamp = datetime.now(DISPLAY_TZ).strftime("%Y%m%d_%H%M%S")

    detail_columns = [
        "거래처",
        "이니셜",
        "품목코드",
        "R코드",
        "Q코드",
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

    shortage_views = ["생산 현황", "사출 현황", "분리 현황", "공용 품목 현황"]
    selected_shortage_view = st.segmented_control(
        "공정별 현황",
        options=shortage_views,
        default=shortage_views[0],
        key="shortage_view_selector_v3",
        width="stretch",
    )
    direct_search_cols = [
        c for c in ["거래처", "이니셜", "품목코드", "R코드", "Q코드", "제품명", "R코드 제품명"] if c in filtered.columns
    ]
    search_col, result_col = st.columns([3.2, 1.0])
    with search_col:
        direct_query = st.text_input(
            "직접 검색",
            value="",
            key="shortage_direct_query_v1",
            placeholder="거래처, 이니셜, 품목코드, R코드, Q코드, 제품명으로 검색하세요",
            help="콤마(,)로 여러 키워드를 입력하면 OR 조건으로 검색합니다.",
        ).strip()
    base_filtered_count = len(filtered)
    if direct_query and direct_search_cols:
        filtered = filter_with_terms_any(filtered, direct_search_cols, direct_query)
    with result_col:
        st.caption(f"표시 {len(filtered):,}건 / 전체 {base_filtered_count:,}건")

    if selected_shortage_view == "생산 현황":
        full_demand_summary = build_summary_group_totals_with_safe_split(filtered)
        with st.expander("전체 수요 요약 (분류별요약 × 안전 포함 여부)", expanded=False):
            st.caption("오더 부족수량 = 안전 미포함, 안전재고 부족수량 = 안전 포함 - 안전 미포함, 총수량 = 오더 부족수량 + 안전재고 부족수량")
            if full_demand_summary.empty:
                st.info("전체 수요 요약을 계산할 데이터가 없습니다.")
            else:
                total_row = full_demand_summary.iloc[0]
                s1, s2, s3 = st.columns(3)
                s1.metric("전체 수요 총수량", f"{float(total_row['총수량']):,.0f}")
                s2.metric("오더 부족수량", f"{float(total_row['오더 부족수량']):,.0f}")
                s3.metric("안전재고 부족수량", f"{float(total_row['안전재고 부족수량']):,.0f}")
                full_demand_summary_display = format_numeric_columns_for_display(full_demand_summary)
                full_demand_summary_column_config = build_auto_column_config(
                    full_demand_summary_display,
                    full_demand_summary_display.columns.tolist(),
                    source_df=full_demand_summary,
                )
                st.dataframe(
                    full_demand_summary_display,
                    use_container_width=True,
                    height=320,
                    column_config=full_demand_summary_column_config,
                    hide_index=True,
                )

        inj_shortage_total = (
            parse_mixed_numeric(filtered["사출생산필요수량"]).sum()
            if "사출생산필요수량" in filtered.columns
            else 0
        )
        c1, c2, c3, c4, c5, c6 = st.columns(6, gap="medium")
        with c1:
            render_dashboard_kpi("부족수량 합계", f"{filtered['부족수량'].sum():,.0f}", "risk")
        with c2:
            render_dashboard_kpi("사출부족수량 합계", f"{inj_shortage_total:,.0f}", "risk")
        with c3:
            render_dashboard_kpi("사출 재고", f"{filtered['사출창고'].sum():,.0f}", "stock")
        with c4:
            render_dashboard_kpi("분리 재고", f"{filtered['분리창고'].sum():,.0f}", "stock")
        with c5:
            render_dashboard_kpi("검사접착 재고", f"{filtered['검사접착창고'].sum():,.0f}", "stock")
        with c6:
            render_dashboard_kpi("누수규격 재고", f"{filtered['누수규격검사 창고'].sum():,.0f}", "stock")

        initial_inj_summary = build_initial_injection_summary(filtered)
        with st.expander("이니셜별 사출부족수량 요약", expanded=False):
            st.caption("사출 부족수량 = 이니셜별(품목코드 단위) 사출 생산 필요수량 합계 - 사출창고 합계 (0 미만은 0)")
            if initial_inj_summary.empty:
                st.info("이니셜별 사출부족수량 요약을 계산할 데이터가 없습니다.")
            else:
                initial_inj_summary_display = format_numeric_columns_for_display(initial_inj_summary)
                initial_inj_summary_column_config = build_auto_column_config(
                    initial_inj_summary_display,
                    initial_inj_summary_display.columns.tolist(),
                    source_df=initial_inj_summary,
                )
                st.dataframe(
                    initial_inj_summary_display,
                    use_container_width=True,
                    height=320,
                    column_config=initial_inj_summary_column_config,
                    hide_index=True,
                )

        p_view = filtered.copy()
        p_view["부족수량"] = parse_mixed_numeric(p_view["부족수량"])
        if "사출생산필요수량" in p_view.columns:
            p_view["사출생산필요수량"] = parse_mixed_numeric(p_view["사출생산필요수량"])
        else:
            p_view["사출생산필요수량"] = 0

        mapped_inj_total = 0.0
        unmatched_inj_total = 0.0
        if "품목코드" in p_view.columns:
            item_prefix = p_view["품목코드"].astype(str).str.upper().str[:1]
            p_rows = p_view[item_prefix == "P"].copy()
            r_rows = p_view[item_prefix == "R"].copy()
        else:
            p_rows = p_view.copy()
            r_rows = p_view.iloc[0:0].copy()

        if p_rows.empty:
            p_view["사출 부족수량"] = p_view["사출생산필요수량"]
            key_cols = [c for c in ["사이트코드", "이니셜", "R코드"] if c in p_view.columns]
            if key_cols and not r_rows.empty and "품목코드" in enriched_df.columns:
                # Fallback: when current filters leave only R rows, recover representative P codes
                # from the full scope using (사이트코드+이니셜+R코드) keys.
                p_universe = enriched_df.copy()
                universe_prefix = p_universe["품목코드"].astype(str).str.upper().str[:1]
                p_universe = p_universe[universe_prefix == "P"]
                if not p_universe.empty and all(c in p_universe.columns for c in key_cols):
                    if "부족수량" in p_universe.columns:
                        p_universe["부족수량_num"] = parse_mixed_numeric(p_universe["부족수량"])
                    else:
                        p_universe["부족수량_num"] = 0
                    if "제품명" not in p_universe.columns:
                        p_universe["제품명"] = "-"

                    p_key_map = (
                        p_universe.sort_values(["부족수량_num", "품목코드"], ascending=[False, True])
                        .drop_duplicates(subset=key_cols, keep="first")[key_cols + ["품목코드", "제품명"]]
                        .rename(columns={"품목코드": "매핑P코드", "제품명": "매핑제품명"})
                    )
                    p_view = p_view.merge(p_key_map, on=key_cols, how="left")
                    mapped_mask = p_view["매핑P코드"].astype(str).str.strip().str.lower().ne("nan")
                    mapped_mask = mapped_mask & p_view["매핑P코드"].astype(str).str.strip().ne("")
                    p_view.loc[mapped_mask, "품목코드"] = p_view.loc[mapped_mask, "매핑P코드"]
                    if "제품명" in p_view.columns:
                        p_view.loc[mapped_mask, "제품명"] = p_view.loc[mapped_mask, "매핑제품명"]
                    p_view = p_view.drop(columns=["매핑P코드", "매핑제품명"], errors="ignore")
        else:
            p_rows["사출 부족수량"] = p_rows["사출생산필요수량"]
            key_cols = [c for c in ["사이트코드", "이니셜", "R코드"] if c in p_rows.columns and c in r_rows.columns]

            if key_cols and not r_rows.empty:
                r_key_inj = (
                    r_rows.groupby(key_cols, as_index=False)["사출생산필요수량"]
                    .sum()
                    .rename(columns={"사출생산필요수량": "연결R 사출수량"})
                )
                p_keys = p_rows[key_cols].drop_duplicates()
                matched_r_total = (
                    r_key_inj.merge(p_keys, on=key_cols, how="inner")["연결R 사출수량"].sum()
                    if not p_keys.empty
                    else 0.0
                )
                unmatched_inj_total = float(r_key_inj["연결R 사출수량"].sum() - matched_r_total)

                p_rows = p_rows.merge(r_key_inj, on=key_cols, how="left")
                p_key_short_sum = p_rows.groupby(key_cols)["부족수량"].transform("sum")
                p_key_count = p_rows.groupby(key_cols)["품목코드"].transform("count")

                mapped_by_short = (
                    parse_mixed_numeric(p_rows["연결R 사출수량"])
                    * p_rows["부족수량"]
                    / p_key_short_sum.replace(0, pd.NA)
                )
                mapped_by_split = (
                    parse_mixed_numeric(p_rows["연결R 사출수량"])
                    / p_key_count.replace(0, pd.NA)
                )
                p_rows["사출 부족수량(연결R)"] = mapped_by_short.where(p_key_short_sum > 0, mapped_by_split).fillna(0)
                p_rows["사출 부족수량"] = p_rows["사출 부족수량"].where(
                    p_rows["사출 부족수량"] > 0, p_rows["사출 부족수량(연결R)"]
                )
                mapped_inj_total = float(p_rows["사출 부족수량(연결R)"].sum())
            else:
                unmatched_inj_total = float(parse_mixed_numeric(r_rows["사출생산필요수량"]).sum())

            p_view = p_rows.copy()

        p_view = p_view[(p_view["부족수량"] > 0) | (p_view["사출 부족수량"] > 0)]
        p_view["표시부족수량"] = p_view["부족수량"] + p_view["사출 부족수량"]
        if "납기일" not in p_view.columns:
            p_view["납기일"] = "-"
        if "사출납기일" in p_view.columns:
            due_text = p_view["납기일"].astype(str).str.strip()
            inj_due_text = p_view["사출납기일"].astype(str).str.strip()
            due_missing = due_text.str.lower().isin({"", "-", "nan", "nat", "none"})
            inj_due_valid = ~inj_due_text.str.lower().isin({"", "-", "nan", "nat", "none"})
            p_view.loc[due_missing & inj_due_valid, "납기일"] = inj_due_text[due_missing & inj_due_valid]

        p_detail_columns = detail_columns.copy()
        if "사출 부족수량" not in p_detail_columns:
            insert_idx = p_detail_columns.index("부족수량") + 1 if "부족수량" in p_detail_columns else len(p_detail_columns)
            p_detail_columns.insert(insert_idx, "사출 부족수량")
        p_table = p_view.sort_values(
            ["표시부족수량", "부족수량", "사출 부족수량", "이니셜", "거래처"],
            ascending=[False, False, False, True, True],
        )[p_detail_columns]
        p_table_ui = p_table.drop(columns=["상태"], errors="ignore")
        p_display_columns = p_table_ui.columns.tolist()
        p_table_display = format_numeric_columns_for_display(p_table_ui)
        p_detail_column_config = build_auto_column_config(p_table_display, p_display_columns, source_df=p_table_ui)
        if "사출 부족수량(연결R)" in p_view.columns:
            st.caption(
                f"R→Q→P 연결 매핑(사이트코드+이니셜+R코드): "
                f"P행 반영 사출부족수량 {mapped_inj_total:,.0f}, "
                f"미매핑 R 사출수량 {unmatched_inj_total:,.0f}"
            )
        st.download_button(
            "엑셀 다운로드",
            data=dataframe_to_excel_bytes(p_table, sheet_name="생산현황"),
            file_name=f"shortage_production_{download_stamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_shortage_tab_p",
            use_container_width=False,
        )

        st.dataframe(
            style_operational_table(p_table_display, p_table_ui),
            use_container_width=True,
            height=700,
            column_order=p_display_columns,
            column_config=p_detail_column_config,
            hide_index=True,
            key="shortage_p_table_v2",
        )

    elif selected_shortage_view == "사출 현황":
        r_summary = build_rcode_summary(filtered)

        r1, r2, r3, r4 = st.columns(4)
        r1.metric("R코드 수", f"{len(r_summary):,}")
        r2.metric(
            "R기준 사출 생산 필요수량 합계",
            f"{r_summary['사출 생산 필요수량 합계'].sum():,.0f}" if not r_summary.empty else "0",
        )
        r3.metric("R기준 사출 재고", f"{r_summary['사출창고 합계'].sum():,.0f}" if not r_summary.empty else "0")
        r4.metric("R기준 분리 재고", f"{r_summary['분리창고 합계'].sum():,.0f}" if not r_summary.empty else "0")
        r_summary_ui = r_summary.drop(columns=["상태"], errors="ignore")
        r_summary_display = format_numeric_columns_for_display(r_summary_ui)
        r_summary_column_config = build_auto_column_config(
            r_summary_display, r_summary_display.columns.tolist(), source_df=r_summary_ui
        )
        st.download_button(
            "엑셀 다운로드",
            data=dataframe_to_excel_bytes(r_summary, sheet_name="사출생산현황"),
            file_name=f"shortage_injection_summary_{download_stamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_shortage_tab_r",
            use_container_width=False,
        )

        st.dataframe(
            style_operational_table(r_summary_display, r_summary_ui),
            use_container_width=True,
            height=700,
            column_config=r_summary_column_config,
            hide_index=True,
            key="shortage_r_table_v2",
        )

    elif selected_shortage_view == "분리 현황":
        q_summary = build_qcode_summary(filtered)
        q1, q2, q3 = st.columns(3)
        q1.metric("Q코드 수", f"{len(q_summary):,}")
        q2.metric("Q기준 부족수량 합계", f"{q_summary['부족수량 합계'].sum():,.0f}")
        q3.metric("Q기준 공정재고 합계", f"{q_summary['공정재고 합계'].sum():,.0f}")

        q_sort_cols = ["Q코드5", "Q코드", "부족수량"] if {"Q코드5", "Q코드", "부족수량"}.issubset(filtered.columns) else ["Q코드", "부족수량"]
        q_sort_asc = [True, True, False] if len(q_sort_cols) == 3 else [True, False]
        q_table = filtered.sort_values(q_sort_cols, ascending=q_sort_asc)[detail_columns]
        q_table_ui = q_table.drop(columns=["상태"], errors="ignore")
        q_display_columns = q_table_ui.columns.tolist()
        q_table_display = format_numeric_columns_for_display(q_table_ui)
        q_detail_column_config = build_auto_column_config(q_table_display, q_display_columns, source_df=q_table_ui)
        st.download_button(
            "엑셀 다운로드",
            data=dataframe_to_excel_bytes(q_table, sheet_name="분리생산현황"),
            file_name=f"shortage_separation_{download_stamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_shortage_tab_q",
            use_container_width=False,
        )
        st.dataframe(
            style_operational_table(q_table_display, q_table_ui),
            use_container_width=True,
            height=700,
            column_order=q_display_columns,
            column_config=q_detail_column_config,
            hide_index=True,
            key="shortage_q_table_v2",
        )

    else:
        st.caption("사출 부족수량 = 수요정보 사출 생산 수량 합계 - 사출창고 합계 (0 미만은 0)")
        st.caption("표시 기준: 품목코드는 P코드만, 납기일/부족수량은 누수규격검사 기준")
        rq_filtered = filtered.copy()
        if "품목코드" in rq_filtered.columns:
            rq_filtered = rq_filtered[rq_filtered["품목코드"].astype(str).str.upper().str.startswith("P")]
        if "부족수량" in rq_filtered.columns:
            rq_filtered["부족수량"] = parse_mixed_numeric(rq_filtered["부족수량"])
        else:
            rq_filtered["부족수량"] = 0
        if "사출생산필요수량" in rq_filtered.columns:
            rq_filtered["사출생산필요수량"] = parse_mixed_numeric(rq_filtered["사출생산필요수량"])
        else:
            rq_filtered["사출생산필요수량"] = 0
        rq_filtered = rq_filtered[(rq_filtered["부족수량"] > 0) | (rq_filtered["사출생산필요수량"] > 0)]

        multi_p_r_codes: set[str] = set()
        if {"R코드5", "P코드5", "부족수량"}.issubset(rq_filtered.columns):
            rq_mapping_scope = rq_filtered.copy()
            rq_mapping_scope["R코드5"] = rq_mapping_scope["R코드5"].astype(str).str.strip()
            rq_mapping_scope["P코드5"] = rq_mapping_scope["P코드5"].astype(str).str.strip()
            rq_mapping_scope = rq_mapping_scope[
                rq_mapping_scope["R코드5"].str.startswith("R") & rq_mapping_scope["P코드5"].str.startswith("P")
            ]

            r_to_p_count = rq_mapping_scope.groupby("R코드5")["P코드5"].nunique()
            multi_p_r_codes = set(r_to_p_count[r_to_p_count >= 2].index.tolist())
            if multi_p_r_codes:
                rq_filtered = rq_filtered[rq_filtered["R코드5"].isin(multi_p_r_codes)]
            else:
                rq_filtered = rq_filtered.iloc[0:0]

        if "R코드 제품명" in rq_filtered.columns:
            rq_product_scope = rq_filtered.copy()
            rq_product_scope["부족수량"] = parse_mixed_numeric(rq_product_scope["부족수량"])
            rq_product_scope["사출생산필요수량"] = parse_mixed_numeric(rq_product_scope["사출생산필요수량"])
            rq_product_scope = rq_product_scope[
                (rq_product_scope["부족수량"] > 0) | (rq_product_scope["사출생산필요수량"] > 0)
            ]
            rq_product_scope["표시부족수량"] = rq_product_scope["부족수량"] + rq_product_scope["사출생산필요수량"]
            rq_product_sum_map = (
                rq_product_scope.groupby("R코드 제품명", as_index=True)["표시부족수량"].sum().sort_values(ascending=False).to_dict()
                if not rq_product_scope.empty
                else {}
            )
            rq_product_options = ["전체"] + [
                p for p in list(rq_product_sum_map.keys()) if str(p).strip() not in {"", "-", "nan", "None"}
            ]
            rq_product_count_map = {
                "전체": float(rq_product_scope["표시부족수량"].sum()) if not rq_product_scope.empty else 0.0,
                **rq_product_sum_map,
            }
            rq_selected_product = st.pills(
                "사출 제품명 (R코드5 1개당 P코드5 2+)",
                options=rq_product_options,
                default="전체",
                key="rq_tab_r_product_pills",
                format_func=lambda x: format_pill_label(x, rq_product_count_map),
            )
            if rq_selected_product != "전체":
                rq_filtered = rq_filtered[rq_filtered["R코드 제품명"] == rq_selected_product]

        rq_summary_tab = build_rq_group_summary(rq_filtered)
        rq_shortage_total = (
            parse_mixed_numeric(rq_filtered["부족수량"]).sum()
            if "부족수량" in rq_filtered.columns
            else 0
        )
        rq_inj_shortage_total = (
            parse_mixed_numeric(rq_summary_tab["사출 부족수량"]).sum()
            if not rq_summary_tab.empty and "사출 부족수량" in rq_summary_tab.columns
            else 0
        )
        r1, r2 = st.columns(2)
        r1.metric("사출 부족수량 합계", f"{rq_inj_shortage_total:,.0f}")
        r2.metric("부족수량 합계", f"{rq_shortage_total:,.0f}")

        if rq_summary_tab.empty:
            st.info("표시할 RQ 그룹 데이터가 없습니다.")
            st.download_button(
                "엑셀 다운로드",
                data=dataframe_to_excel_bytes(pd.DataFrame(columns=detail_columns), sheet_name="사출분리공용"),
                file_name=f"shortage_shared_rq_{download_stamp}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_shortage_tab_rq_empty",
                use_container_width=False,
            )
        else:
            rq_sort_cols = ["R코드5", "Q코드5", "부족수량"] if {"R코드5", "Q코드5", "부족수량"}.issubset(rq_filtered.columns) else ["R코드", "Q코드", "부족수량"]
            rq_sort_asc = [True, True, False]
            rq_view = rq_filtered.copy()
            if "납기일" not in rq_view.columns:
                rq_view["납기일"] = "-"
            if "부족수량" not in rq_view.columns:
                rq_view["부족수량"] = 0
            if "사출생산필요수량" in rq_view.columns:
                rq_view["사출부족수량"] = parse_mixed_numeric(rq_view["사출생산필요수량"])
            else:
                rq_view["사출부족수량"] = 0

            rq_detail_columns = [c for c in detail_columns if c in rq_view.columns]
            if "납기일" not in rq_detail_columns:
                insert_idx = rq_detail_columns.index("파워") + 1 if "파워" in rq_detail_columns else len(rq_detail_columns)
                rq_detail_columns.insert(insert_idx, "납기일")
            if "부족수량" not in rq_detail_columns:
                insert_idx = rq_detail_columns.index("납기일") + 1 if "납기일" in rq_detail_columns else len(rq_detail_columns)
                rq_detail_columns.insert(insert_idx, "부족수량")
            if "사출부족수량" not in rq_detail_columns:
                insert_idx = rq_detail_columns.index("부족수량") + 1 if "부족수량" in rq_detail_columns else len(rq_detail_columns)
                rq_detail_columns.insert(insert_idx, "사출부족수량")

            rq_table = rq_view.sort_values(rq_sort_cols, ascending=rq_sort_asc)[rq_detail_columns]
            rq_table_ui = rq_table.drop(columns=["상태"], errors="ignore")
            rq_display_columns = rq_table_ui.columns.tolist()
            rq_table_display = format_numeric_columns_for_display(rq_table_ui)
            rq_detail_column_config = build_auto_column_config(
                rq_table_display, rq_display_columns, source_df=rq_table_ui
            )
            st.download_button(
                "엑셀 다운로드",
                data=dataframe_to_excel_bytes(rq_table, sheet_name="사출분리공용"),
                file_name=f"shortage_shared_rq_{download_stamp}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_shortage_tab_rq",
                use_container_width=False,
            )
            st.dataframe(
                style_operational_table(rq_table_display, rq_table_ui),
                use_container_width=True,
                height=700,
                column_order=rq_display_columns,
                column_config=rq_detail_column_config,
                hide_index=True,
                key="shortage_rq_table_v2",
            )


def build_leadji_requirement_summary(
    shortage_df: pd.DataFrame, leadji_info: pd.DataFrame, leadji_stock: pd.DataFrame
) -> pd.DataFrame:
    fixed_columns = [
        "리드지코드",
        "리드지명",
        "생산필요수량",
        "리드지필요수량",
        "리드지부족",
        "리드지부족수량",
        "최소납기일",
    ]
    if shortage_df.empty or "품목코드" not in shortage_df.columns:
        return pd.DataFrame(columns=fixed_columns)

    p_shortage = build_leadji_p_shortage(shortage_df)
    if p_shortage.empty:
        return pd.DataFrame(columns=fixed_columns)

    if leadji_info.empty:
        return pd.DataFrame(columns=fixed_columns)

    info_cols = leadji_info.columns.tolist()
    prod_col = pick_first_existing_column(info_cols, ["생산"])
    b1_col = pick_first_existing_column(info_cols, ["B1코드"])
    b1_name_col = pick_first_existing_column(info_cols, ["B1코드명"])

    if prod_col is None and len(info_cols) > 3:
        prod_col = info_cols[3]
    if b1_col is None and len(info_cols) > 12:
        b1_col = info_cols[12]
    if b1_name_col is None and len(info_cols) > 13:
        b1_name_col = info_cols[13]

    if prod_col is None or b1_col is None:
        return pd.DataFrame(columns=fixed_columns)

    selected_cols = [prod_col, b1_col] + ([b1_name_col] if b1_name_col is not None else [])
    mapping = leadji_info[selected_cols].copy()
    for col in selected_cols:
        mapping[col] = mapping[col].astype(str).str.strip().replace({"nan": "", "None": ""})

    mapping["P코드5"] = mapping[prod_col].str.upper().str[:5]
    mapping = mapping[(mapping["P코드5"].str.startswith("P")) & (mapping[b1_col] != "")]
    mapping = mapping.rename(columns={b1_col: "리드지코드"})
    if b1_name_col is not None:
        mapping = mapping.rename(columns={b1_name_col: "리드지명"})
    else:
        mapping["리드지명"] = "-"
    mapping = mapping[["P코드5", "리드지코드", "리드지명"]].drop_duplicates(subset=["P코드5", "리드지코드"], keep="first")

    bs_base = p_shortage.merge(mapping, on="P코드5", how="left")
    unmatched_map = bs_base["리드지코드"].isna() | (bs_base["리드지코드"].astype(str).str.strip() == "")
    bs_base.loc[unmatched_map, "리드지코드"] = "매칭필요:" + bs_base.loc[unmatched_map, "P코드5"].astype(str)
    bs_base.loc[unmatched_map, "리드지명"] = "리드지정보 B1코드 없음"
    bs_base["리드지코드"] = bs_base["리드지코드"].fillna("-")
    bs_base["리드지명"] = bs_base["리드지명"].fillna("-")

    summary = (
        bs_base.groupby(["리드지코드", "리드지명"], as_index=False)
        .agg({"생산필요수량": "sum", "최소납기일": "min"})
        .sort_values(["생산필요수량", "리드지코드"], ascending=[False, True])
    )
    warehouse_columns: list[str] = []
    if not leadji_stock.empty:
        stock_cols = leadji_stock.columns.tolist()
        code_col = pick_first_existing_column(stock_cols, ["품목코드"])
        warehouse_col = pick_first_existing_column(stock_cols, ["창고"])
        qty_col = pick_first_existing_column(stock_cols, ["재고"])
        if code_col and warehouse_col and qty_col:
            stock = leadji_stock[[code_col, warehouse_col, qty_col]].copy()
            stock[code_col] = stock[code_col].astype(str).str.strip()
            stock[warehouse_col] = stock[warehouse_col].astype(str).str.strip()
            stock[qty_col] = parse_mixed_numeric(stock[qty_col])
            stock = stock[(stock[code_col] != "") & (stock[warehouse_col] != "") & (stock[qty_col] > 0)]
            if not stock.empty:
                stock = stock.groupby([code_col, warehouse_col], as_index=False)[qty_col].sum()
                pivot = stock.pivot_table(
                    index=code_col,
                    columns=warehouse_col,
                    values=qty_col,
                    aggfunc="sum",
                    fill_value=0,
                )
                if not pivot.empty:
                    warehouse_totals = pivot.sum(axis=0).sort_values(ascending=False)
                    excluded_warehouse_columns = {"L관창고(자재불량)"}
                    warehouse_columns = [
                        str(c) for c in warehouse_totals.index.tolist() if str(c) not in excluded_warehouse_columns
                    ]
                    pivot = pivot.reindex(columns=warehouse_columns).reset_index().rename(columns={code_col: "리드지코드"})
                    summary = summary.merge(pivot, on="리드지코드", how="left")
                    for w_col in warehouse_columns:
                        summary[w_col] = parse_mixed_numeric(summary[w_col])

    active_warehouse_columns: list[str] = []
    for w_col in warehouse_columns:
        col_sum = parse_mixed_numeric(summary[w_col]).sum() if w_col in summary.columns else 0
        if col_sum > 0:
            active_warehouse_columns.append(w_col)

    summary["리드지필요수량"] = (parse_mixed_numeric(summary["생산필요수량"]) * 1.3).round(0)
    leadji_target_warehouses = ["L관창고(자재)", "C관 공정부자재", "S관 공정부자재", "A관 공정부자재"]
    leadji_stock_total = pd.Series(0.0, index=summary.index)
    for warehouse_name in leadji_target_warehouses:
        matched_col = find_warehouse_column(summary.columns.tolist(), [warehouse_name])
        if matched_col is None:
            continue
        leadji_stock_total = leadji_stock_total + parse_mixed_numeric(summary[matched_col])

    shortage_qty = leadji_stock_total - summary["리드지필요수량"]
    summary["리드지부족"] = ""
    summary.loc[shortage_qty < 0, "리드지부족"] = "🔴"
    summary["리드지부족수량"] = shortage_qty.where(shortage_qty < 0)
    summary["최소납기일"] = pd.to_datetime(summary["최소납기일"], errors="coerce").dt.strftime("%Y-%m-%d").fillna("-")
    return summary[
        [
            "리드지코드",
            "리드지명",
            "생산필요수량",
            "리드지필요수량",
            "리드지부족",
            "리드지부족수량",
            *active_warehouse_columns,
            "최소납기일",
        ]
    ]


def compute_leadji_source_total(shortage_df: pd.DataFrame) -> float:
    p_shortage = build_leadji_p_shortage(shortage_df)
    if p_shortage.empty:
        return 0.0
    return float(parse_mixed_numeric(p_shortage["생산필요수량"]).sum())


def build_leadji_p_shortage(shortage_df: pd.DataFrame) -> pd.DataFrame:
    if shortage_df.empty or "품목코드" not in shortage_df.columns:
        return pd.DataFrame(columns=["P코드5", "생산필요수량", "최소납기일"])

    qty_source_col = LEADJI_REQUIRED_QTY_COL if LEADJI_REQUIRED_QTY_COL in shortage_df.columns else "부족수량"
    if qty_source_col not in shortage_df.columns:
        return pd.DataFrame(columns=["P코드5", "생산필요수량", "최소납기일"])

    if qty_source_col == LEADJI_REQUIRED_QTY_COL and LEADJI_REQUIRED_DUE_COL in shortage_df.columns:
        due_source_col = LEADJI_REQUIRED_DUE_COL
    else:
        due_source_col = "최소납기일" if "최소납기일" in shortage_df.columns else "납기일"

    base = shortage_df.copy()
    base["품목코드"] = base["품목코드"].astype(str).str.strip().str.upper()
    base["P코드5"] = base["품목코드"].str[:5]
    base = base[base["P코드5"].str.startswith("P")]
    base["생산필요수량"] = parse_mixed_numeric(base[qty_source_col])
    base["완료재고수량"] = (
        parse_mixed_numeric(base[LEADJI_COMPLETED_STOCK_COL])
        if LEADJI_COMPLETED_STOCK_COL in base.columns
        else 0
    )
    if due_source_col in base.columns:
        base["납기일_dt"] = pd.to_datetime(base[due_source_col], errors="coerce")
    else:
        base["납기일_dt"] = pd.NaT

    item_shortage = (
        base.groupby("품목코드", as_index=False)
        .agg({"P코드5": "first", "생산필요수량": "sum", "완료재고수량": "max", "납기일_dt": "min"})
    )
    # 품목코드 단위로 필요수량 합산 후 완료재고(누수규격검사 창고)를 1회 차감한다.
    item_shortage["생산필요수량"] = (item_shortage["생산필요수량"] - item_shortage["완료재고수량"]).clip(lower=0)
    item_shortage = item_shortage[item_shortage["생산필요수량"] > 0]

    p_shortage = (
        item_shortage.groupby("P코드5", as_index=False)
        .agg({"생산필요수량": "sum", "납기일_dt": "min"})
        .rename(columns={"납기일_dt": "최소납기일"})
    )
    return p_shortage


def build_pcode5_leadji_requirement_summary(
    shortage_df: pd.DataFrame, leadji_info: pd.DataFrame, leadji_stock: pd.DataFrame
) -> pd.DataFrame:
    fixed_columns = ["생산코드", "리드지코드", "리드지명", "생산필요수량", "최소납기일"]
    if shortage_df.empty or "품목코드" not in shortage_df.columns:
        return pd.DataFrame(columns=fixed_columns)

    p_shortage = build_leadji_p_shortage(shortage_df)
    if p_shortage.empty:
        return pd.DataFrame(columns=fixed_columns)

    info_cols = leadji_info.columns.tolist() if not leadji_info.empty else []
    prod_col = pick_first_existing_column(info_cols, ["생산"])
    b1_col = pick_first_existing_column(info_cols, ["B1코드"])
    b1_name_col = pick_first_existing_column(info_cols, ["B1코드명"])

    if prod_col is None and len(info_cols) > 3:
        prod_col = info_cols[3]
    if b1_col is None and len(info_cols) > 12:
        b1_col = info_cols[12]
    if b1_name_col is None and len(info_cols) > 13:
        b1_name_col = info_cols[13]

    if prod_col is not None and b1_col is not None:
        selected_cols = [prod_col, b1_col] + ([b1_name_col] if b1_name_col is not None else [])
        mapping = leadji_info[selected_cols].copy()
        for col in selected_cols:
            mapping[col] = mapping[col].astype(str).str.strip().replace({"nan": "", "None": ""})
        mapping["P코드5"] = mapping[prod_col].str.upper().str[:5]
        mapping = mapping[(mapping["P코드5"].str.startswith("P")) & (mapping[b1_col] != "")]
        mapping = mapping.rename(columns={b1_col: "리드지코드"})
        if b1_name_col is not None:
            mapping = mapping.rename(columns={b1_name_col: "리드지명"})
        else:
            mapping["리드지명"] = "-"
        mapping = mapping[["P코드5", "리드지코드", "리드지명"]].drop_duplicates(subset=["P코드5", "리드지코드"], keep="first")
        detail = p_shortage.merge(mapping, on="P코드5", how="left")
    else:
        detail = p_shortage.copy()
        detail["리드지코드"] = "-"
        detail["리드지명"] = "-"

    detail["리드지코드"] = detail["리드지코드"].fillna("-")
    detail["리드지명"] = detail["리드지명"].fillna("-")

    summary = (
        detail.groupby(["P코드5", "리드지코드", "리드지명"], as_index=False)
        .agg({"생산필요수량": "sum", "최소납기일": "min"})
        .sort_values(["생산필요수량", "P코드5", "리드지코드"], ascending=[False, True, True])
    )

    warehouse_columns: list[str] = []
    if not leadji_stock.empty:
        stock_cols = leadji_stock.columns.tolist()
        code_col = pick_first_existing_column(stock_cols, ["품목코드"])
        warehouse_col = pick_first_existing_column(stock_cols, ["창고"])
        qty_col = pick_first_existing_column(stock_cols, ["재고"])
        if code_col and warehouse_col and qty_col:
            stock = leadji_stock[[code_col, warehouse_col, qty_col]].copy()
            stock[code_col] = stock[code_col].astype(str).str.strip()
            stock[warehouse_col] = stock[warehouse_col].astype(str).str.strip()
            stock[qty_col] = parse_mixed_numeric(stock[qty_col])
            stock = stock[(stock[code_col] != "") & (stock[warehouse_col] != "") & (stock[qty_col] > 0)]
            if not stock.empty:
                stock = stock.groupby([code_col, warehouse_col], as_index=False)[qty_col].sum()
                pivot = stock.pivot_table(
                    index=code_col,
                    columns=warehouse_col,
                    values=qty_col,
                    aggfunc="sum",
                    fill_value=0,
                )
                if not pivot.empty:
                    warehouse_totals = pivot.sum(axis=0).sort_values(ascending=False)
                    excluded_warehouse_columns = {"L관창고(자재불량)"}
                    warehouse_columns = [
                        str(c) for c in warehouse_totals.index.tolist() if str(c) not in excluded_warehouse_columns
                    ]
                    pivot = pivot.reindex(columns=warehouse_columns).reset_index().rename(columns={code_col: "리드지코드"})
                    summary = summary.merge(pivot, on="리드지코드", how="left")
                    for w_col in warehouse_columns:
                        summary[w_col] = parse_mixed_numeric(summary[w_col])

    active_warehouse_columns: list[str] = []
    for w_col in warehouse_columns:
        col_sum = parse_mixed_numeric(summary[w_col]).sum() if w_col in summary.columns else 0
        if col_sum > 0:
            active_warehouse_columns.append(w_col)

    summary["최소납기일"] = pd.to_datetime(summary["최소납기일"], errors="coerce").dt.strftime("%Y-%m-%d").fillna("-")
    summary = summary.rename(columns={"P코드5": "생산코드"})
    return summary[["생산코드", "리드지코드", "리드지명", "생산필요수량", *active_warehouse_columns, "최소납기일"]]


def merge_leadji_with_order_status(summary_df: pd.DataFrame, leadji_order_df: pd.DataFrame) -> pd.DataFrame:
    merged = summary_df.copy()
    merged["리드지코드_join"] = merged["리드지코드"].map(normalize_leadji_code_key)

    def first_nonempty_text(series: pd.Series) -> str:
        text = series.astype(str).str.strip()
        text = text[(text != "") & (text.str.lower() != "nan") & (text.str.lower() != "none")]
        return text.iloc[0] if not text.empty else "-"

    if leadji_order_df.empty:
        merged["발주수량"] = 0.0
        merged["구매발주수량"] = 0.0
        merged["구매의뢰수량"] = 0.0
        merged["입고예상일자_dt"] = pd.NaT
    else:
        order = leadji_order_df.copy()
        order["리드지코드_join"] = order["리드지코드"].map(normalize_leadji_code_key)
        for qty_col in ["구매발주수량", "구매의뢰수량"]:
            if qty_col not in order.columns:
                order[qty_col] = 0.0
            order[qty_col] = parse_mixed_numeric(order[qty_col])
        if "리드지명" not in order.columns:
            order["리드지명"] = "-"
        order = (
            order.groupby("리드지코드_join", as_index=False)
            .agg(
                {
                    "리드지코드": first_nonempty_text,
                    "리드지명": first_nonempty_text,
                    "발주수량": "sum",
                    "구매발주수량": "sum",
                    "구매의뢰수량": "sum",
                    "입고예상일자_dt": "min",
                }
            )
            .rename(
                columns={
                    "리드지코드": "리드지코드_order",
                    "리드지명": "리드지명_order",
                    "발주수량": "발주수량_join",
                }
            )
        )
        summary_keys = set(merged["리드지코드_join"].dropna().astype(str))
        order_only = order[~order["리드지코드_join"].astype(str).isin(summary_keys)].copy()
        merged = merged.merge(order, on="리드지코드_join", how="left")
        merged["발주수량"] = parse_mixed_numeric(merged["발주수량_join"])
        merged["구매발주수량"] = parse_mixed_numeric(merged["구매발주수량"])
        merged["구매의뢰수량"] = parse_mixed_numeric(merged["구매의뢰수량"])

        if not order_only.empty:
            extra = pd.DataFrame(index=order_only.index)
            for col in summary_df.columns:
                if col in {"리드지코드", "리드지명", "리드지부족", "최소납기일"}:
                    extra[col] = "-"
                elif col == "리드지부족수량":
                    extra[col] = pd.NA
                else:
                    extra[col] = 0.0

            if "리드지코드" in extra.columns:
                extra["리드지코드"] = order_only["리드지코드_order"].where(
                    order_only["리드지코드_order"].astype(str).str.strip() != "",
                    order_only["리드지코드_join"],
                )
            if "리드지명" in extra.columns:
                extra["리드지명"] = order_only["리드지명_order"].fillna("-")
            if "리드지부족" in extra.columns:
                extra["리드지부족"] = ""
            if "최소납기일" in extra.columns:
                extra["최소납기일"] = "-"

            extra["리드지코드_join"] = order_only["리드지코드_join"]
            extra["발주수량"] = parse_mixed_numeric(order_only["발주수량_join"])
            extra["구매발주수량"] = parse_mixed_numeric(order_only["구매발주수량"])
            extra["구매의뢰수량"] = parse_mixed_numeric(order_only["구매의뢰수량"])
            extra["입고예상일자_dt"] = pd.to_datetime(order_only["입고예상일자_dt"], errors="coerce")
            merged = pd.concat([merged, extra], ignore_index=True, sort=False)

    merged["입고예상일자_dt"] = pd.to_datetime(merged["입고예상일자_dt"], errors="coerce")
    merged["입고예상일자"] = merged["입고예상일자_dt"].dt.strftime("%Y-%m-%d").fillna("미확인")

    shortage_raw = parse_mixed_numeric(merged["리드지부족수량"])
    shortage_qty = shortage_raw.where(shortage_raw > 0, -shortage_raw).clip(lower=0)
    shortage_mask = shortage_qty > 0
    missing_due_mask = merged["입고예상일자_dt"].isna()
    enough_order_mask = merged["발주수량"] >= shortage_qty
    has_purchase_order_mask = parse_mixed_numeric(merged["구매발주수량"]) > 0
    has_purchase_request_mask = parse_mixed_numeric(merged["구매의뢰수량"]) > 0
    order_only_mask = parse_mixed_numeric(merged["생산필요수량"]) <= 0

    merged["상태"] = "부족 없음"
    merged.loc[
        order_only_mask & ~missing_due_mask & has_purchase_order_mask & ~has_purchase_request_mask,
        "상태",
    ] = "입고 예정"
    merged.loc[
        order_only_mask & ~missing_due_mask & ~has_purchase_order_mask & has_purchase_request_mask,
        "상태",
    ] = "구매의뢰"
    merged.loc[
        order_only_mask & ~missing_due_mask & has_purchase_order_mask & has_purchase_request_mask,
        "상태",
    ] = "입고 예정+의뢰"
    merged.loc[order_only_mask & missing_due_mask & (has_purchase_order_mask | has_purchase_request_mask), "상태"] = (
        "입고일 미확인"
    )
    merged.loc[shortage_mask & missing_due_mask, "상태"] = "입고일 미확인"
    merged.loc[
        shortage_mask & ~missing_due_mask & enough_order_mask & has_purchase_order_mask & ~has_purchase_request_mask,
        "상태",
    ] = "입고 예정"
    merged.loc[
        shortage_mask & ~missing_due_mask & enough_order_mask & ~has_purchase_order_mask & has_purchase_request_mask,
        "상태",
    ] = "구매의뢰"
    merged.loc[
        shortage_mask & ~missing_due_mask & enough_order_mask & has_purchase_order_mask & has_purchase_request_mask,
        "상태",
    ] = "입고 예정+의뢰"
    merged.loc[shortage_mask & ~missing_due_mask & ~enough_order_mask, "상태"] = "발주부족"

    ordered_cols: list[str] = []
    for col in summary_df.columns:
        ordered_cols.append(col)
        if col == "리드지부족수량":
            ordered_cols.extend(["발주수량", "입고예상일자", "상태"])

    keep_cols = [c for c in ordered_cols if c in merged.columns] + ["입고예상일자_dt"]
    return merged[keep_cols]


def render_leadji_dashboard(
    updated_at: str,
    shortage_df: pd.DataFrame,
    leadji_info: pd.DataFrame,
    leadji_stock: pd.DataFrame,
    leadji_order_df: pd.DataFrame,
) -> None:
    st.subheader("리드지 현황")
    st.caption(f"업데이트: {updated_at}")
    download_stamp = datetime.now(DISPLAY_TZ).strftime("%Y%m%d_%H%M%S")

    summary_df = build_leadji_requirement_summary(shortage_df, leadji_info, leadji_stock)
    if summary_df.empty and leadji_order_df.empty:
        st.warning("리드지재고현황을 계산할 데이터가 없습니다.")
    else:
        summary_df = merge_leadji_with_order_status(summary_df, leadji_order_df)
        st.warning("입고예정일자는 구매의뢰 기준입니다. 실제 입고 일정은 구매팀 확인이 필요합니다.")

        stock_target_names = ["L관창고(자재)", "C관 공정부자재", "S관 공정부자재", "A관 공정부자재"]
        stock_detail_columns: list[str] = []
        summary_df["재고합계"] = 0.0
        for warehouse_name in stock_target_names:
            matched_col = find_warehouse_column(summary_df.columns.tolist(), [warehouse_name])
            display_col = warehouse_name
            if matched_col is not None:
                summary_df[display_col] = parse_mixed_numeric(summary_df[matched_col])
            elif display_col not in summary_df.columns:
                summary_df[display_col] = 0.0
            summary_df["재고합계"] = summary_df["재고합계"] + parse_mixed_numeric(summary_df[display_col])
            stock_detail_columns.append(display_col)

        summary_df["생산 최소 납기일"] = summary_df["최소납기일"] if "최소납기일" in summary_df.columns else "-"
        shortage_numeric = parse_mixed_numeric(summary_df["리드지부족수량"])
        shortage_abs = (-shortage_numeric).clip(lower=0)
        inbound_date = pd.to_datetime(summary_df["입고예상일자_dt"], errors="coerce")
        summary_df["부족수량_abs"] = shortage_abs
        summary_df["우선순위"] = "정상"
        summary_df.loc[(shortage_numeric < 0) & inbound_date.isna(), "우선순위"] = "긴급"
        summary_df.loc[(shortage_numeric < 0) & inbound_date.notna(), "우선순위"] = "확인필요"
        priority_order = {"긴급": 0, "확인필요": 1, "정상": 2}
        summary_df["우선순위정렬"] = summary_df["우선순위"].map(priority_order).fillna(9)
        summary_df = summary_df.sort_values(
            ["우선순위정렬", "부족수량_abs", "리드지코드"],
            ascending=[True, False, True],
        )

        total_codes = summary_df["리드지코드"].astype(str).str.strip().replace("", pd.NA).dropna().nunique()
        shortage_count = int((shortage_numeric < 0).sum())
        inbound_planned_count = int(summary_df["상태"].astype(str).str.contains("입고 예정", regex=False).sum())
        total_shortage_qty = float(shortage_abs.sum())
        st.markdown(
            f"""
            <style>
            .leadji-kpi-grid {{
                display: grid;
                grid-template-columns: repeat(4, minmax(0, 1fr));
                gap: 14px;
                margin: 16px 0 20px;
            }}
            .leadji-kpi-card {{
                border: 1px solid #E5E7EB;
                border-left: 4px solid #1A2B5E;
                border-radius: 12px;
                background: #FFFFFF;
                padding: 16px 18px;
                min-height: 104px;
                box-shadow: 0 8px 22px rgba(15, 23, 42, 0.06);
            }}
            .leadji-kpi-card strong {{
                display: block;
                color: #64748B;
                font-size: 13px;
                font-weight: 700;
                margin-bottom: 8px;
            }}
            .leadji-kpi-card span {{
                display: block;
                color: #374151;
                font-size: 30px;
                font-weight: 850;
                line-height: 1.15;
            }}
            .leadji-kpi-card.risk {{
                background: #FFFFFF;
                border-color: #E5E7EB;
                border-left-color: #DC2626;
            }}
            .leadji-kpi-card.risk strong,
            .leadji-kpi-card.risk span {{
                color: #DC2626;
            }}
            .leadji-kpi-card.inbound {{
                background: #FFFFFF;
                border-color: #E5E7EB;
                border-left-color: #1A2B5E;
            }}
            .leadji-kpi-card.inbound span {{
                color: #1A2B5E;
            }}
            @media (max-width: 900px) {{
                .leadji-kpi-grid {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
            }}
            </style>
            <div class="leadji-kpi-grid">
                <div class="leadji-kpi-card"><strong>전체 리드지코드</strong><span>{total_codes:,.0f}</span></div>
                <div class="leadji-kpi-card risk"><strong>부족 리드지</strong><span>{shortage_count:,.0f}</span></div>
                <div class="leadji-kpi-card inbound"><strong>입고예정</strong><span>{inbound_planned_count:,.0f}</span></div>
                <div class="leadji-kpi-card risk"><strong>총 리드지 부족수량</strong><span>{total_shortage_qty:,.0f}</span></div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        qcol, _ = st.columns([1.35, 2.65])
        with qcol:
            summary_query = st.text_input(
                "검색",
                value="",
                key="leadji_summary_query",
                placeholder="리드지코드, 리드지명, 창고, 상태로 검색하세요",
            ).strip()

        hidden_cols = ["입고예상일자_dt", "우선순위정렬", "부족수량_abs", "최소납기일"]
        summary_visible = summary_df.drop(columns=hidden_cols, errors="ignore")
        summary_search_cols = [c for c in summary_visible.columns if c not in ["생산필요수량"]]
        filtered_visible = filter_with_terms_any(summary_visible, summary_search_cols, summary_query)
        filtered_summary = summary_df.loc[filtered_visible.index].copy()

        priority_columns = [
            "리드지코드",
            "리드지명",
            "리드지부족수량",
            "발주수량",
            "입고예상일자",
            "생산 최소 납기일",
        ]
        priority_rows = (
            filtered_summary[filtered_summary["우선순위"].isin(["긴급", "확인필요"])]
            .sort_values("부족수량_abs", ascending=False)
            .head(10)
        )
        st.markdown(
            f"""
            <div class="dashboard-section-header">
                <h3>우선 확인 필요 리스트</h3>
                <span class="dashboard-count-badge">최대 {min(len(priority_rows), 10):,}건</span>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if priority_rows.empty:
            st.info("우선 확인이 필요한 리드지가 없습니다.")
        else:
            priority_table = priority_rows[[c for c in priority_columns if c in priority_rows.columns]]
            priority_display = format_numeric_columns_for_display(priority_table)
            priority_column_config = build_auto_column_config(
                priority_display, priority_display.columns.tolist(), source_df=priority_table
            )
            priority_styled = style_leadji_shortage_table(priority_display, priority_table)
            st.dataframe(
                priority_styled,
                use_container_width=True,
                height=min(430, 78 + len(priority_table) * 38),
                column_config=priority_column_config,
                hide_index=True,
            )

        st.markdown(
            f"""
            <div class="dashboard-section-header">
                <h3>리드지 목록</h3>
                <span class="dashboard-count-badge">전체 {len(filtered_summary):,}건</span>
            </div>
            <div class="dashboard-section-subtle">핵심 운영 컬럼만 기본 표시합니다. 창고별 수량은 아래 재고 상세 컬럼에서 확인하세요.</div>
            """,
            unsafe_allow_html=True,
        )
        basic_columns = [
            "리드지코드",
            "리드지명",
            "상태",
            "리드지필요수량",
            "재고합계",
            "리드지부족수량",
            "발주수량",
            "입고예상일자",
            "생산 최소 납기일",
        ]
        table_df = filtered_summary[[c for c in basic_columns if c in filtered_summary.columns]]

        download_df = filtered_summary.drop(columns=hidden_cols, errors="ignore")
        st.download_button(
            "엑셀 다운로드",
            data=dataframe_to_excel_bytes(download_df, sheet_name="리드지현황"),
            file_name=f"leadji_status_{download_stamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_leadji_summary",
            use_container_width=False,
        )

        leadji_display = format_numeric_columns_for_display(table_df)
        leadji_column_config = build_auto_column_config(
            leadji_display, leadji_display.columns.tolist(), source_df=table_df
        )
        leadji_column_config.pop("리드지부족", None)
        leadji_styled = style_leadji_shortage_table(leadji_display, table_df)
        st.dataframe(
            leadji_styled,
            use_container_width=True,
            height=620,
            column_config=leadji_column_config,
            hide_index=True,
        )

        with st.expander("재고 상세 컬럼"):
            stock_table_columns = [
                "리드지코드",
                "리드지명",
                "재고합계",
                *stock_detail_columns,
            ]
            stock_table = filtered_summary[[c for c in stock_table_columns if c in filtered_summary.columns]]
            stock_display = format_numeric_columns_for_display(stock_table)
            stock_column_config = build_auto_column_config(
                stock_display, stock_display.columns.tolist(), source_df=stock_table
            )
            st.dataframe(
                stock_display,
                use_container_width=True,
                height=420,
                column_config=stock_column_config,
                hide_index=True,
            )

        source_total = compute_leadji_source_total(shortage_df)
        summary_total = float(parse_mixed_numeric(summary_df["생산필요수량"]).sum())
        verify_diff = summary_total - source_total
        with st.expander("데이터 검증 정보"):
            st.caption(
                f"검증: 품목코드별 ({LEADJI_REQUIRED_QTY_COL} - {LEADJI_COMPLETED_STOCK_COL}) 합계 {source_total:,.0f} / "
                f"리드지 합계 {summary_total:,.0f} / 차이 {verify_diff:,.0f}"
            )


def render_leadji_pcode5_dashboard(
    updated_at: str, shortage_df: pd.DataFrame, leadji_info: pd.DataFrame, leadji_stock: pd.DataFrame
) -> None:
    st.subheader("생산코드별 리드지 현황")
    st.caption(f"업데이트: {updated_at}")
    st.caption(
        f"집계 기준: 품목코드별 ({LEADJI_REQUIRED_QTY_COL} - {LEADJI_COMPLETED_STOCK_COL})를 0 미만 0으로 만든 뒤 P코드 단위 합산(sum)"
    )
    st.caption("기준: 동일 생산코드에 여러 리드지가 매핑되면 생산필요수량이 각 리드지 행에 반복 표시됩니다.")
    download_stamp = datetime.now(DISPLAY_TZ).strftime("%Y%m%d_%H%M%S")

    summary_df = build_pcode5_leadji_requirement_summary(shortage_df, leadji_info, leadji_stock)
    if summary_df.empty:
        st.warning("생산코드별 리드지 현황을 계산할 데이터가 없습니다.")
        return

    qcol, _ = st.columns([3.0, 1.0])
    with qcol:
        summary_query = st.text_input(
            "통합 검색 (생산코드/리드지코드/리드지명/창고)",
            value="",
            key="leadji_pcode5_summary_query",
            placeholder="예: P1234, BS0314, 블리스터케이스, 원료창고",
        ).strip()

    summary_search_cols = [c for c in summary_df.columns if c not in ["생산필요수량"]]
    filtered_summary = filter_with_terms_any(summary_df, summary_search_cols, summary_query)

    c1, c2, c3 = st.columns(3)
    c1.metric("생산코드 수", f"{filtered_summary['생산코드'].astype(str).nunique():,}")
    p_qty_total = (
        filtered_summary.drop_duplicates(subset=["생산코드"], keep="first")["생산필요수량"].sum()
        if not filtered_summary.empty
        else 0
    )
    c2.metric("생산코드 기준 생산필요수량 합계", f"{p_qty_total:,.0f}")
    min_due_dt = pd.to_datetime(filtered_summary["최소납기일"], errors="coerce").min()
    c3.metric("생산 최소 납기일", "-" if pd.isna(min_due_dt) else min_due_dt.strftime("%Y-%m-%d"))

    st.download_button(
        "엑셀 다운로드",
        data=dataframe_to_excel_bytes(filtered_summary, sheet_name="생산코드기준리드지현황"),
        file_name=f"leadji_status_by_production_code_{download_stamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_leadji_pcode5_summary",
        use_container_width=False,
    )

    leadji_p_display = format_numeric_columns_for_display(filtered_summary)
    leadji_p_column_config = build_auto_column_config(
        leadji_p_display, leadji_p_display.columns.tolist(), source_df=filtered_summary
    )
    st.dataframe(
        leadji_p_display,
        use_container_width=True,
        height=700,
        column_config=leadji_p_column_config,
        hide_index=True,
    )



def main() -> None:
    inject_dashboard_theme()
    st.markdown(
        """
        <div class="dashboard-hero">
            <div class="dashboard-hero-title">생산현황</div>
            <p class="dashboard-hero-subtitle">생산 부족 리스크, 공정 재고, 자재 입고 현황을<br>실시간으로 모니터링합니다.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    data_base_dir = BASE_DIR
    updated_at = get_data_updated_at(data_base_dir)

    try:
        refresh_key = build_data_refresh_key(data_base_dir)
        reference_refresh_key = build_reference_refresh_key(data_base_dir)
        leadji_order_refresh_key = build_leadji_order_refresh_key(data_base_dir)
        df, _, _ = load_data(refresh_key, str(data_base_dir))
        leadji_info, leadji_stock = load_leadji_data(reference_refresh_key, str(data_base_dir))
        leadji_order_df = load_leadji_order_data(leadji_order_refresh_key, str(data_base_dir))
    except Exception as exc:
        st.error(f"데이터 로드 실패: {exc}")
        st.stop()

    top_views = ["생산 부족 현황", "리드지 현황", "생산코드별 리드지"]
    with st.sidebar:
        st.markdown(
            """
            <div class="sidebar-brand">
                <span class="sidebar-brand-icon">
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" aria-hidden="true">
                        <path d="M4 19V5" stroke="currentColor" stroke-width="2" stroke-linecap="round"/>
                        <path d="M4 19H20" stroke="currentColor" stroke-width="2" stroke-linecap="round"/>
                        <path d="M8 16V11" stroke="currentColor" stroke-width="2" stroke-linecap="round"/>
                        <path d="M12 16V7" stroke="currentColor" stroke-width="2" stroke-linecap="round"/>
                        <path d="M16 16V9" stroke="currentColor" stroke-width="2" stroke-linecap="round"/>
                    </svg>
                </span>
                <span class="sidebar-brand-title">생산현황</span>
            </div>
            <div class="sidebar-section-title">메뉴</div>
            """,
            unsafe_allow_html=True,
        )
        selected_top_view = st.radio(
            "메뉴",
            options=top_views,
            index=0,
            key="top_view_radio_v2",
            label_visibility="collapsed",
        )
        st.markdown('<div class="sidebar-divider"></div>', unsafe_allow_html=True)

    if selected_top_view == "생산 부족 현황":
        render_shortage_dashboard(df, updated_at)
    elif selected_top_view == "리드지 현황":
        render_leadji_dashboard(updated_at, df, leadji_info, leadji_stock, leadji_order_df)
    else:
        render_leadji_pcode5_dashboard(updated_at, df, leadji_info, leadji_stock)


if __name__ == "__main__":
    main()
