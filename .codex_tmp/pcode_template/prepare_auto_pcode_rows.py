# -*- coding: utf-8 -*-
import json
from pathlib import Path

import pandas as pd


BASE_DIR = Path(__file__).resolve().parents[2]
SNAPSHOT = BASE_DIR / "cloud_snapshots" / "all_item_status_snapshot.csv.gz"
OUTPUT = Path(__file__).resolve().parent / "pcode_auto_rows.json"
FINISHED_GOODS_SHEET_HINTS = ("전체 품목코드 재고", "품목코드 변화 조회결과")

SRC = {
    "site": "사이트코드",
    "primary": "제품대분류",
    "customer_group": "거래처그룹",
    "customer": "거래처",
    "initial": "이니셜",
    "product_class": "신규분류",
    "product": "제품명",
    "power": "파워",
    "due": "납기일",
    "injection_code": "사출코드",
    "separation_code": "분리코드",
    "production_code": "생산코드",
    "order_qty": "오더수량",
    "shortage_qty": "생산부족수량",
    "injection_shortage_qty": "사출부족수량",
    "injection_stock": "사출창고",
    "separation_stock": "분리창고",
    "adhesion_stock": "검사접착창고",
    "leakage_stock": "누수규격검사",
    "process_stock": "공정재고합계",
    "finished_stock": "완제품재고",
    "doi_order": "DOI기준오더",
    "doi": "DOI",
    "stock_signal": "신호",
    "status": "상태",
}

TEXT_COLS = [
    "site",
    "primary",
    "customer_group",
    "customer",
    "initial",
    "product_class",
    "product",
    "power",
    "due",
    "injection_code",
    "separation_code",
    "production_code",
    "stock_signal",
    "status",
]
NUM_COLS = [
    "order_qty",
    "shortage_qty",
    "injection_shortage_qty",
    "injection_stock",
    "separation_stock",
    "adhesion_stock",
    "leakage_stock",
    "process_stock",
    "finished_stock",
    "stock_change",
    "doi_order",
    "doi",
]


def clean_text(series: pd.Series) -> pd.Series:
    return (
        series.fillna("")
        .astype(str)
        .replace({"nan": "", "None": "", "NaT": "", "<NA>": ""})
        .str.strip()
    )


def clean_scalar(value: object) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    text = str(value).strip()
    return "" if text.lower() in {"", "nan", "none", "nat", "null", "<na>"} else text


def normalize_code(value: object) -> str:
    return clean_scalar(value).replace(" ", "").upper()


def pick_col(columns: list[str], candidates: list[str]) -> str | None:
    normalized = {str(col).strip().replace(" ", "").lower(): col for col in columns}
    for candidate in candidates:
        key = candidate.strip().replace(" ", "").lower()
        if key in normalized:
            return normalized[key]
    return None


def summarize_unique(values: pd.Series, head_count: int = 2) -> str:
    cleaned = [clean_scalar(value) for value in values]
    ordered = list(dict.fromkeys(value for value in cleaned if value))
    if not ordered:
        return ""
    if len(ordered) <= head_count:
        return ", ".join(ordered)
    return f"{', '.join(ordered[:head_count])} 외 {len(ordered) - head_count}"


def summarize_signal(values: pd.Series) -> str:
    cleaned = [clean_scalar(value) for value in values]
    cleaned = [value for value in cleaned if value]
    if not cleaned:
        return ""
    priority = {"소진": 0, "감소": 1, "신규": 2, "증가": 3, "유지": 4}
    ordered = sorted(dict.fromkeys(cleaned), key=lambda value: (priority.get(value, 9), value))
    if len(ordered) <= 2:
        return ", ".join(ordered)
    return f"{', '.join(ordered[:2])} 외 {len(ordered) - 2}"


def workbook_has_stock_sheet(path: Path) -> bool:
    try:
        xls = pd.ExcelFile(path)
    except Exception:
        return False
    normalized = {name.replace(" ", "") for name in xls.sheet_names}
    return any(hint.replace(" ", "") in normalized for hint in FINISHED_GOODS_SHEET_HINTS)


def find_finished_goods_file() -> Path | None:
    search_dirs = [BASE_DIR, BASE_DIR / "cloud_snapshots", Path.home() / "Downloads"]
    candidates: list[Path] = []
    for search_dir in search_dirs:
        if not search_dir.exists():
            continue
        for path in search_dir.glob("*.xlsx"):
            if path.name.startswith("~$"):
                continue
            normalized_name = path.name.replace(" ", "")
            looks_relevant = (
                ("재고변화" in normalized_name and "품목코드" in normalized_name)
                or path.name == "완제품_재고변화_uploaded.xlsx"
                or ("LOT" in normalized_name.upper() and "재고" in normalized_name and "품목코드" in normalized_name)
            )
            if looks_relevant and workbook_has_stock_sheet(path):
                candidates.append(path)
    if not candidates:
        return None
    return max(candidates, key=lambda p: p.stat().st_mtime)


def read_finished_goods_summary() -> tuple[pd.DataFrame, str]:
    path = find_finished_goods_file()
    output_columns = [
        "production_code",
        "finished_stock",
        "stock_change",
        "doi_order",
        "doi",
        "stock_ratio",
        "stock_signal",
        "stock_action",
    ]
    if path is None:
        return pd.DataFrame(columns=output_columns), ""

    xls = pd.ExcelFile(path)
    sheet_map = {name.replace(" ", ""): name for name in xls.sheet_names}
    sheet_name = next(
        (sheet_map[hint.replace(" ", "")] for hint in FINISHED_GOODS_SHEET_HINTS if hint.replace(" ", "") in sheet_map),
        xls.sheet_names[0],
    )

    parsed = pd.DataFrame()
    for header_row in (1, 0, 2):
        try:
            candidate = pd.read_excel(path, sheet_name=sheet_name, header=header_row)
        except Exception:
            continue
        candidate.columns = [str(col).strip() for col in candidate.columns]
        code_col = pick_col(candidate.columns.tolist(), ["품목코드", "제품코드", "ITEM_ID", "제품 코드"])
        if code_col:
            parsed = candidate
            break

    if parsed.empty:
        return pd.DataFrame(columns=output_columns), str(path)

    columns = parsed.columns.tolist()
    code_col = pick_col(columns, ["품목코드", "제품코드", "ITEM_ID", "제품 코드"])
    end_stock_col = pick_col(columns, ["종료재고", "기말재고", "현재재고", "재고"])
    change_col = pick_col(columns, ["변화", "재고변화", "증감", "증감수량"])
    order_col = pick_col(columns, ["오더", "오더수량", "총오더", "ORDER_QTY"])
    doi_col = pick_col(columns, ["DOI(일)", "DOI", "DOI일"])
    ratio_col = pick_col(columns, ["비율", "재고비율", "증감비율"])
    signal_col = pick_col(columns, ["신호", "SIGNAL", "Signal"])
    action_col = pick_col(columns, ["대응판단", "판단", "조치판단"])

    stock = pd.DataFrame({"production_code": parsed[code_col].map(normalize_code)})
    stock = stock[stock["production_code"].str.startswith("P", na=False)].copy()
    if stock.empty:
        return pd.DataFrame(columns=output_columns), str(path)

    source = parsed.loc[stock.index]
    stock["finished_stock"] = pd.to_numeric(source[end_stock_col], errors="coerce").fillna(0) if end_stock_col else 0
    stock["stock_change"] = pd.to_numeric(source[change_col], errors="coerce").fillna(0) if change_col else 0
    stock["doi_order"] = pd.to_numeric(source[order_col], errors="coerce").fillna(0) if order_col else 0
    stock["doi"] = pd.to_numeric(source[doi_col], errors="coerce").fillna(0) if doi_col else 0
    stock["stock_ratio"] = source[ratio_col].map(clean_scalar) if ratio_col else ""
    stock["stock_signal"] = source[signal_col].map(clean_scalar) if signal_col else ""
    stock["stock_action"] = source[action_col].map(clean_scalar) if action_col else ""

    grouped = (
        stock.groupby("production_code", as_index=False)
        .agg(
            {
                "finished_stock": "sum",
                "stock_change": "sum",
                "doi_order": "sum",
                "doi": lambda s: s[s > 0].median() if (s > 0).any() else 0,
                "stock_ratio": lambda s: summarize_unique(s, head_count=1),
                "stock_signal": summarize_signal,
                "stock_action": lambda s: summarize_unique(s, head_count=1),
            }
        )
    )
    doi_order_mask = grouped["doi_order"] > 0
    grouped.loc[doi_order_mask, "doi"] = grouped.loc[doi_order_mask, "finished_stock"] / grouped.loc[doi_order_mask, "doi_order"] * 181
    grouped["doi"] = pd.to_numeric(grouped["doi"], errors="coerce").fillna(0)
    return grouped[output_columns], str(path)


def main() -> None:
    df = pd.read_csv(SNAPSHOT, low_memory=False)
    data = pd.DataFrame()
    for key, col in SRC.items():
        if col in df.columns:
            data[key] = df[col]
        else:
            data[key] = 0 if key in NUM_COLS else ""
    if "stock_change" not in data.columns:
        data["stock_change"] = 0
    if "stock_ratio" not in data.columns:
        data["stock_ratio"] = ""
    if "stock_action" not in data.columns:
        data["stock_action"] = ""

    for col in TEXT_COLS:
        data[col] = clean_text(data[col])
    for col in NUM_COLS:
        if col not in data.columns:
            data[col] = 0
        data[col] = pd.to_numeric(data[col], errors="coerce").fillna(0)

    data.loc[data["customer_group"].isin(["", "nan", "None"]), "customer_group"] = "거래처 미지정"
    data.loc[data["primary"].isin(["", "nan", "None"]), "primary"] = "기타"
    data = data[data["production_code"].str.startswith("P", na=False)].copy()

    finished_goods, finished_goods_path = read_finished_goods_summary()
    if not finished_goods.empty:
        data = data.merge(finished_goods, on="production_code", how="left", suffixes=("", "_stockfile"))
        for col in ["finished_stock", "stock_change", "doi_order", "doi"]:
            stock_col = f"{col}_stockfile"
            if stock_col in data.columns:
                data[col] = pd.to_numeric(data[stock_col], errors="coerce").combine_first(data[col])
                data = data.drop(columns=[stock_col])
        for col in ["stock_ratio", "stock_signal", "stock_action"]:
            stock_col = f"{col}_stockfile"
            if stock_col in data.columns:
                data[col] = clean_text(data[stock_col]).where(clean_text(data[stock_col]).ne(""), data[col])
                data = data.drop(columns=[stock_col])

    for col in ["finished_stock", "stock_change", "doi_order", "doi"]:
        data[col] = pd.to_numeric(data.get(col, 0), errors="coerce").fillna(0)
    for col in ["stock_ratio", "stock_signal", "stock_action"]:
        data[col] = clean_text(data.get(col, pd.Series("", index=data.index)))

    rows = []
    for rec in data.to_dict(orient="records"):
        rows.append(
            {
                "관": rec["site"],
                "거래처그룹": rec["customer_group"],
                "제품분류": rec["primary"],
                "생산코드": rec["production_code"],
                "분리코드": rec["separation_code"],
                "사출코드": rec["injection_code"],
                "제품명": rec["product"],
                "파워": rec["power"],
                "오더수량1": float(rec["order_qty"]),
                "오더수량2": 0.0,
                "제품부족수량": float(rec["shortage_qty"]),
                "사출부족수량": float(rec["injection_shortage_qty"]),
                "사출재고": float(rec["injection_stock"]),
                "분리재고": float(rec["separation_stock"]),
                "검사접착재고": float(rec["adhesion_stock"]),
                "누수규격검사재고": float(rec["leakage_stock"]),
                "완제품재고": float(rec["finished_stock"]),
                "재고변화": float(rec["stock_change"]),
                "DOI기준오더": float(rec["doi_order"]),
                "DOI": float(rec["doi"]),
                "재고비율": rec["stock_ratio"],
                "재고신호": rec["stock_signal"],
                "대응판단": rec["stock_action"],
                "납기일": rec["due"],
                "이니셜": rec["initial"],
                "기존상태": rec["status"],
                "비고": "",
            }
        )

    customers = sorted({row["거래처그룹"] for row in rows if row["거래처그룹"]})
    payload = {
        "snapshot": str(SNAPSHOT),
        "finished_goods_stock": finished_goods_path,
        "row_count": len(rows),
        "customers": customers,
        "rows": rows,
    }
    OUTPUT.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    print(
        json.dumps(
            {
                "output": str(OUTPUT),
                "rows": len(rows),
                "customers": len(customers),
                "finished_goods_stock": finished_goods_path,
            },
            ensure_ascii=False,
        )
    )


if __name__ == "__main__":
    main()
