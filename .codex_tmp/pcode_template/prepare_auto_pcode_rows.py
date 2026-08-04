# -*- coding: utf-8 -*-
import json
from pathlib import Path

import pandas as pd


BASE_DIR = Path(__file__).resolve().parents[2]
SNAPSHOT = BASE_DIR / "cloud_snapshots" / "all_item_status_snapshot.csv.gz"
OUTPUT = Path(__file__).resolve().parent / "pcode_auto_rows.json"

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
    "doi": "DOI",
    "signal": "신호",
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
    "signal",
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
    "doi",
]


def clean_text(series: pd.Series) -> pd.Series:
    return (
        series.fillna("")
        .astype(str)
        .replace({"nan": "", "None": "", "NaT": "", "<NA>": ""})
        .str.strip()
    )


def main() -> None:
    df = pd.read_csv(SNAPSHOT, low_memory=False)
    data = pd.DataFrame()
    for key, col in SRC.items():
        if col in df.columns:
            data[key] = df[col]
        else:
            data[key] = 0 if key in NUM_COLS else ""

    for col in TEXT_COLS:
        data[col] = clean_text(data[col])
    for col in NUM_COLS:
        data[col] = pd.to_numeric(data[col], errors="coerce").fillna(0)

    data.loc[data["customer_group"].isin(["", "nan", "None"]), "customer_group"] = "거래처 미지정"
    data.loc[data["primary"].isin(["", "nan", "None"]), "primary"] = "기타"
    data = data[data["production_code"].str.startswith("P", na=False)].copy()

    rows = []
    for rec in data.to_dict(orient="records"):
        rows.append(
            {
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
                "DOI": float(rec["doi"]),
                "납기일": rec["due"],
                "이니셜": rec["initial"],
                "신호": rec["signal"],
                "기존상태": rec["status"],
                "비고": "",
            }
        )

    customers = sorted({row["거래처그룹"] for row in rows if row["거래처그룹"]})
    payload = {
        "snapshot": str(SNAPSHOT),
        "row_count": len(rows),
        "customers": customers,
        "rows": rows,
    }
    OUTPUT.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    print(json.dumps({"output": str(OUTPUT), "rows": len(rows), "customers": len(customers)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
