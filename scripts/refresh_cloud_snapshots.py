from __future__ import annotations

import os
import sys
from pathlib import Path

import pandas as pd


REPO_ROOT = Path(__file__).resolve().parents[1]
os.chdir(REPO_ROOT)
sys.path.insert(0, str(REPO_ROOT))

import app  # noqa: E402


def write_snapshot(name: str, df: pd.DataFrame) -> None:
    path = REPO_ROOT / "cloud_snapshots" / name
    compression: str | dict[str, object] = "infer"
    if name.endswith(".gz"):
        compression = {"method": "gzip", "mtime": 0}
    df.to_csv(path, index=False, encoding="utf-8-sig", compression=compression)
    print(f"wrote {name}: shape={df.shape} bytes={path.stat().st_size}")


def load_existing_meta(snapshot_dir: Path) -> dict[str, str]:
    path = snapshot_dir / "snapshot_meta.csv"
    if not path.exists():
        return {}
    try:
        meta = pd.read_csv(path, encoding="utf-8-sig")
    except Exception:
        return {}
    if meta.empty or not {"key", "value"}.issubset(meta.columns):
        return {}
    return {str(row["key"]): str(row["value"]) for _, row in meta.iterrows()}


def main() -> None:
    snapshot_dir = REPO_ROOT / "cloud_snapshots"
    snapshot_dir.mkdir(exist_ok=True)
    meta_values = load_existing_meta(snapshot_dir)

    try:
        data_refresh_key = app.build_data_refresh_key(REPO_ROOT)
        shortage_df, file_info_df, process_map_df = app.load_data(data_refresh_key, str(REPO_ROOT))
        inventory_risk_df = app.build_inventory_risk_snapshot(data_refresh_key, str(REPO_ROOT))
    except Exception as exc:
        print(f"skipped data snapshots: {exc}")
    else:
        write_snapshot("shortage_snapshot.csv.gz", shortage_df)
        write_snapshot("shortage_file_info.csv.gz", file_info_df)
        write_snapshot("process_map.csv.gz", process_map_df)
        write_snapshot("inventory_risk_snapshot.csv.gz", inventory_risk_df)
        meta_values["data_updated_at"] = app.get_data_updated_at(REPO_ROOT)

    try:
        all_item_refresh_key = app.build_all_item_refresh_key(REPO_ROOT)
        all_item_df, code_mismatch_df = app.build_all_item_status_snapshot(all_item_refresh_key, str(REPO_ROOT))
    except Exception as exc:
        print(f"skipped all item snapshots: {exc}")
    else:
        write_snapshot(app.ALL_ITEM_SNAPSHOT_FILE, all_item_df)
        write_snapshot(app.CODE_MISMATCH_SNAPSHOT_FILE, code_mismatch_df)
        meta_values["all_item_updated_at"] = app.get_all_item_updated_at(REPO_ROOT)

    try:
        leadji_refresh_key = app.build_leadji_status_refresh_key(REPO_ROOT)
        leadji_shortage_df, leadji_info_df, leadji_stock_df, leadji_order_df = app.load_leadji_status_snapshot(
            leadji_refresh_key, str(REPO_ROOT)
        )
    except Exception as exc:
        print(f"skipped leadji snapshots: {exc}")
    else:
        write_snapshot("leadji_shortage_snapshot.csv.gz", leadji_shortage_df)
        write_snapshot("leadji_info.csv.gz", leadji_info_df)
        write_snapshot("leadji_stock.csv.gz", leadji_stock_df)
        write_snapshot("leadji_order.csv.gz", leadji_order_df)
        meta_values["leadji_updated_at"] = app.get_leadji_status_updated_at(REPO_ROOT)

    meta = pd.DataFrame(
        [
            {"key": "data_updated_at", "value": meta_values.get("data_updated_at", "-")},
            {"key": "all_item_updated_at", "value": meta_values.get("all_item_updated_at", "-")},
            {"key": "leadji_updated_at", "value": meta_values.get("leadji_updated_at", "-")},
        ]
    )
    meta.to_csv(snapshot_dir / "snapshot_meta.csv", index=False, encoding="utf-8-sig")
    print(meta.to_string(index=False))


if __name__ == "__main__":
    main()
