from __future__ import annotations

import os
import sys
from pathlib import Path

import pandas as pd


REPO_ROOT = Path(__file__).resolve().parents[1]
os.chdir(REPO_ROOT)
sys.path.insert(0, str(REPO_ROOT))

import app  # noqa: E402


VSS_MANAGED_APS_SNAPSHOT_NAMES = frozenset(
    {
        "current_snapshot_set.json",
        "aps_snapshot_refresh_status.json",
        "aps_snapshot_refresh_state.json",
        "wip_inventory_snapshot.csv.gz",
        "shortage_snapshot.csv.gz",
        "shortage_file_info.csv.gz",
        "process_map.csv.gz",
        "shortage_snapshot_asite.csv.gz",
        "shortage_file_info_asite.csv.gz",
        "process_map_asite.csv.gz",
        "shortage_snapshot_csite.csv.gz",
        "shortage_file_info_csite.csv.gz",
        "process_map_csite.csv.gz",
        "shortage_snapshot_ssite.csv.gz",
        "shortage_file_info_ssite.csv.gz",
        "process_map_ssite.csv.gz",
    }
)


def is_vss_managed_aps_snapshot_name(name: str) -> bool:
    snapshot_name = name.replace("\\", "/").lstrip("/")
    return snapshot_name.startswith("sets/") or snapshot_name in VSS_MANAGED_APS_SNAPSHOT_NAMES


def write_snapshot(name: str, df: pd.DataFrame) -> None:
    if is_vss_managed_aps_snapshot_name(name):
        raise RuntimeError(f"VSS-managed APS snapshot writes are blocked: {name}")
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


def build_snapshot_meta_frame(meta_values: dict[str, str]) -> pd.DataFrame:
    preferred_keys = ["data_updated_at", "all_item_updated_at", "leadji_updated_at"]
    rows: list[dict[str, str]] = []
    seen: set[str] = set()

    for key in preferred_keys:
        rows.append({"key": key, "value": meta_values.get(key, "-")})
        seen.add(key)

    for key, value in meta_values.items():
        if key in seen:
            continue
        rows.append({"key": key, "value": value})
        seen.add(key)

    return pd.DataFrame(rows)


def main() -> None:
    snapshot_dir = REPO_ROOT / "cloud_snapshots"
    snapshot_dir.mkdir(exist_ok=True)
    meta_values = load_existing_meta(snapshot_dir)

    print("skipped APS shortage snapshots: managed by scripts/refresh_snapshot.py")
    try:
        data_refresh_key = app.build_data_refresh_key(REPO_ROOT)
        inventory_risk_df = app.build_inventory_risk_snapshot(data_refresh_key, str(REPO_ROOT))
    except Exception as exc:
        print(f"skipped inventory risk snapshot: {exc}")
    else:
        write_snapshot("inventory_risk_snapshot.csv.gz", inventory_risk_df)

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

    meta = build_snapshot_meta_frame(meta_values)
    meta.to_csv(snapshot_dir / "snapshot_meta.csv", index=False, encoding="utf-8-sig")
    print(meta.to_string(index=False))


if __name__ == "__main__":
    main()
