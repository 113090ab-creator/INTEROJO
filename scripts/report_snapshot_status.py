from __future__ import annotations

import csv
import gzip
import json
import sys
from datetime import datetime
from pathlib import Path


try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SNAPSHOT_DIR = PROJECT_ROOT / "cloud_snapshots"
STATUS_PATH = SNAPSHOT_DIR / "aps_snapshot_refresh_status.json"
META_PATH = SNAPSHOT_DIR / "snapshot_meta.csv"


def read_json(path: Path) -> dict[str, object]:
    try:
        with path.open("r", encoding="utf-8") as file:
            data = json.load(file)
    except Exception:
        return {}
    return data if isinstance(data, dict) else {}


def read_meta(path: Path) -> dict[str, str]:
    if not path.exists():
        return {}
    try:
        with path.open("r", encoding="utf-8-sig", newline="") as file:
            return {str(row.get("key", "")): str(row.get("value", "")) for row in csv.DictReader(file)}
    except Exception:
        return {}


def count_csv_rows(path: Path) -> int | None:
    opener = gzip.open if path.suffix == ".gz" else open
    try:
        with opener(path, "rt", encoding="utf-8-sig", newline="") as file:
            row_count = sum(1 for _ in file)
    except Exception:
        return None
    return max(row_count - 1, 0)


def print_snapshot_file(path: Path) -> None:
    rows = count_csv_rows(path) if path.suffix in {".csv", ".gz"} or path.name.endswith(".csv.gz") else None
    stat = path.stat()
    modified_at = datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
    row_text = "-" if rows is None else f"{rows:,}"
    print(f"- {path.name}: rows={row_text} size={stat.st_size:,} bytes modified={modified_at}")


def main() -> int:
    status = read_json(STATUS_PATH)
    meta = read_meta(META_PATH)

    print("APS snapshot report")
    print(f"- status: {status.get('status', '-')}")
    print(f"- checked_at(KST): {status.get('checked_at', '-')}")
    print(f"- slot_key: {status.get('slot_key', '-')}")
    print(f"- PLAN api_updated_at: {status.get('api_updated_at', meta.get('data_updated_at', '-'))}")
    print(f"- WIP api_updated_at: {status.get('wip_api_updated_at', meta.get('wip_inventory_updated_at', '-'))}")
    print(f"- snapshot PLAN meta: {meta.get('data_updated_at', '-')}")
    print(f"- snapshot WIP meta: {meta.get('wip_inventory_updated_at', '-')}")

    results = status.get("results")
    if isinstance(results, dict):
        print("- results:")
        for key, value in results.items():
            print(f"  - {key}: {value}")

    print("- files:")
    for path in sorted(SNAPSHOT_DIR.glob("*")):
        if path.is_file() and path.name.endswith((".csv", ".csv.gz", ".json")):
            print_snapshot_file(path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
