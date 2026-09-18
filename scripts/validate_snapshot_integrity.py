from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_SNAPSHOT_DIR = PROJECT_ROOT / "cloud_snapshots"
CURRENT_SNAPSHOT_SET_NAME = "current_snapshot_set.json"
STATUS_NAME = "aps_snapshot_refresh_status.json"
META_NAME = "snapshot_meta.csv"
MANIFEST_NAME = "manifest.json"
SETS_DIR_NAME = "sets"
CONFLICT_MARKERS = ("<<<<<<<", "=======", ">>>>>>>")
ALLOWED_STATUS_VALUES = {
    "WAITING_FOR_PLAN",
    "WAITING_FOR_WIP",
    "DELAYED",
    "FAILED",
    "PUBLISHED",
    "READY",
    "BUILDING",
    "VALIDATING",
    "PUBLISHING",
}


class SnapshotIntegrityError(RuntimeError):
    pass


def read_json_file(path: Path) -> dict[str, object]:
    if not path.exists():
        raise SnapshotIntegrityError(f"missing required JSON file: {path}")
    text = path.read_text(encoding="utf-8-sig")
    for marker in CONFLICT_MARKERS:
        if marker in text:
            raise SnapshotIntegrityError(f"git conflict marker found in {path}: {marker}")
    try:
        payload = json.loads(text)
    except json.JSONDecodeError as exc:
        raise SnapshotIntegrityError(f"invalid JSON in {path}: {exc}") from exc
    if not isinstance(payload, dict):
        raise SnapshotIntegrityError(f"JSON file must contain an object: {path}")
    return payload


def validate_no_conflict_markers(path: Path) -> None:
    if not path.exists():
        return
    text = path.read_text(encoding="utf-8-sig")
    for marker in CONFLICT_MARKERS:
        if marker in text:
            raise SnapshotIntegrityError(f"git conflict marker found in {path}: {marker}")


def clean_text(value: object) -> str:
    return str(value or "").strip()


def require_fields(path: Path, payload: dict[str, object], fields: set[str]) -> None:
    missing = sorted(field for field in fields if not clean_text(payload.get(field, "")))
    if missing:
        raise SnapshotIntegrityError(f"{path} missing required fields: {', '.join(missing)}")


def validate_all_json_files(snapshot_dir: Path) -> None:
    json_paths = [snapshot_dir / CURRENT_SNAPSHOT_SET_NAME, snapshot_dir / STATUS_NAME]
    json_paths.extend(sorted((snapshot_dir / SETS_DIR_NAME).glob(f"*/{MANIFEST_NAME}")))
    for path in json_paths:
        if path.exists():
            read_json_file(path)


def validate_current_set(snapshot_dir: Path) -> tuple[dict[str, object], dict[str, object]]:
    current_path = snapshot_dir / CURRENT_SNAPSHOT_SET_NAME
    current = read_json_file(current_path)
    require_fields(
        current_path,
        current,
        {"set_id", "slot_key", "manifest_path", "plan_updated_at", "wip_updated_at", "published_at"},
    )

    manifest_path = snapshot_dir / clean_text(current.get("manifest_path", ""))
    manifest = read_json_file(manifest_path)
    require_fields(manifest_path, manifest, {"set_id", "slot_key", "plan_updated_at", "wip_updated_at", "created_at"})

    comparable_fields = ("set_id", "slot_key", "plan_updated_at", "wip_updated_at")
    for field in comparable_fields:
        current_value = clean_text(current.get(field, ""))
        manifest_value = clean_text(manifest.get(field, ""))
        if current_value != manifest_value:
            raise SnapshotIntegrityError(
                f"current set and manifest disagree on {field}: current={current_value!r}, manifest={manifest_value!r}"
            )
    return current, manifest


def validate_status(snapshot_dir: Path, current: dict[str, object]) -> None:
    status_path = snapshot_dir / STATUS_NAME
    if not status_path.exists():
        return
    status = read_json_file(status_path)
    status_value = clean_text(status.get("status", "")).upper()
    if status_value not in ALLOWED_STATUS_VALUES:
        raise SnapshotIntegrityError(f"{status_path} has unsupported status value: {status.get('status')!r}")

    if status_value != "PUBLISHED":
        return

    status_set_id = clean_text(status.get("set_id", ""))
    status_slot_key = clean_text(status.get("slot_key", ""))
    current_set_id = clean_text(current.get("set_id", ""))
    current_slot_key = clean_text(current.get("slot_key", ""))
    if status_set_id and status_set_id != current_set_id:
        raise SnapshotIntegrityError(
            f"published status set_id does not match current set: status={status_set_id!r}, current={current_set_id!r}"
        )
    if status_slot_key and status_slot_key != current_slot_key:
        raise SnapshotIntegrityError(
            f"published status slot_key does not match current set: status={status_slot_key!r}, current={current_slot_key!r}"
        )


def validate_snapshot_integrity(snapshot_dir: Path) -> None:
    if not snapshot_dir.exists():
        raise SnapshotIntegrityError(f"snapshot directory does not exist: {snapshot_dir}")
    validate_no_conflict_markers(snapshot_dir / META_NAME)
    validate_all_json_files(snapshot_dir)
    current, _manifest = validate_current_set(snapshot_dir)
    validate_status(snapshot_dir, current)


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate production APS snapshot JSON integrity.")
    parser.add_argument("--snapshot-dir", type=Path, default=DEFAULT_SNAPSHOT_DIR)
    args = parser.parse_args()
    try:
        validate_snapshot_integrity(args.snapshot_dir)
    except SnapshotIntegrityError as exc:
        print(f"snapshot integrity validation failed: {exc}", file=sys.stderr)
        return 1
    print(f"snapshot integrity validation passed: {args.snapshot_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
