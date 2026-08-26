import argparse
import json
import logging
import os
import sys
from datetime import datetime, time, timedelta
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
LOG_PATH = PROJECT_ROOT / "outputs" / "aps_snapshot_refresh.log"
STATE_PATH = PROJECT_ROOT / "outputs" / "aps_snapshot_refresh_state.json"
DEFAULT_SITES = ("C관", "A관", "S관", "전체")
DEFAULT_SLOT_TIMES = ("07:30", "16:00")


def configure_logging() -> None:
    LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[
            logging.FileHandler(LOG_PATH, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )


def parse_sites(value: str) -> list[str]:
    sites = [site.strip() for site in value.split(",") if site.strip()]
    return sites or list(DEFAULT_SITES)


def parse_slot_times(value: str) -> list[time]:
    slots: list[time] = []
    for raw_slot in value.split(","):
        slot_text = raw_slot.strip()
        if not slot_text:
            continue
        slots.append(datetime.strptime(slot_text, "%H:%M").time())
    return sorted(slots)


def current_slot_key(now: datetime, slots: list[time]) -> str:
    candidates = [datetime.combine(now.date(), slot, tzinfo=now.tzinfo) for slot in slots]
    past_candidates = [candidate for candidate in candidates if candidate <= now]
    if past_candidates:
        slot_dt = max(past_candidates)
    else:
        previous_day = now.date() - timedelta(days=1)
        slot_dt = datetime.combine(previous_day, slots[-1], tzinfo=now.tzinfo)
    return slot_dt.strftime("%Y-%m-%d %H:%M")


def read_state() -> dict[str, object]:
    if not STATE_PATH.exists():
        return {"completed_slots": {}}
    try:
        with STATE_PATH.open("r", encoding="utf-8") as state_file:
            state = json.load(state_file)
    except Exception:
        return {"completed_slots": {}}
    if not isinstance(state, dict):
        return {"completed_slots": {}}
    completed_slots = state.get("completed_slots")
    if not isinstance(completed_slots, dict):
        state["completed_slots"] = {}
    return state


def write_state(state: dict[str, object]) -> None:
    STATE_PATH.parent.mkdir(parents=True, exist_ok=True)
    temp_path = STATE_PATH.with_suffix(STATE_PATH.suffix + ".tmp")
    with temp_path.open("w", encoding="utf-8") as state_file:
        json.dump(state, state_file, ensure_ascii=False, indent=2)
    temp_path.replace(STATE_PATH)


def is_slot_completed(state: dict[str, object], slot_key: str) -> bool:
    completed_slots = state.get("completed_slots")
    return isinstance(completed_slots, dict) and slot_key in completed_slots


def mark_slot_completed(
    state: dict[str, object],
    slot_key: str,
    api_updated_at: str,
    results: dict[str, str],
    app_module,
) -> None:
    completed_slots = state.setdefault("completed_slots", {})
    if not isinstance(completed_slots, dict):
        completed_slots = {}
        state["completed_slots"] = completed_slots
    completed_slots[slot_key] = {
        "completed_at": datetime.now(app_module.DISPLAY_TZ).strftime("%Y-%m-%d %H:%M:%S"),
        "api_updated_at": api_updated_at,
        "results": results,
    }
    if len(completed_slots) > 20:
        for old_key in sorted(completed_slots)[:-20]:
            completed_slots.pop(old_key, None)
    write_state(state)


def parse_updated_at(app_module, value: object):
    return app_module.parse_updated_at_value(value)


def is_snapshot_current(app_module, site_filter: str, api_updated_at: str) -> bool:
    api_dt = parse_updated_at(app_module, api_updated_at)
    if api_dt is None:
        return False
    snapshot_dt = parse_updated_at(
        app_module,
        app_module.get_cloud_shortage_snapshot_updated_at(site_filter, "-"),
    )
    if snapshot_dt is None:
        return False
    return snapshot_dt.timestamp() + 1 >= api_dt.timestamp()


def refresh_site(app_module, site_filter: str, api_updated_at: str, only_if_stale: bool) -> str:
    if only_if_stale and is_snapshot_current(app_module, site_filter, api_updated_at):
        return "skip-current"

    refresh_key = app_module.build_api_shortage_refresh_key(app_module.BASE_DIR, site_filter)
    df, file_info_df, process_map_df = app_module.load_api_shortage_data(
        f"{refresh_key}:snapshot-refresh:{datetime.now(app_module.DISPLAY_TZ).isoformat()}",
        str(app_module.BASE_DIR),
        site_filter,
    )
    if df.empty:
        return "skip-empty"

    updated_at = api_updated_at
    if not updated_at or updated_at == "-":
        updated_at = app_module.get_plan_api_updated_at()
    if not updated_at or updated_at == "-":
        updated_at = datetime.now(app_module.DISPLAY_TZ).strftime("%Y-%m-%d %H:%M:%S")

    ok = app_module.write_cloud_shortage_snapshot(df, file_info_df, process_map_df, updated_at, site_filter)
    if not ok:
        raise RuntimeError(f"snapshot write failed: {site_filter}")
    return f"refreshed rows={len(df):,} updated_at={updated_at}"


def main() -> int:
    parser = argparse.ArgumentParser(description="Refresh APS shortage snapshots for the Streamlit dashboard.")
    parser.add_argument(
        "--sites",
        default=",".join(DEFAULT_SITES),
        help="Comma-separated site filters. Default: C관,A관,S관,전체",
    )
    parser.add_argument(
        "--only-if-stale",
        action="store_true",
        help="Skip refresh when the saved snapshot is already current for the APS API timestamp.",
    )
    parser.add_argument(
        "--scheduled",
        action="store_true",
        help="Use slot state so failed scheduled refreshes retry every run until the slot succeeds.",
    )
    parser.add_argument(
        "--slot-times",
        default=",".join(DEFAULT_SLOT_TIMES),
        help="Comma-separated APS refresh slot times in HH:MM. Default: 07:30,16:00",
    )
    args = parser.parse_args()

    configure_logging()
    slot_key = ""
    state: dict[str, object] = {"completed_slots": {}}
    if args.scheduled:
        slots = parse_slot_times(args.slot_times)
        if not slots:
            logging.error("No valid slot times configured: %s", args.slot_times)
            return 2
        state = read_state()
        now = datetime.now().astimezone()
        slot_key = current_slot_key(now, slots)
        if is_slot_completed(state, slot_key):
            logging.info("slot %s already completed; skip without API call", slot_key)
            return 0
        logging.info("slot %s is pending; attempting snapshot refresh", slot_key)

    os.environ.setdefault("INTEROJO_FORCE_PLAN_API", "1")
    sys.path.insert(0, str(PROJECT_ROOT))

    import app  # noqa: PLC0415

    if not app.is_plan_api_configured():
        logging.error("APS API key is not configured.")
        return 2

    api_updated_at = app.get_plan_api_updated_at()
    logging.info("APS API updated_at=%s sites=%s only_if_stale=%s", api_updated_at, args.sites, args.only_if_stale)

    exit_code = 0
    results: dict[str, str] = {}
    for site_filter in parse_sites(args.sites):
        try:
            result = refresh_site(app, site_filter, api_updated_at, args.only_if_stale)
            results[site_filter] = result
            logging.info("%s: %s", site_filter, result)
        except Exception as exc:
            exit_code = 1
            results[site_filter] = f"failed: {exc}"
            logging.exception("%s: refresh failed: %s", site_filter, exc)
    if args.scheduled and exit_code == 0 and slot_key:
        mark_slot_completed(state, slot_key, api_updated_at, results, app)
        logging.info("slot %s completed", slot_key)
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
