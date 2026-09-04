import argparse
import atexit
import json
import logging
import os
import sys
from datetime import datetime, time, timedelta
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
LOG_PATH = PROJECT_ROOT / "outputs" / "aps_snapshot_refresh.log"
STATE_PATH = PROJECT_ROOT / "outputs" / "aps_snapshot_refresh_state.json"
STATUS_PATH = PROJECT_ROOT / "outputs" / "aps_snapshot_refresh_status.json"
CLOUD_STATE_PATH = PROJECT_ROOT / "cloud_snapshots" / "aps_snapshot_refresh_state.json"
CLOUD_STATUS_PATH = PROJECT_ROOT / "cloud_snapshots" / "aps_snapshot_refresh_status.json"
LOCK_PATH = PROJECT_ROOT / "outputs" / "aps_snapshot_refresh.lock"
DEFAULT_SITES = ("C관", "A관", "S관", "전체")
DEFAULT_SLOT_TIMES = ("07:30", "16:00")
DEFAULT_SLOT_GRACE_MINUTES = 10
DEFAULT_SLOT_LOOKAHEAD_MINUTES = 5
LOCK_STALE_MINUTES = 120


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


def current_slot_key(now: datetime, slots: list[time], lookahead_minutes: int = 0) -> str:
    candidates = [datetime.combine(now.date(), slot, tzinfo=now.tzinfo) for slot in slots]
    lookahead_until = now + timedelta(minutes=max(lookahead_minutes, 0))
    upcoming_candidates = [candidate for candidate in candidates if now < candidate <= lookahead_until]
    if upcoming_candidates:
        slot_dt = min(upcoming_candidates)
        return slot_dt.strftime("%Y-%m-%d %H:%M")

    past_candidates = [candidate for candidate in candidates if candidate <= now]
    if past_candidates:
        slot_dt = max(past_candidates)
    else:
        previous_day = now.date() - timedelta(days=1)
        slot_dt = datetime.combine(previous_day, slots[-1], tzinfo=now.tzinfo)
    return slot_dt.strftime("%Y-%m-%d %H:%M")


def read_state() -> dict[str, object]:
    for path in (STATE_PATH, CLOUD_STATE_PATH):
        if not path.exists():
            continue
        try:
            with path.open("r", encoding="utf-8") as state_file:
                state = json.load(state_file)
        except Exception:
            continue
        if not isinstance(state, dict):
            continue
        completed_slots = state.get("completed_slots")
        if not isinstance(completed_slots, dict):
            state["completed_slots"] = {}
        return state
    return {"completed_slots": {}}


def write_json_atomic(path: Path, payload: dict[str, object]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = path.with_suffix(path.suffix + ".tmp")
    with temp_path.open("w", encoding="utf-8") as json_file:
        json.dump(payload, json_file, ensure_ascii=False, indent=2)
    temp_path.replace(path)


def write_state(state: dict[str, object]) -> None:
    write_json_atomic(STATE_PATH, state)
    write_json_atomic(CLOUD_STATE_PATH, state)


def write_status(status: dict[str, object]) -> None:
    write_json_atomic(STATUS_PATH, status)
    write_json_atomic(CLOUD_STATUS_PATH, status)


def release_refresh_lock() -> None:
    try:
        LOCK_PATH.unlink(missing_ok=True)
    except Exception:
        pass


def acquire_refresh_lock() -> bool:
    LOCK_PATH.parent.mkdir(parents=True, exist_ok=True)
    try:
        fd = os.open(str(LOCK_PATH), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    except FileExistsError:
        try:
            lock_age = datetime.now().timestamp() - LOCK_PATH.stat().st_mtime
            if lock_age > LOCK_STALE_MINUTES * 60:
                LOCK_PATH.unlink(missing_ok=True)
                return acquire_refresh_lock()
        except Exception:
            pass
        logging.info("another APS snapshot refresh is already running; skip this run")
        return False

    with os.fdopen(fd, "w", encoding="utf-8") as lock_file:
        lock_file.write(f"pid={os.getpid()} created_at={datetime.now().isoformat()}\n")
    atexit.register(release_refresh_lock)
    return True


def is_slot_completed(state: dict[str, object], slot_key: str) -> bool:
    completed_slots = state.get("completed_slots")
    if not isinstance(completed_slots, dict) or slot_key not in completed_slots:
        return False
    slot_info = completed_slots.get(slot_key)
    if not isinstance(slot_info, dict):
        return False
    return bool(str(slot_info.get("wip_api_updated_at", "")).strip())


def mark_slot_completed(
    state: dict[str, object],
    slot_key: str,
    api_updated_at: str,
    wip_api_updated_at: str,
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
        "wip_api_updated_at": wip_api_updated_at,
        "results": results,
    }
    if len(completed_slots) > 20:
        for old_key in sorted(completed_slots)[:-20]:
            completed_slots.pop(old_key, None)
    write_state(state)


def parse_updated_at(app_module, value: object):
    return app_module.parse_updated_at_value(value)


def parse_slot_key_datetime(app_module, slot_key: str):
    try:
        return datetime.strptime(slot_key, "%Y-%m-%d %H:%M").replace(tzinfo=app_module.DISPLAY_TZ)
    except ValueError:
        return None


def is_api_update_ready_for_slot(
    app_module,
    api_updated_at: str,
    slot_key: str,
    grace_minutes: int,
) -> bool:
    api_dt = parse_updated_at(app_module, api_updated_at)
    slot_dt = parse_slot_key_datetime(app_module, slot_key)
    if api_dt is None or slot_dt is None:
        return False
    return api_dt >= slot_dt - timedelta(minutes=max(grace_minutes, 0))


def is_snapshot_current(app_module, site_filter: str, api_updated_at: str, wip_api_updated_at: str = "") -> bool:
    api_dt = parse_updated_at(app_module, api_updated_at)
    if api_dt is None:
        return False
    snapshot_dt = parse_updated_at(
        app_module,
        app_module.get_cloud_shortage_snapshot_updated_at(site_filter, "-"),
    )
    if snapshot_dt is None:
        return False
    if snapshot_dt.timestamp() + 1 < api_dt.timestamp():
        return False

    wip_text = str(wip_api_updated_at or "").strip()
    if wip_text and wip_text != "-":
        wip_label = app_module.get_cloud_shortage_wip_source_label(site_filter)
        if app_module.format_reference_timestamp(wip_text) not in wip_label:
            return False
    return True


def is_wip_ready_for_plan(app_module, plan_updated_at: str, wip_updated_at: str) -> bool:
    plan_dt = parse_updated_at(app_module, plan_updated_at)
    wip_dt = parse_updated_at(app_module, wip_updated_at)
    if plan_dt is None or wip_dt is None:
        return False
    return wip_dt.timestamp() + 1 >= plan_dt.timestamp()


def format_wip_pending_message(plan_updated_at: str, wip_updated_at: str) -> str:
    return (
        "APS WIP API 기준시각이 APS PLAN 기준시각보다 오래되어 WIP 데이터 갱신을 기다립니다. "
        f"PLAN={plan_updated_at}, WIP={wip_updated_at}"
    )


def refresh_wip_inventory(app_module, only_if_stale: bool) -> tuple[str, str]:
    result = app_module.refresh_cloud_wip_inventory_snapshot(only_if_stale=only_if_stale)
    updated_at = str(result.get("updated_at", "-"))
    status = str(result.get("status", "unknown"))
    rows = int(result.get("rows", 0) or 0)
    return updated_at, f"{status} rows={rows:,} updated_at={updated_at}"


def refresh_site(app_module, site_filter: str, api_updated_at: str, wip_api_updated_at: str, only_if_stale: bool) -> str:
    if only_if_stale and is_snapshot_current(app_module, site_filter, api_updated_at, wip_api_updated_at):
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
    parser.add_argument(
        "--slot-grace-minutes",
        type=int,
        default=DEFAULT_SLOT_GRACE_MINUTES,
        help="How many minutes before a slot an APS updated_at is allowed to count for that slot. Default: 10",
    )
    parser.add_argument(
        "--slot-lookahead-minutes",
        type=int,
        default=DEFAULT_SLOT_LOOKAHEAD_MINUTES,
        help="Start checking an upcoming slot this many minutes before the configured slot time. Default: 5",
    )
    parser.add_argument(
        "--wip-only",
        action="store_true",
        help="Refresh only the APS WIP raw/normalized inventory snapshots.",
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
        slot_key = current_slot_key(now, slots, args.slot_lookahead_minutes)
        if is_slot_completed(state, slot_key):
            logging.debug("slot %s already completed; skip without API call", slot_key)
            return 0
        logging.info("slot %s is pending; attempting snapshot refresh", slot_key)
        if not acquire_refresh_lock():
            return 0

    os.environ.setdefault("INTEROJO_FORCE_PLAN_API", "1")
    sys.path.insert(0, str(PROJECT_ROOT))

    import app  # noqa: PLC0415

    if not app.is_plan_api_configured():
        logging.error("APS API key is not configured.")
        return 2

    api_updated_at = app.get_plan_api_updated_at()
    wip_api_updated_at = "-"
    logging.info("APS API updated_at=%s sites=%s only_if_stale=%s", api_updated_at, args.sites, args.only_if_stale)
    if args.scheduled and slot_key:
        write_status(
            {
                "checked_at": datetime.now(app.DISPLAY_TZ).strftime("%Y-%m-%d %H:%M:%S"),
                "slot_key": slot_key,
                "api_updated_at": api_updated_at,
                "wip_api_updated_at": wip_api_updated_at,
                "status": "checking",
                "sites": parse_sites(args.sites),
            }
        )
    if args.scheduled and slot_key and not is_api_update_ready_for_slot(
        app,
        api_updated_at,
        slot_key,
        args.slot_grace_minutes,
    ):
        logging.warning(
            "slot %s remains pending: APS API updated_at=%s is older than the slot window "
            "(grace=%s minutes); retry on next scheduled run",
            slot_key,
            api_updated_at,
            args.slot_grace_minutes,
        )
        write_status(
            {
                "checked_at": datetime.now(app.DISPLAY_TZ).strftime("%Y-%m-%d %H:%M:%S"),
                "slot_key": slot_key,
                "api_updated_at": api_updated_at,
                "wip_api_updated_at": wip_api_updated_at,
                "status": "pending_api_update",
                "sites": parse_sites(args.sites),
            }
        )
        return 0

    exit_code = 0
    results: dict[str, str] = {}
    try:
        wip_api_updated_at, wip_result = refresh_wip_inventory(app, args.only_if_stale)
        results["WIP"] = wip_result
        logging.info("WIP: %s", wip_result)
        if not is_wip_ready_for_plan(app, api_updated_at, wip_api_updated_at):
            exit_code = 1
            results["WIP"] = format_wip_pending_message(api_updated_at, wip_api_updated_at)
            logging.warning(results["WIP"])
            if args.scheduled and slot_key:
                write_status(
                    {
                        "checked_at": datetime.now(app.DISPLAY_TZ).strftime("%Y-%m-%d %H:%M:%S"),
                        "slot_key": slot_key,
                        "api_updated_at": api_updated_at,
                        "wip_api_updated_at": wip_api_updated_at,
                        "status": "pending_wip_update",
                        "sites": parse_sites(args.sites),
                        "results": results,
                    }
                )
            return exit_code
    except Exception as exc:
        exit_code = 1
        results["WIP"] = f"failed: {exc}"
        logging.exception("WIP refresh failed: %s", exc)
        if args.scheduled and slot_key:
            write_status(
                {
                    "checked_at": datetime.now(app.DISPLAY_TZ).strftime("%Y-%m-%d %H:%M:%S"),
                    "slot_key": slot_key,
                    "api_updated_at": api_updated_at,
                    "wip_api_updated_at": wip_api_updated_at,
                    "status": "failed",
                    "sites": parse_sites(args.sites),
                    "results": results,
                }
            )
        return exit_code

    if args.wip_only:
        if args.scheduled and exit_code == 0 and slot_key:
            mark_slot_completed(state, slot_key, api_updated_at, wip_api_updated_at, results, app)
            write_status(
                {
                    "checked_at": datetime.now(app.DISPLAY_TZ).strftime("%Y-%m-%d %H:%M:%S"),
                    "slot_key": slot_key,
                    "api_updated_at": api_updated_at,
                    "wip_api_updated_at": wip_api_updated_at,
                    "status": "completed",
                    "sites": [],
                    "results": results,
                }
            )
        return 0

    for site_filter in parse_sites(args.sites):
        try:
            result = refresh_site(app, site_filter, api_updated_at, wip_api_updated_at, args.only_if_stale)
            results[site_filter] = result
            logging.info("%s: %s", site_filter, result)
        except Exception as exc:
            exit_code = 1
            results[site_filter] = f"failed: {exc}"
            logging.exception("%s: refresh failed: %s", site_filter, exc)
    if args.scheduled and exit_code == 0 and slot_key:
        mark_slot_completed(state, slot_key, api_updated_at, wip_api_updated_at, results, app)
        write_status(
            {
                "checked_at": datetime.now(app.DISPLAY_TZ).strftime("%Y-%m-%d %H:%M:%S"),
                "slot_key": slot_key,
                "api_updated_at": api_updated_at,
                "wip_api_updated_at": wip_api_updated_at,
                "status": "completed",
                "sites": parse_sites(args.sites),
                "results": results,
            }
        )
        logging.info("slot %s completed", slot_key)
    elif args.scheduled and slot_key:
        write_status(
            {
                "checked_at": datetime.now(app.DISPLAY_TZ).strftime("%Y-%m-%d %H:%M:%S"),
                "slot_key": slot_key,
                "api_updated_at": api_updated_at,
                "wip_api_updated_at": wip_api_updated_at,
                "status": "failed",
                "sites": parse_sites(args.sites),
                "results": results,
            }
        )
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
