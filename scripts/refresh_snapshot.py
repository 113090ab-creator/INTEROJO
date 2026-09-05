from __future__ import annotations

import argparse
import logging
import os
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]
LOG_PATH = PROJECT_ROOT / "outputs" / "refresh_snapshot.log"
STATUS_SNAPSHOT_NAME = "aps_snapshot_refresh_status.json"
DEFAULT_SITES = ("C관", "A관", "S관", "전체")
try:
    MIN_ROW_COUNT_RATIO = float(os.getenv("SNAPSHOT_MIN_ROW_COUNT_RATIO", "0.5"))
except ValueError:
    MIN_ROW_COUNT_RATIO = 0.5
try:
    SNAPSHOT_WAIT_DELAY_MINUTES = int(os.getenv("SNAPSHOT_WAIT_DELAY_MINUTES", "60"))
except ValueError:
    SNAPSHOT_WAIT_DELAY_MINUTES = 60
SHORTAGE_REQUIRED_COLUMNS = {
    "거래처",
    "이니셜",
    "품목코드",
    "제품명",
    "납기일",
    "부족수량",
    "R코드",
    "Q코드",
    "공정재고 합계",
}
FILE_INFO_REQUIRED_COLUMNS = {"재고파일", "수요파일", "행수(현황표)"}


class SnapshotRefreshError(RuntimeError):
    pass


class SnapshotWaiting(RuntimeError):
    def __init__(
        self,
        status: str,
        message: str,
        plan_updated_at: str,
        wip_updated_at: str,
        slot_key: str = "",
    ) -> None:
        super().__init__(message)
        self.status = status
        self.plan_updated_at = plan_updated_at
        self.wip_updated_at = wip_updated_at
        self.slot_key = slot_key


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


def clean_log_text(value: object, max_length: int = 260) -> str:
    text = str(value or "").strip()
    text = re.sub(r"https?://\S+", "[URL]", text)
    text = re.sub(r"(?i)(x-api-key|api[_-]?key|authorization|token)\s*[:=]\s*[^;\s]+", r"\1=[REDACTED]", text)
    if len(text) > max_length:
        return text[: max_length - 3].rstrip() + "..."
    return text


def parse_sites(value: str) -> list[str]:
    sites = [site.strip() for site in value.split(",") if site.strip()]
    return sites or list(DEFAULT_SITES)


def validate_dataframe(name: str, df: pd.DataFrame, required_columns: set[str]) -> None:
    if not isinstance(df, pd.DataFrame):
        raise SnapshotRefreshError(f"{name}: DataFrame이 아닙니다.")
    if df.empty:
        raise SnapshotRefreshError(f"{name}: 0건 응답입니다.")
    missing = sorted(required_columns.difference(df.columns))
    if missing:
        raise SnapshotRefreshError(f"{name}: 필수 컬럼 누락: {', '.join(missing)}")


def validate_wip_snapshot(app, inventory: pd.DataFrame) -> None:
    validate_dataframe("WIP", inventory, set(app.WIP_INVENTORY_COLUMNS))
    if app.parse_mixed_numeric(inventory["재고량"]).fillna(0).sum() <= 0:
        raise SnapshotRefreshError("WIP: 재고량 합계가 0입니다.")


def validate_plan_snapshot(site: str, raw: pd.DataFrame) -> None:
    if not isinstance(raw, pd.DataFrame) or raw.empty:
        raise SnapshotRefreshError(f"{site} PLAN: 0건 응답입니다.")


def validate_shortage_snapshot(site: str, df: pd.DataFrame, file_info_df: pd.DataFrame) -> None:
    validate_dataframe(f"{site} 생산부족", df, SHORTAGE_REQUIRED_COLUMNS)
    validate_dataframe(f"{site} 파일정보", file_info_df, FILE_INFO_REQUIRED_COLUMNS)


def validate_row_count_guard(name: str, new_rows: int, existing_rows: int) -> None:
    if MIN_ROW_COUNT_RATIO <= 0 or existing_rows <= 0:
        return
    minimum_rows = int(existing_rows * MIN_ROW_COUNT_RATIO)
    if new_rows < minimum_rows:
        raise SnapshotRefreshError(
            f"{name}: 행수가 기존 정상 스냅샷 대비 급감했습니다. "
            f"new={new_rows:,}, existing={existing_rows:,}, minimum={minimum_rows:,}."
        )


def load_app_module():
    os.chdir(PROJECT_ROOT)
    sys.path.insert(0, str(PROJECT_ROOT))
    os.environ.setdefault("INTEROJO_FORCE_PLAN_API", "1")
    import app  # noqa: PLC0415

    return app


def comparable_status(payload: dict[str, object]) -> dict[str, object]:
    return {key: value for key, value in payload.items() if key != "checked_at"}


def write_status(app, status: str, **payload: object) -> None:
    import snapshot_storage

    status_payload: dict[str, object] = {
        "checked_at": datetime.now(app.DISPLAY_TZ).strftime("%Y-%m-%d %H:%M:%S"),
        "status": status,
        **payload,
    }
    current_manifest = app.get_published_snapshot_set_manifest()
    if current_manifest:
        status_payload["current_set"] = {
            "set_id": current_manifest.get("set_id", ""),
            "slot_key": current_manifest.get("slot_key", ""),
            "plan_updated_at": current_manifest.get("plan_updated_at", ""),
            "wip_updated_at": current_manifest.get("wip_updated_at", ""),
            "created_at": current_manifest.get("created_at", ""),
        }
    try:
        existing = snapshot_storage.read_json_snapshot(app.CLOUD_SNAPSHOT_DIR, STATUS_SNAPSHOT_NAME)
        if comparable_status(existing) == comparable_status(status_payload):
            return
        snapshot_storage.write_json_snapshot_atomic(
            app.CLOUD_SNAPSHOT_DIR,
            STATUS_SNAPSHOT_NAME,
            status_payload,
        )
    except Exception as exc:
        logging.warning("status write failed: %s", clean_log_text(exc))


def validate_existing_snapshots(app, sites: list[str]) -> None:
    inventory, _ = app.load_cloud_wip_inventory_snapshot_with_label()
    validate_wip_snapshot(app, inventory)
    logging.info("validated existing WIP snapshot rows=%s", f"{len(inventory):,}")

    manifest = app.get_published_snapshot_set_manifest()
    if manifest:
        logging.info(
            "validated current snapshot set=%s slot=%s plan=%s wip=%s",
            manifest.get("set_id", ""),
            manifest.get("slot_key", ""),
            manifest.get("plan_updated_at", ""),
            manifest.get("wip_updated_at", ""),
        )

    for site in sites:
        df, file_info_df, _ = app.load_cloud_shortage_snapshot(site)
        validate_shortage_snapshot(site, df, file_info_df)
        logging.info("validated existing shortage snapshot site=%s rows=%s", site, f"{len(df):,}")


def latest_observed_minutes_old(app, plan_updated_at: str, wip_updated_at: str) -> float | None:
    candidates = [
        parsed
        for parsed in (app.parse_updated_at_value(plan_updated_at), app.parse_updated_at_value(wip_updated_at))
        if parsed is not None
    ]
    if not candidates:
        return None
    latest = max(candidates, key=lambda item: item.timestamp())
    return (datetime.now(app.DISPLAY_TZ) - latest).total_seconds() / 60


def maybe_delayed_status(app, waiting_status: str, plan_updated_at: str, wip_updated_at: str) -> str:
    age_minutes = latest_observed_minutes_old(app, plan_updated_at, wip_updated_at)
    if age_minutes is not None and age_minutes > SNAPSHOT_WAIT_DELAY_MINUTES:
        return app.REFRESH_STATUS_DELAYED
    return waiting_status


def raise_if_slots_not_ready(app, plan_updated_at: str, wip_updated_at: str) -> None:
    slot_status = app.compare_aps_snapshot_slots(plan_updated_at, wip_updated_at)
    if slot_status == app.REFRESH_STATUS_READY:
        return
    slot_key = app.get_latest_aps_snapshot_slot(plan_updated_at, wip_updated_at)
    status = maybe_delayed_status(app, slot_status, plan_updated_at, wip_updated_at)
    if status == app.REFRESH_STATUS_DELAYED:
        message = (
            "PLAN/WIP 같은 회차 데이터 대기 시간이 기준을 초과했습니다. "
            f"PLAN={plan_updated_at}, WIP={wip_updated_at}."
        )
    elif slot_status == app.REFRESH_STATUS_WAITING_FOR_WIP:
        message = f"PLAN={plan_updated_at}, WIP={wip_updated_at}. WIP 같은 회차 데이터를 기다립니다."
    else:
        message = f"PLAN={plan_updated_at}, WIP={wip_updated_at}. PLAN 같은 회차 데이터를 기다립니다."
    raise SnapshotWaiting(status, message, plan_updated_at, wip_updated_at, slot_key)


def current_snapshot_set_matches(app, sites: list[str], plan_updated_at: str, wip_updated_at: str) -> bool:
    manifest = app.get_published_snapshot_set_manifest()
    if not manifest:
        return False
    if str(manifest.get("plan_updated_at", "")).strip() != str(plan_updated_at).strip():
        return False
    if str(manifest.get("wip_updated_at", "")).strip() != str(wip_updated_at).strip():
        return False
    for site in sites:
        snapshot_df, file_info_df, _ = app.load_cloud_shortage_snapshot(site)
        if snapshot_df.empty or file_info_df.empty:
            return False
    return True


def file_info_value(file_info_df: pd.DataFrame, column: str) -> str:
    if not isinstance(file_info_df, pd.DataFrame) or file_info_df.empty or column not in file_info_df.columns:
        return ""
    return str(file_info_df.iloc[0].get(column, "")).strip()


def validate_site_uses_staged_sources(
    app,
    site: str,
    file_info_df: pd.DataFrame,
    plan_updated_at: str,
    wip_updated_at: str,
) -> None:
    demand_source = file_info_value(file_info_df, "수요파일")
    stock_source = file_info_value(file_info_df, "재고파일")
    plan_text = app.format_reference_timestamp(plan_updated_at)
    wip_text = app.format_reference_timestamp(wip_updated_at)
    if plan_text not in demand_source:
        raise SnapshotRefreshError(f"{site}: 생산부족 결과가 staged PLAN 기준시각을 기록하지 않았습니다.")
    if wip_text not in stock_source:
        raise SnapshotRefreshError(f"{site}: 생산부족 결과가 staged WIP 기준시각을 기록하지 않았습니다.")


def is_transient_source_error(error_text: object) -> bool:
    text = str(error_text or "")
    transient_tokens = [
        "HTTP 524",
        "HTTP 520",
        "HTTP 522",
        "HTTP 503",
        "did not contain data rows",
        "response was truncated",
        "응답에서 행 데이터를 찾지 못했습니다",
        "응답이 비어",
        "대상 창고 데이터를 찾지 못했습니다",
    ]
    return any(token in text for token in transient_tokens)


def build_validated_snapshot_set(
    app,
    sites: list[str],
    plan_updated_at: str,
    wip_updated_at: str,
) -> tuple[
    pd.DataFrame,
    str,
    dict[str, pd.DataFrame],
    dict[str, tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]],
    dict[str, str],
]:
    existing_inventory, _ = app.load_cloud_wip_inventory_snapshot_with_label()
    inventory, source_label, raw_snapshot_dir, wip_error = app.build_wip_inventory_snapshot_from_api(wip_updated_at)
    if wip_error:
        raise SnapshotRefreshError(f"APS WIP API 조회 실패: {clean_log_text(wip_error)}")
    validate_wip_snapshot(app, inventory)
    validate_row_count_guard("WIP", len(inventory), len(existing_inventory))
    logging.info(
        "staged WIP snapshot rows=%s source=%s raw_dir=%s",
        f"{len(inventory):,}",
        clean_log_text(source_label),
        clean_log_text(raw_snapshot_dir),
    )

    plan_frames: dict[str, pd.DataFrame] = {}
    shortage_results: dict[str, tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]] = {}
    results: dict[str, str] = {"WIP": f"staged rows={len(inventory):,} updated_at={wip_updated_at}"}
    for site in sites:
        plan_raw, plan_error = app.read_aps_plan_operations_dataframe(app.APS_PLAN_SHORTAGE_OPERATIONS, site)
        if plan_error:
            raise SnapshotRefreshError(f"{site}: APS PLAN API 조회 실패: {clean_log_text(plan_error)}")
        validate_plan_snapshot(site, plan_raw)
        plan_frames[site] = plan_raw

        refresh_key = app.build_api_shortage_refresh_key(app.BASE_DIR, site)
        refresh_key = (
            f"{refresh_key}:validated-set:{app.resolve_aps_snapshot_slot(plan_updated_at)}:"
            f"{plan_updated_at}:{wip_updated_at}:{datetime.now(app.DISPLAY_TZ).isoformat()}"
        )
        df, file_info_df, process_map_df = app.build_api_shortage_data_from_frames(
            plan_raw,
            refresh_key,
            str(app.BASE_DIR),
            site,
            plan_updated_at=plan_updated_at,
            inventory_df=inventory,
            inventory_source_label=source_label,
        )
        validate_shortage_snapshot(site, df, file_info_df)
        validate_site_uses_staged_sources(app, site, file_info_df, plan_updated_at, wip_updated_at)
        existing_df, _, _ = app.load_cloud_shortage_snapshot(site)
        validate_row_count_guard(f"{site} 생산부족", len(df), len(existing_df))
        shortage_results[site] = (df, file_info_df, process_map_df)
        results[site] = f"staged rows={len(df):,} updated_at={plan_updated_at}"
        logging.info("staged shortage snapshot site=%s rows=%s", site, f"{len(df):,}")
    return inventory, source_label, plan_frames, shortage_results, results


def refresh_snapshots(
    app,
    sites: list[str],
    require_remote_storage: bool,
    dry_run: bool,
    only_if_stale: bool,
) -> int:
    import snapshot_storage

    if require_remote_storage and not snapshot_storage.is_remote_snapshot_storage_enabled():
        raise SnapshotRefreshError(
            "SNAPSHOT_STORAGE_BACKEND=s3와 SNAPSHOT_S3_* 설정이 필요합니다. Public 저장소에는 생산 스냅샷을 쓰지 않습니다."
        )

    if not app.is_plan_api_configured():
        raise SnapshotRefreshError(f"{app.PLAN_API_KEY_ENV}가 설정되어 있지 않습니다.")

    storage_label = snapshot_storage.describe_snapshot_storage()
    sites = [app.normalize_shortage_snapshot_site_filter(site) for site in sites]
    logging.info("snapshot storage=%s", storage_label)
    logging.info(
        "API timeouts plan=%ss wip=%ss retry_attempts=%s",
        app.PLAN_API_TIMEOUT_SECONDS,
        app.APS_WIP_API_TIMEOUT_SECONDS,
        app.PLAN_API_RETRY_ATTEMPTS,
    )

    plan_updated_at = app.get_plan_api_updated_at()
    wip_updated_at = app.get_aps_wip_api_updated_at()
    logging.info("APS PLAN updated_at=%s APS WIP updated_at=%s", plan_updated_at, wip_updated_at)
    write_status(
        app,
        app.REFRESH_STATUS_CHECKING,
        api_updated_at=plan_updated_at,
        wip_api_updated_at=wip_updated_at,
        slot_key=app.get_latest_aps_snapshot_slot(plan_updated_at, wip_updated_at),
        storage=storage_label,
        sites=sites,
    )
    if not plan_updated_at or plan_updated_at == "-":
        raise SnapshotWaiting(
            app.REFRESH_STATUS_WAITING_FOR_PLAN,
            "APS PLAN API 기준시각을 확인하지 못했습니다. PLAN 데이터 갱신 대기 중입니다.",
            plan_updated_at,
            wip_updated_at,
        )

    raise_if_slots_not_ready(app, plan_updated_at, wip_updated_at)
    slot_key = app.resolve_aps_snapshot_slot(plan_updated_at)

    if only_if_stale and current_snapshot_set_matches(app, sites, plan_updated_at, wip_updated_at):
        manifest = app.get_published_snapshot_set_manifest()
        current_set = app.get_published_snapshot_set_pointer()
        logging.info("validated snapshot set already current set=%s", manifest.get("set_id", ""))
        write_status(
            app,
            app.REFRESH_STATUS_PUBLISHED,
            api_updated_at=plan_updated_at,
            wip_api_updated_at=wip_updated_at,
            slot_key=slot_key,
            set_id=manifest.get("set_id", ""),
            published_at=current_set.get("published_at", "") or manifest.get("created_at", ""),
            storage=storage_label,
            sites=sites,
            skipped=True,
            reason="stored validated snapshot set is already current",
        )
        return 0

    write_status(
        app,
        app.REFRESH_STATUS_READY,
        api_updated_at=plan_updated_at,
        wip_api_updated_at=wip_updated_at,
        slot_key=slot_key,
        storage=storage_label,
        sites=sites,
    )

    try:
        write_status(
            app,
            app.REFRESH_STATUS_BUILDING,
            api_updated_at=plan_updated_at,
            wip_api_updated_at=wip_updated_at,
            slot_key=slot_key,
            storage=storage_label,
            sites=sites,
        )
        inventory, source_label, plan_frames, shortage_results, results = build_validated_snapshot_set(
            app,
            sites,
            plan_updated_at,
            wip_updated_at,
        )
    except SnapshotRefreshError as exc:
        reason = clean_log_text(exc)
        waiting_status = (
            app.REFRESH_STATUS_WAITING_FOR_WIP
            if "WIP" in reason and is_transient_source_error(reason)
            else app.REFRESH_STATUS_WAITING_FOR_PLAN
            if "PLAN" in reason and is_transient_source_error(reason)
            else ""
        )
        if waiting_status:
            status = maybe_delayed_status(app, waiting_status, plan_updated_at, wip_updated_at)
            if status != app.REFRESH_STATUS_DELAYED:
                raise SnapshotWaiting(status, reason, plan_updated_at, wip_updated_at, slot_key) from exc
        raise

    write_status(
        app,
        app.REFRESH_STATUS_VALIDATING,
        api_updated_at=plan_updated_at,
        wip_api_updated_at=wip_updated_at,
        slot_key=slot_key,
        storage=storage_label,
        sites=sites,
        results=results,
    )

    if dry_run:
        logging.info("dry run complete; validated snapshot set was not published")
        write_status(
            app,
            app.REFRESH_STATUS_READY,
            api_updated_at=plan_updated_at,
            wip_api_updated_at=wip_updated_at,
            slot_key=slot_key,
            storage=storage_label,
            sites=sites,
            dry_run=True,
            results=results,
        )
        return 0

    write_status(
        app,
        app.REFRESH_STATUS_PUBLISHING,
        api_updated_at=plan_updated_at,
        wip_api_updated_at=wip_updated_at,
        slot_key=slot_key,
        storage=storage_label,
        sites=sites,
        results=results,
    )
    manifest = app.write_validated_snapshot_set(
        plan_updated_at,
        wip_updated_at,
        sites,
        plan_frames,
        inventory,
        source_label,
        shortage_results,
        update_flat_compat=True,
    )
    set_id = str(manifest.get("set_id", ""))
    published_at = datetime.now(app.DISPLAY_TZ).strftime("%Y-%m-%d %H:%M:%S")
    logging.info("validated snapshot set published set=%s slot=%s", set_id, manifest.get("slot_key", ""))
    write_status(
        app,
        app.REFRESH_STATUS_PUBLISHED,
        api_updated_at=plan_updated_at,
        wip_api_updated_at=wip_updated_at,
        slot_key=slot_key,
        set_id=set_id,
        published_at=published_at,
        storage=storage_label,
        sites=sites,
        results=results,
    )
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Refresh APS validated snapshot sets for the Streamlit dashboard.")
    parser.add_argument("--sites", default=",".join(DEFAULT_SITES), help="Comma-separated site filters.")
    parser.add_argument("--require-remote-storage", action="store_true", help="Fail unless private remote storage is used.")
    parser.add_argument("--dry-run", action="store_true", help="Validate API outputs without publishing a snapshot set.")
    parser.add_argument("--only-if-stale", action="store_true", help="Skip when the published validated set already matches APS 기준시각.")
    parser.add_argument("--validate-existing", action="store_true", help="Validate already saved snapshots without API calls.")
    args = parser.parse_args()

    configure_logging()
    app = load_app_module()
    sites = parse_sites(args.sites)

    try:
        if args.validate_existing:
            validate_existing_snapshots(app, sites)
            return 0
        return refresh_snapshots(app, sites, args.require_remote_storage, args.dry_run, args.only_if_stale)
    except SnapshotWaiting as exc:
        reason = clean_log_text(exc)
        logging.warning("snapshot refresh waiting: %s", reason)
        write_status(
            app,
            exc.status,
            reason=reason,
            waiting_for=exc.status if exc.status != app.REFRESH_STATUS_DELAYED else "",
            api_updated_at=exc.plan_updated_at,
            wip_api_updated_at=exc.wip_updated_at,
            slot_key=exc.slot_key or app.get_latest_aps_snapshot_slot(exc.plan_updated_at, exc.wip_updated_at),
            sites=[app.normalize_shortage_snapshot_site_filter(site) for site in sites],
            storage="remote" if args.require_remote_storage else "local-or-remote",
        )
        return 0
    except Exception as exc:
        reason = clean_log_text(exc)
        logging.exception("snapshot refresh failed: %s", reason)
        try:
            write_status(
                app,
                app.REFRESH_STATUS_FAILED,
                reason=reason,
                sites=[app.normalize_shortage_snapshot_site_filter(site) for site in sites],
                storage="remote" if args.require_remote_storage else "local-or-remote",
            )
        except Exception:
            pass
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
