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


class SnapshotPendingError(SnapshotRefreshError):
    pass


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


def validate_shortage_snapshot(site: str, df: pd.DataFrame, file_info_df: pd.DataFrame) -> None:
    validate_dataframe(f"{site} 생산부족", df, SHORTAGE_REQUIRED_COLUMNS)
    validate_dataframe(f"{site} 파일정보", file_info_df, FILE_INFO_REQUIRED_COLUMNS)


def ensure_wip_ready_for_plan(app, plan_updated_at: str, wip_updated_at: str) -> None:
    plan_dt = app.parse_updated_at_value(plan_updated_at)
    wip_dt = app.parse_updated_at_value(wip_updated_at)
    if plan_dt is None:
        raise SnapshotRefreshError("APS PLAN API 기준시각을 확인하지 못했습니다.")
    if wip_dt is None:
        raise SnapshotPendingError("APS WIP API 기준시각을 확인하지 못했습니다. WIP 데이터 갱신 대기 중입니다.")
    if wip_dt.timestamp() + 1 < plan_dt.timestamp():
        raise SnapshotPendingError(
            "APS WIP API 기준시각이 APS PLAN 기준시각보다 오래되었습니다. "
            f"PLAN={plan_updated_at}, WIP={wip_updated_at}. WIP 데이터 갱신 대기 중입니다."
        )


def write_status(app, status: str, **payload: object) -> None:
    import snapshot_storage

    status_payload = {
        "checked_at": datetime.now(app.DISPLAY_TZ).strftime("%Y-%m-%d %H:%M:%S"),
        "status": status,
        **payload,
    }
    try:
        snapshot_storage.write_json_snapshot_atomic(
            app.CLOUD_SNAPSHOT_DIR,
            STATUS_SNAPSHOT_NAME,
            status_payload,
        )
    except Exception as exc:
        logging.warning("status write failed: %s", clean_log_text(exc))


def load_app_module():
    os.chdir(PROJECT_ROOT)
    sys.path.insert(0, str(PROJECT_ROOT))
    os.environ.setdefault("INTEROJO_FORCE_PLAN_API", "1")
    import app  # noqa: PLC0415

    return app


def validate_existing_snapshots(app, sites: list[str]) -> None:
    inventory, _ = app.load_cloud_wip_inventory_snapshot_with_label()
    validate_wip_snapshot(app, inventory)
    logging.info("validated existing WIP snapshot rows=%s", f"{len(inventory):,}")

    for site in sites:
        df, file_info_df, _ = app.load_cloud_shortage_snapshot(site)
        validate_shortage_snapshot(site, df, file_info_df)
        logging.info("validated existing shortage snapshot site=%s rows=%s", site, f"{len(df):,}")


def snapshots_current_for_updated_at(app, sites: list[str], api_updated_at: str, wip_api_updated_at: str) -> bool:
    inventory, _ = app.load_cloud_wip_inventory_snapshot_with_label()
    if inventory.empty:
        return False
    if wip_api_updated_at != "-" and not app.is_cloud_wip_inventory_snapshot_current(wip_api_updated_at):
        return False

    for site in sites:
        if not app.is_cloud_snapshot_fresh(app.shortage_snapshot_meta_key(site), api_updated_at):
            return False
        snapshot_df, file_info_df, _ = app.load_cloud_shortage_snapshot(site)
        if snapshot_df.empty or file_info_df.empty:
            return False
    return True


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
    logging.info("snapshot storage=%s", storage_label)
    logging.info(
        "API timeouts plan=%ss wip=%ss retry_attempts=%s",
        app.PLAN_API_TIMEOUT_SECONDS,
        app.APS_WIP_API_TIMEOUT_SECONDS,
        app.PLAN_API_RETRY_ATTEMPTS,
    )

    api_updated_at = app.get_plan_api_updated_at()
    if not api_updated_at or api_updated_at == "-":
        raise SnapshotRefreshError("APS PLAN API 기준시각을 확인하지 못했습니다.")

    wip_api_updated_at = app.get_aps_wip_api_updated_at()
    logging.info("APS PLAN updated_at=%s APS WIP updated_at=%s", api_updated_at, wip_api_updated_at)
    write_status(
        app,
        "running",
        api_updated_at=api_updated_at,
        wip_api_updated_at=wip_api_updated_at,
        storage=storage_label,
        sites=sites,
    )
    ensure_wip_ready_for_plan(app, api_updated_at, wip_api_updated_at)

    if only_if_stale and snapshots_current_for_updated_at(app, sites, api_updated_at, wip_api_updated_at):
        logging.info("snapshots already current for APS PLAN updated_at=%s; skipping write", api_updated_at)
        write_status(
            app,
            "completed",
            api_updated_at=api_updated_at,
            wip_api_updated_at=wip_api_updated_at,
            storage=storage_label,
            sites=sites,
            skipped=True,
            reason="stored snapshots are already current",
        )
        return 0

    inventory, source_label, raw_snapshot_dir, wip_error = app.build_wip_inventory_snapshot_from_api(wip_api_updated_at)
    if wip_error:
        raise SnapshotRefreshError(f"APS WIP API 조회 실패: {clean_log_text(wip_error)}")
    validate_wip_snapshot(app, inventory)
    logging.info(
        "prepared WIP snapshot rows=%s source=%s raw_dir=%s",
        f"{len(inventory):,}",
        clean_log_text(source_label),
        clean_log_text(raw_snapshot_dir),
    )

    prepared_sites: list[tuple[str, pd.DataFrame, pd.DataFrame, pd.DataFrame]] = []
    for site in sites:
        refresh_key = app.build_api_shortage_refresh_key(app.BASE_DIR, site)
        refresh_key = f"{refresh_key}:github-actions:{datetime.now(app.DISPLAY_TZ).isoformat()}"
        df, file_info_df, process_map_df = app.load_api_shortage_data(refresh_key, str(app.BASE_DIR), site)
        validate_shortage_snapshot(site, df, file_info_df)
        prepared_sites.append((site, df, file_info_df, process_map_df))
        logging.info("prepared shortage snapshot site=%s rows=%s", site, f"{len(df):,}")

    if dry_run:
        logging.info("dry run complete; snapshots were not written")
        write_status(
            app,
            "completed",
            api_updated_at=api_updated_at,
            wip_api_updated_at=wip_api_updated_at,
            storage=storage_label,
            sites=sites,
            dry_run=True,
            results={"WIP": f"validated rows={len(inventory):,}"},
        )
        return 0

    if not app.write_cloud_wip_inventory_snapshot(inventory, wip_api_updated_at, source_label):
        raise SnapshotRefreshError("WIP 스냅샷 저장 실패")
    logging.info("wrote WIP snapshot path=%s", snapshot_storage.display_snapshot_uri(app.WIP_INVENTORY_SNAPSHOT_FILE))

    results: dict[str, str] = {"WIP": f"refreshed rows={len(inventory):,} updated_at={wip_api_updated_at}"}
    for site, df, file_info_df, process_map_df in prepared_sites:
        if not app.write_cloud_shortage_snapshot(df, file_info_df, process_map_df, api_updated_at, site):
            raise SnapshotRefreshError(f"{site}: 생산부족 스냅샷 저장 실패")
        snapshot_name, info_name, process_name = app.shortage_snapshot_file_names(site)
        results[site] = f"refreshed rows={len(df):,} updated_at={api_updated_at}"
        logging.info(
            "wrote shortage snapshot site=%s rows=%s paths=%s,%s,%s",
            site,
            f"{len(df):,}",
            snapshot_storage.display_snapshot_uri(snapshot_name),
            snapshot_storage.display_snapshot_uri(info_name),
            snapshot_storage.display_snapshot_uri(process_name),
        )

    write_status(
        app,
        "completed",
        api_updated_at=api_updated_at,
        wip_api_updated_at=wip_api_updated_at,
        storage=storage_label,
        sites=sites,
        results=results,
    )
    logging.info("snapshot refresh completed sites=%s", ",".join(sites))
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Refresh APS snapshots for the Streamlit dashboard.")
    parser.add_argument("--sites", default=",".join(DEFAULT_SITES), help="Comma-separated site filters.")
    parser.add_argument("--require-remote-storage", action="store_true", help="Fail unless private remote storage is used.")
    parser.add_argument("--dry-run", action="store_true", help="Validate API outputs without writing snapshots.")
    parser.add_argument("--only-if-stale", action="store_true", help="Skip when saved snapshots already match APS 기준시각.")
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
    except SnapshotPendingError as exc:
        reason = clean_log_text(exc)
        logging.warning("snapshot refresh pending: %s", reason)
        try:
            write_status(
                app,
                "pending",
                reason=reason,
                sites=sites,
                storage="remote" if args.require_remote_storage else "local-or-remote",
            )
        except Exception:
            pass
        return 1
    except Exception as exc:
        reason = clean_log_text(exc)
        logging.exception("snapshot refresh failed: %s", reason)
        try:
            write_status(
                app,
                "failed",
                reason=reason,
                sites=sites,
                storage="remote" if args.require_remote_storage else "local-or-remote",
            )
        except Exception:
            pass
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
