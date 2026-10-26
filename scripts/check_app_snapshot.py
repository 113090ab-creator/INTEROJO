"""Check the snapshot actually rendered by the public Streamlit app."""

import argparse
import asyncio
import json
import os
import re
import sys
import time
import urllib.request
from datetime import datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo


ROOT = Path(__file__).resolve().parents[1]
KST = ZoneInfo("Asia/Seoul")


def slot_order(slot: str) -> tuple:
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2} (AM|PM)", slot):
        raise ValueError("invalid snapshot slot")
    date, period = slot.split()
    return date, int(period == "PM")


def required_slot(now: datetime) -> str:
    now = now.astimezone(KST)
    minute = now.hour * 60 + now.minute
    if minute > 17 * 60 + 55:
        return f"{now:%Y-%m-%d} PM"
    if minute > 9 * 60 + 55:
        return f"{now:%Y-%m-%d} AM"
    return f"{now.date() - timedelta(days=1):%Y-%m-%d} PM"


def parse_rendered_snapshot(text: str) -> dict:
    match = re.search(r"\bSet:\s*([A-Za-z0-9_-]+)", text)
    if not match:
        return {}
    status = re.search(r"상태\s*:\s*(최신|갱신 중|갱신 지연|갱신 실패)", text[match.end():])
    return {"set_id": match.group(1), "display_status": status.group(1) if status else ""}


def check_rendered_snapshot(expected: dict, observed: dict, now: datetime, require_sync: bool = False) -> tuple[bool, str]:
    if slot_order(expected["slot_key"]) < slot_order(required_slot(now)):
        return False, "published snapshot missed its operational deadline"
    if observed.get("set_id") != expected["set_id"]:
        return False, "app has not rendered the published snapshot"
    if require_sync and observed.get("sync_revision") != "validated-mirror-v1":
        return False, "app has not deployed the independent snapshot recovery code"
    minute = now.astimezone(KST).hour * 60 + now.astimezone(KST).minute
    target = f"{now.astimezone(KST):%Y-%m-%d} {'PM' if minute >= 16 * 60 else 'AM'}"
    if minute < 8 * 60:
        target = required_slot(now)
    if slot_order(expected["slot_key"]) >= slot_order(target) and observed.get("display_status") != "최신":
        return False, "app snapshot matches but its operational status is not current"
    return True, "published snapshot is rendered by the app"


def refresh_expected_pointer(expected: dict) -> None:
    repository = os.environ.get("GITHUB_REPOSITORY", "113090ab-creator/INTEROJO")
    url = f"https://raw.githubusercontent.com/{repository}/main/cloud_snapshots/current_snapshot_set.json?check={int(time.time())}"
    try:
        with urllib.request.urlopen(url, timeout=10) as response:
            current = json.load(response)
        def order(pointer):
            return (slot_order(pointer["slot_key"]), *(str(pointer.get(key, "")) for key in
                    ("plan_updated_at", "wip_updated_at", "published_at")))
        if current.get("set_id") and order(current) >= order(expected):
            expected.update(current)
    except Exception:
        pass  # The checked-out published pointer is still an authoritative lower bound.


async def verify(url: str, expected: dict, timeout: int) -> dict:
    from playwright.async_api import async_playwright
    observed = {}
    reason = "app has not rendered the published snapshot"
    deadline = time.monotonic() + timeout
    async with async_playwright() as playwright:
        browser = await playwright.chromium.launch(headless=True)
        page = await browser.new_page()
        try:
            while time.monotonic() < deadline:
                try:
                    await page.goto(url, wait_until="domcontentloaded", timeout=45000)
                    wait_until = min(deadline, time.monotonic() + 35)
                    while time.monotonic() < wait_until:
                        for frame in page.frames:
                            text = await frame.locator("body").inner_text(timeout=5000)
                            current = parse_rendered_snapshot(text)
                            if current:
                                observed = current
                                marker = frame.locator('[data-testid="snapshot-health"]')
                                if await marker.count():
                                    observed["sync_revision"] = await marker.get_attribute("data-sync-revision")
                                if observed.get("set_id") != expected["set_id"]:
                                    await asyncio.to_thread(refresh_expected_pointer, expected)
                                ok, reason = check_rendered_snapshot(expected, observed, datetime.now(KST), require_sync=True)
                                if ok:
                                    return {"ok": True, "reason": reason, "observed": observed}
                        await asyncio.sleep(3)
                except Exception as exc:
                    reason = f"app probe could not read the rendered snapshot ({type(exc).__name__})"
                await asyncio.sleep(min(15, max(0, deadline - time.monotonic())))
        finally:
            await browser.close()
    return {"ok": False, "reason": reason, "observed": observed}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--expected", type=Path, default=ROOT / "cloud_snapshots/current_snapshot_set.json")
    parser.add_argument("--url", default=os.environ.get("STREAMLIT_URL", "https://interojo-2.streamlit.app/"))
    parser.add_argument("--timeout", type=int, default=480)
    parser.add_argument("--report", type=Path)
    args = parser.parse_args()
    expected = json.loads(args.expected.read_text(encoding="utf-8-sig"))
    now = datetime.now(KST)
    if slot_order(expected["slot_key"]) < slot_order(required_slot(now)):
        result = {"ok": False, "reason": "published snapshot missed its operational deadline", "observed": {}}
    else:
        result = asyncio.run(verify(args.url, expected, args.timeout))
    result.update({"expected_set_id": expected["set_id"], "expected_slot": expected["slot_key"],
                   "required_slot": required_slot(now), "checked_at": datetime.now(KST).isoformat()})
    report = json.dumps(result, ensure_ascii=False, indent=2)
    print(report)
    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(report, encoding="utf-8")
    summary = os.environ.get("GITHUB_STEP_SUMMARY")
    if summary:
        with open(summary, "a", encoding="utf-8") as file:
            file.write(f"### Live snapshot verification: {'PASS' if result['ok'] else 'FAIL'}\n\n```json\n{report}\n```\n")
    return 0 if result["ok"] else 1


if __name__ == "__main__":
    sys.exit(main())
