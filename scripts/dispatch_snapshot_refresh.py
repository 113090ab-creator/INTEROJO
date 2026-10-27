"""Start the official GitHub refresh from the Windows scheduler."""

from __future__ import annotations

import argparse
import base64
import json
import logging
import os
import subprocess
import sys
import time
import urllib.error
import urllib.request
from datetime import datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo


ROOT = Path(__file__).resolve().parents[1]
REPOSITORY = "113090ab-creator/INTEROJO"
WORKFLOW = "refresh_snapshot.yml"
KST = ZoneInfo("Asia/Seoul")


def target_slot(now: datetime) -> str:
    now = now.astimezone(KST)
    if now.hour < 8:
        return f"{now.date() - timedelta(days=1):%Y-%m-%d} PM"
    return f"{now:%Y-%m-%d} {'AM' if now.hour < 16 else 'PM'}"


def slot_order(slot: str) -> tuple[str, int]:
    date, period = slot.split()
    datetime.strptime(date, "%Y-%m-%d")
    if period not in {"AM", "PM"}:
        raise ValueError("Invalid snapshot period")
    return date, int(period == "PM")


def github_token() -> str:
    token = os.environ.get("GH_TOKEN") or os.environ.get("GITHUB_TOKEN")
    if token:
        return token
    result = subprocess.run(
        ["git", "credential", "fill"],
        input=f"protocol=https\nhost=github.com\npath={REPOSITORY}.git\n\n",
        capture_output=True, text=True, cwd=ROOT, timeout=30,
        env=dict(os.environ, GIT_TERMINAL_PROMPT="0", GCM_INTERACTIVE="Never"),
    )
    fields = dict(line.split("=", 1) for line in result.stdout.splitlines() if "=" in line)
    if result.returncode or not fields.get("password"):
        raise RuntimeError("GitHub login unavailable; sign in to this repository with GitHub Desktop.")
    return fields["password"]


class GitHub:
    def __init__(self, token: str):
        self.token = token

    def request(self, path: str, payload: dict | None = None):
        request = urllib.request.Request(
            f"https://api.github.com/repos/{REPOSITORY}/{path}",
            data=json.dumps(payload).encode() if payload is not None else None,
            headers={"Accept": "application/vnd.github+json",
                     "Authorization": f"Bearer {self.token}",
                     "User-Agent": "INTEROJO-external-scheduler",
                     "Content-Type": "application/json"},
        )
        try:
            with urllib.request.urlopen(request, timeout=30) as response:
                body = response.read()
                return json.loads(body) if body else {}
        except urllib.error.HTTPError as error:
            raise RuntimeError(f"GitHub API HTTP {error.code} ({path.split('?')[0]})") from None
        except urllib.error.URLError:
            raise RuntimeError("GitHub network request failed") from None

    def pointer(self) -> dict:
        result = self.request("contents/cloud_snapshots/current_snapshot_set.json?ref=main")
        return json.loads(base64.b64decode(result["content"]))


def run(github: GitHub, check_only: bool = False, wait_seconds: int = 6600) -> int:
    target = target_slot(datetime.now(KST))
    workflow = github.request(f"actions/workflows/{WORKFLOW}")
    if workflow.get("state") != "active":
        raise RuntimeError("Official snapshot refresh workflow is not active")
    pointer = github.pointer()
    logging.info("Target=%s published=%s set=%s", target, pointer.get("slot_key"), pointer.get("set_id"))
    if check_only:
        logging.info("CHECK OK: GitHub login and official workflow are available; no dispatch requested")
        return 0
    if slot_order(pointer["slot_key"]) >= slot_order(target):
        logging.info("ALREADY PUBLISHED: target snapshot is current")
        return 0

    recent = github.request(f"actions/workflows/{WORKFLOW}/runs?per_page=50")["workflow_runs"]
    active = [run for run in recent if run["status"] != "completed" and run["head_branch"] == "main"]
    if active:
        # The workflow's existing concurrency group also serializes backup refreshes.
        run_id = min(active, key=lambda item: item["id"])["id"]
        logging.info("Watching existing official refresh: https://github.com/%s/actions/runs/%s", REPOSITORY, run_id)
    else:
        previous_max_id = max((run["id"] for run in recent), default=0)
        github.request(f"actions/workflows/{WORKFLOW}/dispatches", {"ref": "main", "inputs": {"sites": "C관,A관,S관,전체"}})
        logging.info("DISPATCHED: official main workflow; waiting for GitHub run")
        run_id = None
        for _ in range(24):
            time.sleep(5)
            runs = github.request(f"actions/workflows/{WORKFLOW}/runs?event=workflow_dispatch&per_page=20")["workflow_runs"]
            new = [run for run in runs if run["id"] > previous_max_id and run["head_branch"] == "main"]
            if new:
                run_id = max(new, key=lambda item: item["id"])["id"]
                break
        if run_id is None:
            raise RuntimeError("GitHub accepted dispatch but no new run appeared within two minutes")
        logging.info("RUN: https://github.com/%s/actions/runs/%s", REPOSITORY, run_id)

    deadline = time.monotonic() + wait_seconds
    while time.monotonic() < deadline:
        result = github.request(f"actions/runs/{run_id}")
        if result["status"] == "completed":
            pointer = github.pointer()
            if result.get("conclusion") != "success":
                raise RuntimeError(f"Official refresh or live verification failed: run={run_id} conclusion={result.get('conclusion')}")
            if slot_order(pointer["slot_key"]) < slot_order(target):
                raise RuntimeError(f"Workflow completed without publishing target {target}")
            logging.info("SUCCESS: published=%s set=%s run=%s", pointer["slot_key"], pointer["set_id"], run_id)
            return 0
        time.sleep(30)
    raise RuntimeError(f"Official refresh did not complete within {wait_seconds} seconds; run={run_id}")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true", help="Check GitHub access without starting a refresh")
    args = parser.parse_args()
    log_dir = ROOT / "outputs" / "snapshot_scheduler"
    log_dir.mkdir(parents=True, exist_ok=True)
    handlers = [logging.FileHandler(log_dir / f"dispatch_{datetime.now(KST):%Y%m%d}.log", encoding="utf-8")]
    if sys.stdout is not None:
        handlers.append(logging.StreamHandler(sys.stdout))
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s", handlers=handlers)
    try:
        return run(GitHub(github_token()), check_only=args.check)
    except Exception as error:
        # Avoid printing request objects, credential helper output, or access tokens.
        logging.error("%s", error if isinstance(error, RuntimeError) else type(error).__name__)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
