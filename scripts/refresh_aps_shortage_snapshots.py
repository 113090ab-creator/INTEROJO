from __future__ import annotations

import argparse
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
OFFICIAL_REFRESH_SCRIPT = PROJECT_ROOT / "scripts" / "refresh_snapshot.py"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Retired APS flat snapshot refresher. Production APS snapshots are now "
            "published only through the Validated Snapshot Set refresh path."
        )
    )
    parser.add_argument("--sites", help="Ignored legacy option.", default="")
    parser.add_argument("--only-if-stale", action="store_true", help="Ignored legacy option.")
    parser.add_argument("--scheduled", action="store_true", help="Ignored legacy option.")
    parser.add_argument("--slot-times", help="Ignored legacy option.", default="")
    parser.add_argument("--slot-grace-minutes", help="Ignored legacy option.", default="")
    parser.add_argument("--slot-lookahead-minutes", help="Ignored legacy option.", default="")
    parser.add_argument("--wip-only", action="store_true", help="Ignored legacy option.")
    parser.parse_known_args(argv)

    print(
        "\n".join(
            [
                "This legacy APS flat snapshot refresher is retired.",
                "Production cloud_snapshots writes are blocked to preserve the Single Writer rule.",
                f"Use the official VSS refresh path instead: python {OFFICIAL_REFRESH_SCRIPT}",
            ]
        ),
        file=sys.stderr,
    )
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
