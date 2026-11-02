"""Connect to local common sources when they are available on this machine."""

import json
import sys
from pathlib import Path


def _load_common_client():
    app_dir = Path(__file__).resolve().parent
    config_path = next(
        (
            directory / "공통데이터_연결.json"
            for directory in (app_dir, *app_dir.parents)
            if (directory / "공통데이터_연결.json").is_file()
        ),
        None,
    )
    if config_path is None:
        return None

    config = json.loads(config_path.read_text(encoding="utf-8-sig"))
    if not config.get("enabled"):
        return None

    code_dir = Path(config["root"]) / "05_공통계산"
    # Streamlit Cloud does not have the PC's shared source folder. The app
    # already supports COMMON=None and reads its deployed snapshots instead.
    if not (code_dir / "common_sources.py").is_file():
        return None

    code_path = str(code_dir)
    if code_path not in sys.path:
        sys.path.insert(0, code_path)

    import common_sources

    return common_sources.client(config["project_id"])


COMMON = _load_common_client()
