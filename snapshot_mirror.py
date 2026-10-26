"""Refresh published APS data independently of Streamlit's Git checkout."""

import gzip
import hashlib
import json
import logging
import re
import shutil
import threading
import time
import urllib.parse
import urllib.request
import uuid
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path, PurePosixPath
from typing import Callable


POINTER = "current_snapshot_set.json"
STATUS = "aps_snapshot_refresh_status.json"
META = "snapshot_meta.csv"
REFRESH_SECONDS = 300
_mirrors = {}
_mirrors_lock = threading.Lock()


def read_json(path: Path) -> dict:
    try:
        value = json.loads(path.read_bytes())
        return value if isinstance(value, dict) else {}
    except (OSError, ValueError):
        return {}


def snapshot_order(pointer: dict) -> tuple:
    return tuple(str(pointer.get(key, "")) for key in (
        "slot_key", "plan_updated_at", "wip_updated_at", "published_at"
    ))


def validate_pointer(pointer: dict) -> str:
    set_id = str(pointer.get("set_id", ""))
    if not re.fullmatch(r"[A-Za-z0-9_-]+", set_id):
        raise ValueError("invalid published snapshot set id")
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2} (AM|PM)", str(pointer.get("slot_key", ""))):
        raise ValueError("invalid published snapshot slot")
    if pointer.get("manifest_path") != f"sets/{set_id}/manifest.json":
        raise ValueError("invalid published manifest path")
    for key in ("plan_updated_at", "wip_updated_at", "published_at"):
        if not pointer.get(key):
            raise ValueError(f"published snapshot is missing {key}")
    return set_id


def validate_manifest(pointer: dict, manifest: dict) -> None:
    set_id = validate_pointer(pointer)
    for key in ("set_id", "slot_key", "plan_updated_at", "wip_updated_at"):
        if manifest.get(key) != pointer.get(key):
            raise ValueError(f"published pointer and manifest disagree: {key}")
    files = manifest.get("files")
    if not isinstance(files, dict) or not files:
        raise ValueError("published manifest has no files")
    for name, entry in files.items():
        if not isinstance(entry, dict) or name in (".", "..") or PurePosixPath(name).name != name or "\\" in name:
            raise ValueError("invalid published filename")
        if entry.get("path") != f"sets/{set_id}/{name}":
            raise ValueError("invalid published data path")
        if not re.fullmatch(r"[a-f0-9]{64}", str(entry.get("checksum", ""))):
            raise ValueError("invalid published data checksum")


def content_matches(data: bytes, name: str, checksum: str) -> bool:
    try:
        content = gzip.decompress(data) if name.endswith(".gz") else data
        return hashlib.sha256(content).hexdigest() == checksum
    except (OSError, EOFError):
        return False


def atomic_write(path: Path, data: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(f".{path.name}.{uuid.uuid4().hex}.tmp")
    try:
        temp.write_bytes(data)
        temp.replace(path)
    finally:
        temp.unlink(missing_ok=True)


class SnapshotMirror:
    def __init__(self, source: Path, cache: Path, repository: str, branch: str = "main",
                 fetch: Callable[[str], bytes] | None = None):
        self.source, self.cache = Path(source), Path(cache)
        self.repository, self.branch = repository, branch
        self.fetch = fetch or self.fetch_public_file
        self.lock = threading.Lock()
        self.last_check = float("-inf")
        self.last_error = ""

    def fetch_public_file(self, name: str) -> bytes:
        path = urllib.parse.quote(f"cloud_snapshots/{name}", safe="/")
        branch = urllib.parse.quote(self.branch, safe="")
        url = f"https://raw.githubusercontent.com/{self.repository}/{branch}/{path}"
        if not name.startswith("sets/"):
            url += f"?check={int(time.time() // REFRESH_SECONDS)}"
        request = urllib.request.Request(url, headers={
            "User-Agent": "INTEROJO-published-snapshot-mirror", "Cache-Control": "no-cache"
        })
        with urllib.request.urlopen(request, timeout=10) as response:
            return response.read()

    def bundle(self) -> Path | None:
        active = read_json(self.cache / "active.json")
        version = str(active.get("version", ""))
        if not re.fullmatch(r"[a-f0-9]{64}", version):
            return None
        directory = self.cache / "bundles" / version
        pointer = read_json(directory / POINTER)
        local = read_json(self.source / POINTER)
        if not pointer or snapshot_order(pointer) < snapshot_order(local):
            return None
        if snapshot_order(pointer) == snapshot_order(local):
            mirror_checked = str(read_json(directory / STATUS).get("checked_at", ""))
            local_checked = str(read_json(self.source / STATUS).get("checked_at", ""))
            if mirror_checked < local_checked:
                return None
        return directory

    def read_path(self, name: str) -> Path:
        local = self.source / name
        if name.startswith("sets/") and ".." not in PurePosixPath(name).parts and "\\" not in name:
            cached = self.cache / name
            return cached if cached.is_file() else local
        bundle = self.bundle()
        if bundle is None:
            return local
        if name in (POINTER, STATUS, META):
            candidate = bundle / name
            return candidate if candidate.is_file() else local
        pointer = read_json(bundle / POINTER)
        manifest = read_json(self.cache / pointer.get("manifest_path", ""))
        entry = manifest.get("files", {}).get(name, {})
        if isinstance(entry, dict) and entry.get("path"):
            candidate = self.cache / entry["path"]
            if candidate.is_file():
                return candidate
        return local

    def version(self) -> str:
        bundle = self.bundle()
        if bundle is not None:
            return bundle.name
        pointer = read_json(self.source / POINTER)
        status = read_json(self.source / STATUS)
        return f"local:{pointer.get('set_id', '')}:{status.get('checked_at', '')}"

    def refresh(self, force: bool = False) -> str:
        with self.lock:
            if not force and time.monotonic() - self.last_check < REFRESH_SECONDS:
                return self.version()
            self.last_check = time.monotonic()
            try:
                self._refresh()
                self.last_error = ""
            except Exception as exc:
                self.last_error = type(exc).__name__
                logging.warning("Published snapshot mirror check failed (%s); retaining last validated data", self.last_error)
            return self.version()

    def _refresh(self) -> None:
        pointer_bytes = self.fetch(POINTER)
        pointer = json.loads(pointer_bytes)
        set_id = validate_pointer(pointer)
        active = self.bundle()
        current = read_json((active or self.source) / POINTER)
        if snapshot_order(pointer) < snapshot_order(current):
            return
        manifest_bytes = self.fetch(pointer["manifest_path"])
        manifest = json.loads(manifest_bytes)
        validate_manifest(pointer, manifest)
        destination = self.cache / "sets" / set_id
        if destination.exists():
            if read_json(destination / "manifest.json").get("files") != manifest["files"]:
                raise ValueError("published immutable set changed contents")
        else:
            stage = self.cache / "staging" / uuid.uuid4().hex
            stage.mkdir(parents=True, exist_ok=True)
            try:
                def copy_file(item):
                    name, entry = item
                    local_file = self.source / entry["path"]
                    data = local_file.read_bytes() if local_file.is_file() else b""
                    if not content_matches(data, name, entry["checksum"]):
                        data = self.fetch(entry["path"])
                    if not content_matches(data, name, entry["checksum"]):
                        raise ValueError(f"published file content check failed: {name}")
                    (stage / name).write_bytes(data)
                with ThreadPoolExecutor(max_workers=4) as pool:
                    list(pool.map(copy_file, manifest["files"].items()))
                (stage / "manifest.json").write_bytes(manifest_bytes)
                destination.parent.mkdir(parents=True, exist_ok=True)
                stage.replace(destination)
            finally:
                if stage.exists():
                    # Only this newly created staging directory can be removed.
                    if stage.resolve().parent == (self.cache / "staging").resolve():
                        shutil.rmtree(stage)
        status_bytes = self.fetch(STATUS)
        status = json.loads(status_bytes)
        if not isinstance(status, dict):
            raise ValueError("invalid published status")
        meta_bytes = self.fetch(META)
        version = hashlib.sha256(pointer_bytes + status_bytes + meta_bytes).hexdigest()
        if read_json(self.cache / "active.json").get("version") == version:
            return
        bundle = self.cache / "bundles" / version
        atomic_write(bundle / POINTER, pointer_bytes)
        atomic_write(bundle / STATUS, status_bytes)
        atomic_write(bundle / META, meta_bytes)
        atomic_write(self.cache / "active.json", json.dumps({"version": version}).encode())
        logging.info("Published snapshot mirror ready: set=%s slot=%s", set_id, pointer["slot_key"])
        self.prune_cache(set_id, version)

    def prune_cache(self, active_set: str, active_bundle: str) -> None:
        protected_set = str(read_json(self.source / POINTER).get("set_id", ""))
        for name, keep, protected in (("sets", 4, {active_set, protected_set}),
                                      ("bundles", 32, {active_bundle})):
            root = self.cache / name
            if not root.is_dir():
                continue
            directories = sorted((p for p in root.iterdir() if p.is_dir()),
                                 key=lambda p: p.stat().st_mtime, reverse=True)
            for directory in directories[keep:]:
                if directory.name in protected or time.time() - directory.stat().st_mtime < 86400:
                    continue
                # Remove only old, generated cache directories within this exact cache root.
                if not directory.is_symlink() and directory.resolve().parent == root.resolve():
                    shutil.rmtree(directory, ignore_errors=True)


def get_mirror(source: Path, cache: Path, repository: str, branch: str = "main") -> SnapshotMirror:
    key = (str(Path(source).resolve()), str(Path(cache).resolve()), repository, branch)
    with _mirrors_lock:
        if key not in _mirrors:
            _mirrors[key] = SnapshotMirror(source, cache, repository, branch)
        return _mirrors[key]
