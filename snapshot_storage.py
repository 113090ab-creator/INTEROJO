from __future__ import annotations

import json
import mimetypes
import os
import uuid
from dataclasses import dataclass
from pathlib import Path
from typing import Callable


SNAPSHOT_STORAGE_BACKEND_ENV = "SNAPSHOT_STORAGE_BACKEND"
SNAPSHOT_S3_BUCKET_ENV = "SNAPSHOT_S3_BUCKET"
SNAPSHOT_S3_PREFIX_ENV = "SNAPSHOT_S3_PREFIX"
SNAPSHOT_S3_REGION_ENV = "SNAPSHOT_S3_REGION"
SNAPSHOT_S3_ENDPOINT_URL_ENV = "SNAPSHOT_S3_ENDPOINT_URL"
SNAPSHOT_S3_ACCESS_KEY_ID_ENV = "SNAPSHOT_S3_ACCESS_KEY_ID"
SNAPSHOT_S3_SECRET_ACCESS_KEY_ENV = "SNAPSHOT_S3_SECRET_ACCESS_KEY"
SNAPSHOT_STORAGE_TIMEOUT_SECONDS_ENV = "SNAPSHOT_STORAGE_TIMEOUT_SECONDS"

_secret_resolver: Callable[[str, str], str] | None = None


@dataclass(frozen=True)
class S3SnapshotConfig:
    bucket: str
    prefix: str
    region: str
    endpoint_url: str
    access_key_id: str
    secret_access_key: str
    timeout_seconds: int


def configure_secret_resolver(resolver: Callable[[str, str], str]) -> None:
    global _secret_resolver
    _secret_resolver = resolver


def get_setting(name: str, default: str = "") -> str:
    value = ""
    if _secret_resolver is not None:
        try:
            value = _secret_resolver(name, "")
        except Exception:
            value = ""
    if value is None or str(value).strip() == "":
        value = os.environ.get(name, default)
    return str(value or "").strip()


def get_snapshot_storage_backend() -> str:
    backend = get_setting(SNAPSHOT_STORAGE_BACKEND_ENV, "").strip().lower()
    if backend:
        return backend
    if get_setting(SNAPSHOT_S3_BUCKET_ENV, ""):
        return "s3"
    return "local"


def is_remote_snapshot_storage_enabled() -> bool:
    return get_snapshot_storage_backend() in {"s3", "r2"}


def normalize_snapshot_name(name: str) -> str:
    clean_name = str(name or "").replace("\\", "/").lstrip("/")
    if not clean_name or clean_name.startswith("../") or "/../" in clean_name or clean_name == "..":
        raise ValueError(f"invalid snapshot name: {name!r}")
    return clean_name


def normalize_s3_prefix(prefix: str) -> str:
    return str(prefix or "").strip().strip("/")


def get_s3_config() -> S3SnapshotConfig:
    bucket = get_setting(SNAPSHOT_S3_BUCKET_ENV, "")
    if not bucket:
        raise RuntimeError(f"{SNAPSHOT_S3_BUCKET_ENV} is not configured.")

    access_key_id = get_setting(SNAPSHOT_S3_ACCESS_KEY_ID_ENV, "") or os.environ.get("AWS_ACCESS_KEY_ID", "")
    secret_access_key = get_setting(SNAPSHOT_S3_SECRET_ACCESS_KEY_ENV, "") or os.environ.get(
        "AWS_SECRET_ACCESS_KEY",
        "",
    )
    if not access_key_id or not secret_access_key:
        raise RuntimeError(
            f"{SNAPSHOT_S3_ACCESS_KEY_ID_ENV}/{SNAPSHOT_S3_SECRET_ACCESS_KEY_ENV} are not configured."
        )

    try:
        timeout_seconds = int(get_setting(SNAPSHOT_STORAGE_TIMEOUT_SECONDS_ENV, "30"))
    except ValueError:
        timeout_seconds = 30

    return S3SnapshotConfig(
        bucket=bucket,
        prefix=normalize_s3_prefix(get_setting(SNAPSHOT_S3_PREFIX_ENV, "interojo/snapshots")),
        region=get_setting(SNAPSHOT_S3_REGION_ENV, "ap-northeast-2"),
        endpoint_url=get_setting(SNAPSHOT_S3_ENDPOINT_URL_ENV, ""),
        access_key_id=access_key_id,
        secret_access_key=secret_access_key,
        timeout_seconds=max(timeout_seconds, 5),
    )


def build_s3_key(config: S3SnapshotConfig, name: str) -> str:
    snapshot_name = normalize_snapshot_name(name)
    return f"{config.prefix}/{snapshot_name}" if config.prefix else snapshot_name


def get_s3_client(config: S3SnapshotConfig | None = None):
    config = config or get_s3_config()
    try:
        import boto3
        from botocore.config import Config
    except ImportError as exc:
        raise RuntimeError("boto3 is required when SNAPSHOT_STORAGE_BACKEND=s3.") from exc

    return boto3.client(
        "s3",
        endpoint_url=config.endpoint_url or None,
        region_name=config.region,
        aws_access_key_id=config.access_key_id,
        aws_secret_access_key=config.secret_access_key,
        config=Config(
            connect_timeout=config.timeout_seconds,
            read_timeout=config.timeout_seconds,
            retries={"max_attempts": 3, "mode": "standard"},
        ),
    )


def snapshot_path(snapshot_dir: Path, name: str) -> Path:
    return snapshot_dir / normalize_snapshot_name(name)


def snapshot_exists(snapshot_dir: Path, name: str) -> bool:
    if is_remote_snapshot_storage_enabled():
        try:
            config = get_s3_config()
            client = get_s3_client(config)
            client.head_object(Bucket=config.bucket, Key=build_s3_key(config, name))
        except Exception:
            return False
        return True
    return snapshot_path(snapshot_dir, name).exists()


def snapshot_signature(snapshot_dir: Path, name: str) -> str:
    if is_remote_snapshot_storage_enabled():
        try:
            config = get_s3_config()
            client = get_s3_client(config)
            head = client.head_object(Bucket=config.bucket, Key=build_s3_key(config, name))
        except Exception:
            return f"{name}:missing"
        etag = str(head.get("ETag", "")).strip('"')
        last_modified = head.get("LastModified", "")
        size = int(head.get("ContentLength", 0) or 0)
        return f"s3:{name}:{size}:{last_modified}:{etag}"

    path = snapshot_path(snapshot_dir, name)
    try:
        stat = path.stat()
    except OSError:
        return f"{name}:missing"
    return f"local:{name}:{stat.st_size}:{stat.st_mtime_ns}"


def read_snapshot_bytes(snapshot_dir: Path, name: str) -> bytes:
    if is_remote_snapshot_storage_enabled():
        config = get_s3_config()
        client = get_s3_client(config)
        obj = client.get_object(Bucket=config.bucket, Key=build_s3_key(config, name))
        return obj["Body"].read()
    return snapshot_path(snapshot_dir, name).read_bytes()


def guess_content_type(name: str) -> str:
    if name.endswith(".json"):
        return "application/json; charset=utf-8"
    if name.endswith(".csv") or name.endswith(".csv.gz"):
        return "text/csv; charset=utf-8"
    guessed, _ = mimetypes.guess_type(name)
    return guessed or "application/octet-stream"


def write_snapshot_bytes_atomic(snapshot_dir: Path, name: str, data: bytes) -> None:
    snapshot_name = normalize_snapshot_name(name)
    if is_remote_snapshot_storage_enabled():
        config = get_s3_config()
        client = get_s3_client(config)
        final_key = build_s3_key(config, snapshot_name)
        temp_key = f"{final_key}.tmp.{uuid.uuid4().hex}"
        client.put_object(
            Bucket=config.bucket,
            Key=temp_key,
            Body=data,
            ContentType=guess_content_type(snapshot_name),
            Metadata={"interojo-snapshot-temp": "true"},
        )
        try:
            client.copy_object(
                Bucket=config.bucket,
                Key=final_key,
                CopySource={"Bucket": config.bucket, "Key": temp_key},
                ContentType=guess_content_type(snapshot_name),
                MetadataDirective="REPLACE",
                Metadata={"interojo-snapshot": "true"},
            )
        finally:
            try:
                client.delete_object(Bucket=config.bucket, Key=temp_key)
            except Exception:
                pass
        return

    path = snapshot_path(snapshot_dir, snapshot_name)
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = path.with_suffix(path.suffix + f".{uuid.uuid4().hex}.tmp")
    try:
        temp_path.write_bytes(data)
        temp_path.replace(path)
    except Exception:
        try:
            temp_path.unlink(missing_ok=True)
        except Exception:
            pass
        raise


def read_json_snapshot(snapshot_dir: Path, name: str) -> dict[str, object]:
    try:
        payload = json.loads(read_snapshot_bytes(snapshot_dir, name).decode("utf-8-sig"))
    except Exception:
        return {}
    return payload if isinstance(payload, dict) else {}


def write_json_snapshot_atomic(snapshot_dir: Path, name: str, payload: dict[str, object]) -> None:
    data = json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")
    write_snapshot_bytes_atomic(snapshot_dir, name, data)


def display_snapshot_uri(name: str) -> str:
    snapshot_name = normalize_snapshot_name(name)
    if is_remote_snapshot_storage_enabled():
        config = get_s3_config()
        key = build_s3_key(config, snapshot_name)
        return f"s3://{config.bucket}/{key}"
    return str(snapshot_name)


def describe_snapshot_storage() -> str:
    if is_remote_snapshot_storage_enabled():
        config = get_s3_config()
        prefix = f"/{config.prefix}" if config.prefix else ""
        endpoint = f" ({config.endpoint_url})" if config.endpoint_url else ""
        return f"S3 비공개 저장소: s3://{config.bucket}{prefix}{endpoint}"
    return "로컬 cloud_snapshots 폴더"
