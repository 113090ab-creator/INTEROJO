import importlib.util
import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]


class SnapshotStorageImportTests(unittest.TestCase):
    def test_snapshot_storage_loads_without_sys_modules_entry(self) -> None:
        spec = importlib.util.spec_from_file_location(
            "isolated_snapshot_storage_import",
            PROJECT_ROOT / "snapshot_storage.py",
        )
        self.assertIsNotNone(spec)
        self.assertIsNotNone(spec.loader)
        module = importlib.util.module_from_spec(spec)

        spec.loader.exec_module(module)

        config = module.GitHubSnapshotConfig(
            repository="owner/repo",
            branch="main",
            prefix="cloud_snapshots",
            token="",
            timeout_seconds=15,
            cache_ttl_seconds=15,
        )
        self.assertEqual(config.repository, "owner/repo")


if __name__ == "__main__":
    unittest.main()
