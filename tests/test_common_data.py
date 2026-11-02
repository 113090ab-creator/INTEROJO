import json
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]


class CommonDataTests(unittest.TestCase):
    def setUp(self):
        self.workspace = tempfile.TemporaryDirectory()
        self.addCleanup(self.workspace.cleanup)
        self.root = Path(self.workspace.name)
        self.app_dir = self.root / "app"
        self.app_dir.mkdir()
        shutil.copy2(PROJECT_ROOT / "_common_data.py", self.app_dir)

    def write_config(self, *, enabled=True, root=None, directory=None):
        config = {
            "enabled": enabled,
            "root": str(root or self.root / "unavailable_common_root"),
            "project_id": "interojo",
        }
        (directory or self.app_dir).joinpath("공통데이터_연결.json").write_text(
            json.dumps(config), encoding="utf-8-sig"
        )

    def run_import(self, code):
        result = subprocess.run(
            [
                sys.executable,
                "-I",
                "-B",
                "-c",
                "import sys; sys.path.insert(0, sys.argv[1]); " + code,
                str(self.app_dir),
            ],
            cwd=self.root,
            capture_output=True,
            text=True,
            timeout=30,
        )
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def assert_disconnected(self):
        self.run_import(
            "before = list(sys.path); "
            "from _common_data import COMMON; "
            "assert COMMON is None; "
            "assert 'common_sources' not in sys.modules; "
            "assert sys.path == before"
        )

    def test_missing_config_allows_import(self):
        self.assert_disconnected()

    def test_disabled_connection_does_not_import_local_sources(self):
        self.write_config(enabled=False)
        self.assert_disconnected()

    def test_unavailable_local_folder_allows_cloud_import(self):
        self.write_config()
        self.assert_disconnected()

    def test_windows_config_allows_import_on_other_machines(self):
        self.write_config(root=r"C:\unavailable-workstation\15. 통합관리")
        self.assert_disconnected()

    def test_existing_folder_without_module_allows_import(self):
        (self.root / "05_공통계산").mkdir()
        self.write_config(root=self.root)
        self.assert_disconnected()

    def test_available_local_sources_keep_working_with_parent_config(self):
        code_dir = self.root / "05_공통계산"
        code_dir.mkdir()
        (code_dir / "common_sources.py").write_text(
            "def client(project_id):\n    return {'project_id': project_id}\n",
            encoding="utf-8",
        )
        self.write_config(root=self.root, directory=self.root)
        self.run_import(
            "from _common_data import COMMON; "
            "assert COMMON == {'project_id': 'interojo'}"
        )

    def test_installed_module_dependency_errors_are_not_hidden(self):
        code_dir = self.root / "05_공통계산"
        code_dir.mkdir()
        (code_dir / "common_sources.py").write_text(
            "import missing_common_source_dependency\n", encoding="utf-8"
        )
        self.write_config(root=self.root)
        self.run_import(
            "\ntry:\n    import _common_data\n"
            "except ModuleNotFoundError as exc:\n"
            "    assert exc.name == 'missing_common_source_dependency'\n"
            "else:\n    raise AssertionError('dependency error was hidden')\n"
        )

    def test_snapshot_storage_imports_without_local_sources(self):
        self.write_config()
        shutil.copy2(PROJECT_ROOT / "snapshot_storage.py", self.app_dir)
        self.run_import(
            "import snapshot_storage; "
            "assert snapshot_storage.COMMON is None; "
            "assert snapshot_storage.snapshot_path(__import__('pathlib').Path('cloud_snapshots'), "
            "'snapshot_meta.csv').name == 'snapshot_meta.csv'"
        )


if __name__ == "__main__":
    unittest.main()
