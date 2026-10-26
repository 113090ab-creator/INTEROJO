import gzip
import hashlib
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock


sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import snapshot_mirror as mirror_module


def published(slot="2026-10-26 AM", set_id="aps_new", checked="2026-10-26 10:32:53"):
    pointer = {"set_id": set_id, "slot_key": slot, "manifest_path": f"sets/{set_id}/manifest.json",
               "plan_updated_at": "2026-10-26 07:24:25", "wip_updated_at": "2026-10-26 07:35:19",
               "published_at": "2026-10-26 09:27:01"}
    manifest = {**pointer, "files": {}}
    files = {}
    for name, text in (("shortage_snapshot.csv.gz", "item,quantity\nnew,7\n"),
                       ("wip_inventory_snapshot.csv.gz", "item,quantity\nnew,3\n")):
        data = text.encode()
        path = f"sets/{set_id}/{name}"
        files[path] = gzip.compress(data)
        manifest["files"][name] = {"path": path, "checksum": hashlib.sha256(data).hexdigest(), "rows": 1}
    files[mirror_module.POINTER] = json.dumps(pointer).encode()
    files[pointer["manifest_path"]] = json.dumps(manifest).encode()
    files[mirror_module.STATUS] = json.dumps({"status": "PUBLISHED", "set_id": set_id,
                                            "slot_key": slot, "checked_at": checked}).encode()
    files[mirror_module.META] = b"key,value\ndata_updated_at,2026-10-26 07:24:25\n"
    return files


def write_tree(root, files):
    for name, data in files.items():
        path = root / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(data)


class SnapshotMirrorTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.source = Path(self.temp.name) / "source"
        self.source.mkdir()
        write_tree(self.source, published("2026-10-24 AM", "aps_old", "2026-10-24 11:02:13"))
        self.remote = published()
        self.fetch = mock.Mock(side_effect=lambda name: self.remote[name])
        self.mirror = mirror_module.SnapshotMirror(self.source, Path(self.temp.name) / "mirror",
                                                   "owner/repo", fetch=self.fetch)

    def test_stale_checkout_is_replaced_for_reads_without_changing_git_files(self):
        original = (self.source / mirror_module.POINTER).read_bytes()
        self.mirror.refresh()
        pointer = mirror_module.read_json(self.mirror.read_path(mirror_module.POINTER))
        self.assertEqual(pointer["set_id"], "aps_new")
        self.assertEqual((self.source / mirror_module.POINTER).read_bytes(), original)
        data = gzip.decompress(self.mirror.read_path("shortage_snapshot.csv.gz").read_bytes())
        self.assertIn(b"new,7", data)

    def test_failed_download_preserves_previous_validated_set(self):
        self.mirror.refresh()
        before = self.mirror.version()
        self.remote = published("2026-10-26 PM", "aps_evening")
        del self.remote["sets/aps_evening/wip_inventory_snapshot.csv.gz"]
        self.mirror.refresh(force=True)
        self.assertEqual(self.mirror.version(), before)
        self.assertEqual(mirror_module.read_json(self.mirror.read_path(mirror_module.POINTER))["set_id"], "aps_new")
        self.assertFalse((self.mirror.cache / "sets/aps_evening").exists())

    def test_checksum_failure_does_not_publish_partial_snapshot(self):
        self.remote["sets/aps_new/shortage_snapshot.csv.gz"] = gzip.compress(b"wrong contents")
        self.mirror.refresh()
        self.assertIsNone(self.mirror.bundle())
        self.assertFalse((self.mirror.cache / "active.json").exists())
        self.assertEqual(self.mirror.read_path(mirror_module.POINTER), self.source / mirror_module.POINTER)

    def test_network_error_retains_data_and_throttles_retries(self):
        self.fetch.side_effect = OSError("offline")
        self.mirror.refresh()
        self.mirror.refresh()
        self.assertEqual(self.fetch.call_count, 1)
        self.assertEqual(self.mirror.last_error, "OSError")

    def test_unchanged_publish_does_not_invalidate_data_cache(self):
        self.mirror.refresh()
        path = self.mirror.read_path(mirror_module.POINTER)
        modified = path.stat().st_mtime_ns
        self.mirror.refresh(force=True)
        self.assertEqual(path.stat().st_mtime_ns, modified)

    def test_fresh_checkout_uses_local_data_without_remote_data_downloads(self):
        write_tree(self.source, self.remote)
        self.mirror.refresh()
        names = [call.args[0] for call in self.fetch.call_args_list]
        self.assertNotIn("sets/aps_new/shortage_snapshot.csv.gz", names)
        self.assertIsNotNone(self.mirror.bundle())

    def test_remote_never_regresses_a_newer_checkout(self):
        write_tree(self.source, published("2026-10-26 PM", "aps_local_evening"))
        self.mirror.refresh()
        self.assertIsNone(self.mirror.bundle())

    def test_checkout_catches_up_after_mirror_and_newer_status_wins(self):
        self.mirror.refresh()
        write_tree(self.source, published(checked="2026-10-26 11:32:53"))
        self.assertIsNone(self.mirror.bundle())

    def test_next_period_is_adopted_without_restart(self):
        self.mirror.refresh()
        before = self.mirror.version()
        self.remote = published("2026-10-26 PM", "aps_evening")
        self.mirror.refresh(force=True)
        self.assertNotEqual(self.mirror.version(), before)
        self.assertEqual(mirror_module.read_json(self.mirror.read_path(mirror_module.POINTER))["slot_key"], "2026-10-26 PM")

    def test_immutable_published_set_cannot_be_overwritten_in_place(self):
        self.mirror.refresh()
        before = self.mirror.version()
        path = "sets/aps_new/manifest.json"
        manifest = json.loads(self.remote[path])
        manifest["files"]["shortage_snapshot.csv.gz"]["checksum"] = "a" * 64
        self.remote[path] = json.dumps(manifest).encode()
        self.mirror.refresh(force=True)
        self.assertEqual(self.mirror.version(), before)
        self.assertIn(b"new,7", gzip.decompress(self.mirror.read_path("shortage_snapshot.csv.gz").read_bytes()))

    def test_cache_cleanup_protects_active_and_source_sets(self):
        import os
        import time
        self.mirror.refresh()
        sets = self.mirror.cache / "sets"
        for index in range(8):
            directory = sets / f"aps_old_cache_{index}"
            directory.mkdir()
            os.utime(directory, (time.time() - 172800 - index, time.time() - 172800 - index))
        deployed = sets / "aps_old"
        deployed.mkdir()
        os.utime(deployed, (0, 0))
        self.mirror.prune_cache("aps_new", self.mirror.bundle().name)
        self.assertTrue((sets / "aps_new").is_dir())
        self.assertTrue(deployed.is_dir())
        self.assertLess(len(list(sets.iterdir())), 10)

    def test_path_traversal_manifest_is_rejected(self):
        path = "sets/aps_new/manifest.json"
        manifest = json.loads(self.remote[path])
        manifest["files"]["../escape.csv.gz"] = {"path": "../escape.csv.gz", "checksum": "a" * 64}
        self.remote[path] = json.dumps(manifest).encode()
        self.mirror.refresh()
        self.assertIsNone(self.mirror.bundle())

    def test_app_context_and_csv_reads_use_same_validated_mirror(self):
        import app
        with (mock.patch.object(app, "CLOUD_SNAPSHOT_DIR", self.source),
              mock.patch.object(app, "get_published_snapshot_mirror", return_value=self.mirror)):
            app.read_cloud_snapshot_context_cached.clear()
            app.read_cloud_snapshot_csv.clear()
            context = app.get_cloud_snapshot_context()
            self.assertEqual(context["current_set"]["set_id"], "aps_new")
            frame = app.load_cloud_snapshot_csv("sets/aps_new/shortage_snapshot.csv.gz")
            self.assertEqual(frame.loc[0, "quantity"], 7)
            app.read_cloud_snapshot_context_cached.clear()
            app.read_cloud_snapshot_csv.clear()


if __name__ == "__main__":
    unittest.main()
