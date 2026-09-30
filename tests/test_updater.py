import hashlib
import json
import os
import sqlite3
import tarfile
import tempfile
import unittest
import zipfile
from pathlib import Path

import updater


class UpdaterTests(unittest.TestCase):
    def test_checksum_is_verified_before_install(self):
        with tempfile.TemporaryDirectory() as tmp:
            package = Path(tmp) / "update.zip"
            package.write_bytes(b"release package")
            digest = hashlib.sha256(package.read_bytes()).hexdigest()
            self.assertEqual(updater.verify_sha256(package, f"{digest}  update.zip"), digest)
            with self.assertRaisesRegex(ValueError, "校验失败"):
                updater.verify_sha256(package, "0" * 64)

    def test_extractor_rejects_zip_path_traversal(self):
        with tempfile.TemporaryDirectory() as tmp:
            package = Path(tmp) / "bad.zip"
            with zipfile.ZipFile(package, "w") as archive:
                archive.writestr("../outside.txt", "bad")
            with self.assertRaisesRegex(ValueError, "不安全路径"):
                updater.extract_archive(package, Path(tmp) / "extract")

    def test_extractor_rejects_tar_symlink(self):
        with tempfile.TemporaryDirectory() as tmp:
            package = Path(tmp) / "bad.tar.gz"
            with tarfile.open(package, "w:gz") as archive:
                link = tarfile.TarInfo("BKTC_Ledger/escape")
                link.type = tarfile.SYMTYPE
                link.linkname = "../../outside"
                archive.addfile(link)
            with self.assertRaisesRegex(ValueError, "符号链接超出目标目录"):
                updater.extract_archive(package, Path(tmp) / "extract")

    @unittest.skipIf(os.name == "nt", "app-bundle symlinks are only used on macOS")
    def test_extractor_preserves_safe_internal_tar_symlink(self):
        with tempfile.TemporaryDirectory() as tmp:
            package = Path(tmp) / "app.tar.gz"
            with tarfile.open(package, "w:gz") as archive:
                data = b"executable"
                file_entry = tarfile.TarInfo("BKTC_Ledger.app/Contents/MacOS/app")
                file_entry.size = len(data)
                file_entry.mode = 0o755
                import io
                archive.addfile(file_entry, io.BytesIO(data))
                link = tarfile.TarInfo("BKTC_Ledger.app/Contents/MacOS/alias")
                link.type = tarfile.SYMTYPE
                link.linkname = "app"
                archive.addfile(link)
            output = updater.extract_archive(package, Path(tmp) / "extract")
            alias = output / "BKTC_Ledger.app/Contents/MacOS/alias"
            self.assertTrue(alias.is_symlink())
            self.assertEqual(alias.read_bytes(), b"executable")
            self.assertTrue(os.access(alias, os.X_OK))

    def test_sqlite_snapshot_includes_wal_committed_data(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "ledger.db"
            snapshot = Path(tmp) / "backup.db"
            connection = sqlite3.connect(source)
            connection.execute("PRAGMA journal_mode=WAL")
            connection.execute("CREATE TABLE entries (value TEXT)")
            connection.execute("INSERT INTO entries VALUES ('kept')")
            connection.commit()
            updater.snapshot_sqlite(source, snapshot)
            self.assertEqual(
                sqlite3.connect(snapshot).execute("SELECT value FROM entries").fetchone()[0],
                "kept",
            )
            connection.close()

    def test_failed_health_check_restores_program_database_and_workbook(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            target, stage = root / "app", root / "stage"
            target.mkdir()
            stage.mkdir()
            (target / "program.bin").write_text("old")
            (stage / "program.bin").write_text("new")
            db = root / "ledger.db"
            connection = sqlite3.connect(db)
            connection.execute("CREATE TABLE entries (value TEXT)")
            connection.execute("INSERT INTO entries VALUES ('before')")
            connection.commit()
            connection.close()
            db_backup = root / "ledger_backup.db"
            updater.snapshot_sqlite(db, db_backup)
            connection = sqlite3.connect(db)
            connection.execute("UPDATE entries SET value='partial-migration'")
            connection.commit()
            connection.close()
            xlsx = root / "ledger.xlsx"
            xlsx.write_bytes(b"old workbook")
            xlsx_backup = root / "ledger_backup.xlsx"
            xlsx_backup.write_bytes(b"old workbook")
            health = root / "healthy.json"
            result = root / "update-result.json"
            calls = []

            class FailedApp:
                def poll(self):
                    return 1

                def terminate(self):
                    pass

                def wait(self, timeout=None):
                    return 0

                def kill(self):
                    pass

            def launch(args, cwd):
                calls.append(args)
                return FailedApp()

            state = {
                "target_dir": str(target), "stage_dir": str(stage),
                "previous_dir": str(root / "previous"), "parent_pid": 10,
                "executable_relative": "program.bin", "health_marker": str(health),
                "result_path": str(result), "database_path": str(db),
                "database_backup": str(db_backup), "xlsx_path": str(xlsx),
                "xlsx_backup": str(xlsx_backup), "xlsx_existed": True,
            }
            outcome = updater.apply_update(
                state, wait_for_exit=lambda *_: None, launch=launch, health_timeout=0.01
            )
            self.assertFalse(outcome["ok"])
            self.assertEqual((target / "program.bin").read_text(), "old")
            connection = sqlite3.connect(db)
            self.assertEqual(connection.execute("SELECT value FROM entries").fetchone()[0], "before")
            connection.close()
            self.assertEqual(xlsx.read_bytes(), b"old workbook")
            self.assertEqual(len(calls), 2)
            self.assertEqual(json.loads(result.read_text()) ["ok"], False)

    def test_healthy_startup_commits_new_program_and_discards_stage(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            target = root / "app"
            stage_parent = root / "stage-parent"
            stage = stage_parent / "program"
            target.mkdir()
            stage.mkdir(parents=True)
            (target / "program.bin").write_text("old")
            (stage / "program.bin").write_text("new")
            marker = root / "healthy.json"
            temp_dir = root / "download"
            temp_dir.mkdir()
            calls = []

            class HealthyApp:
                def poll(self):
                    return None

            def launch(args, cwd):
                calls.append(args)
                Path(args[2]).write_text('{"version":"1.14.4"}')
                return HealthyApp()

            result = updater.apply_update(
                {
                    "target_dir": str(target), "stage_dir": str(stage),
                    "stage_parent": str(stage_parent), "temporary_dir": str(temp_dir),
                    "previous_dir": str(root / "previous"), "parent_pid": 10,
                    "executable_relative": "program.bin", "health_marker": str(marker),
                    "result_path": str(root / "result.json"), "version": "1.14.4",
                },
                wait_for_exit=lambda *_: None, launch=launch, health_timeout=0.1,
            )
            self.assertTrue(result["ok"])
            self.assertEqual((target / "program.bin").read_text(), "new")
            self.assertFalse((root / "previous").exists())
            self.assertFalse(stage_parent.exists())
            self.assertFalse(temp_dir.exists())
            self.assertEqual(len(calls), 1)


if __name__ == "__main__":
    unittest.main()
