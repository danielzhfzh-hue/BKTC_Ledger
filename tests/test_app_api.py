import contextlib
import io
import json
import tempfile
import unittest
from pathlib import Path

import app
import core
from create_portable_starter import create_starter_files


class ApiMigrationTests(unittest.TestCase):
    def test_schema_for_js_keeps_options_and_derived_flag_in_fixed_positions(self):
        schema = app.schema_for_js()
        contract_fields = {field[0]: field for field in schema["合同订单"]}
        term_fields = {field[0]: field for field in schema["付款条件"]}

        self.assertEqual(contract_fields["设备型号"][2], [])
        self.assertTrue(contract_fields["设备型号"][3])
        self.assertFalse(contract_fields["JOB No"][3])
        self.assertIn("预付款", term_fields["款类"][2])
        self.assertFalse(term_fields["款类"][3])

    def test_default_legacy_source_is_not_reused_after_switching_database(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            legacy = root / "legacy.records.json"
            payload = {"tables": core.TABLES, "version": 1, "预警规则": {}}
            payload.update(core.empty_data())
            legacy.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
            api = app.Api(str(root / "ledger.xlsx"), str(root / "default.db"),
                          legacy_json_path=str(legacy))
            self.assertEqual(api.legacy_json_path, str(legacy))

            api._set_data_path(str(root / "another.db"))

            self.assertIsNone(api.legacy_json_path)
            state = api.load_state()
            self.assertEqual(state["database_path"], str(root / "another.db"))
            self.assertFalse(state["migration"])

    def test_portable_default_finds_data_next_to_executable_or_parent(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            app_dir = root / "BKTC_Ledger"
            app_dir.mkdir()
            executable = app_dir / "BKTC_Ledger.exe"
            parent_db = root / "BKTC_Ledger.db"
            parent_db.write_bytes(b"db")
            self.assertEqual(
                app._portable_default("BKTC_Ledger.db", str(executable)),
                str(parent_db),
            )
            local_db = app_dir / "BKTC_Ledger.db"
            local_db.write_bytes(b"db")
            self.assertEqual(
                app._portable_default("BKTC_Ledger.db", str(executable)),
                str(local_db),
            )

    def test_portable_default_prefers_sibling_data_directory(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            executable = root / "releases" / "macOS" / "BKTC_Ledger.app" / "Contents" / "MacOS" / "BKTC_Ledger"
            data_dir = root / "releases" / "macOS" / "data"
            data_dir.mkdir(parents=True)
            data_file = data_dir / "BKTC_Ledger.db"
            data_file.write_bytes(b"db")
            self.assertEqual(
                app._portable_default("BKTC_Ledger.db", str(executable)),
                str(data_file),
            )

    def test_ephemeral_path_is_rejected(self):
        self.assertTrue(app._is_ephemeral_path("/private/tmp/bktc-ledger-audit/ledger.db"))
        self.assertFalse(app._is_ephemeral_path("/Users/danielzhu/projects/订单整理/BKTC_Ledger/data/BKTC_Ledger.db"))

    def test_mac_release_copy_is_redirected_to_canonical_project_data(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            release_path = str(root / "releases" / "macOS" / "data" / "BKTC_Ledger.db")
            expected = root / "data" / "BKTC_Ledger.db"
            expected.parent.mkdir()
            expected.write_bytes(b"db")

            self.assertEqual(
                app._prefer_canonical_mac_data(
                    release_path, "BKTC_Ledger.db", "darwin", str(root)
                ),
                str(expected),
            )
            self.assertEqual(
                app._prefer_canonical_mac_data(
                    release_path, "BKTC_Ledger.db", "win32", str(root)
                ),
                release_path,
            )

    def test_source_mode_uses_project_data_directory(self):
        legacy = "/Users/example/production.db"

        path = app._source_default("BKTC_Ledger.db", legacy, platform="win32")

        self.assertEqual(path, str(Path(app.APP_DIR) / "data" / "BKTC_Ledger.db"))

    def test_portable_starter_contains_openable_database_and_workbook(self):
        with tempfile.TemporaryDirectory() as tmp:
            cp1252_stdout = io.TextIOWrapper(io.BytesIO(), encoding="cp1252")
            with contextlib.redirect_stdout(cp1252_stdout):
                database_path, workbook_path = create_starter_files(
                    Path(tmp) / "中文目录"
                )

            self.assertTrue(Path(database_path).is_file())
            self.assertTrue(Path(workbook_path).is_file())
            self.assertEqual(app.database.get_revision(database_path), 1)
            self.assertEqual(app.database.load_database(database_path), core.empty_data())

    def test_selected_database_and_workbook_paths_are_persisted(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            config_path = root / "config.json"
            database_path = root / "chosen.db"
            workbook_path = root / "chosen.xlsx"
            api = app.Api(
                str(root / "initial.xlsx"),
                str(root / "initial.db"),
                config_path=str(config_path),
            )

            state = api.load_state(
                store=str(database_path), xlsx=str(workbook_path)
            )

            self.assertEqual(state["database_path"], str(database_path))
            self.assertEqual(state["xlsx_path"], str(workbook_path))
            self.assertEqual(
                app._load_config(str(config_path)),
                {
                    "database_path": str(database_path),
                    "xlsx_path": str(workbook_path),
                    "operator_name": api.operator_name,
                },
            )

    def test_operator_name_is_persisted_and_used_for_audit(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            config_path = root / "config.json"
            api = app.Api(
                str(root / "ledger.xlsx"), str(root / "ledger.db"),
                config_path=str(config_path), operator_name="初始操作人",
            )
            state = api.load_state()
            api.set_operator_name("张三")
            data = state["data"]
            data["合同订单"] = [{
                "记录ID": "contract-1", "JOB No": "26BS001", "客户": "客户A"
            }]

            result = api.save_data(data, state["rules"], {
                "source": "manual_save", "actions": ["new_order"]
            })
            audit = api.get_audit_events({}, 20)

            self.assertGreater(result["audit_change_count"], 0)
            self.assertEqual(audit["events"][0]["operator_name"], "张三")
            self.assertIn("new_order", audit["events"][0]["actions"])
            self.assertEqual(app._load_config(str(config_path))["operator_name"], "张三")

    def test_second_app_instance_cannot_overwrite_a_newer_revision(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            db = root / "shared.db"
            first = app.Api(str(root / "ledger.xlsx"), str(db))
            second = app.Api(str(root / "ledger.xlsx"), str(db))
            first_state = first.load_state()
            second_state = second.load_state()

            first.save_data(first_state["data"], first_state["rules"])

            with self.assertRaises(app.database.StaleImportError):
                second.save_data(second_state["data"], second_state["rules"])

    def test_unpaid_rows_api_returns_json_safe_public_fields(self):
        data = core.empty_data()
        data["合同订单"] = [{"JOB No": "26BS001", "客户": "测试客户"}]
        data["付款条件"] = [{"JOB No": "26BS001", "款类": "预付款", "比例%": 100}]
        data["设备台账"] = [{
            "JOB No": "26BS001", "製造番号": "26BS001-001", "機番": "A1",
            "未税单价": 100, "是否无偿": "否", "发货批次": "1",
        }]
        data["发货批次"] = [{"JOB No": "26BS001", "发货批次": "1"}]

        rows = app.Api("", "/tmp/bktc-query-test.db").get_unpaid_rows(data)

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["客户"], "测试客户")
        self.assertEqual(rows[0]["JOB No"], "26BS001")
        self.assertEqual(rows[0]["未回收金额"], 113)
        self.assertNotIn("inv_dt", rows[0])
        json.dumps(rows, ensure_ascii=False)


if __name__ == "__main__":
    unittest.main()
