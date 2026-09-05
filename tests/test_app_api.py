import contextlib
import io
import json
import tempfile
import unittest
from pathlib import Path
from unittest import mock

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
            api = app.Api(str(root / "default.db"), legacy_json_path=str(legacy))
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

    def test_selected_database_is_persisted_and_workbook_path_is_derived(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            config_path = root / "config.json"
            database_path = root / "chosen.db"
            api = app.Api(
                str(root / "initial.db"),
                config_path=str(config_path),
            )

            state = api.load_state(store=str(database_path))

            self.assertEqual(state["database_path"], str(database_path))
            self.assertEqual(state["xlsx_path"], str(root / "chosen.xlsx"))
            self.assertFalse(state["xlsx_exists"])
            self.assertEqual(
                app._load_config(str(config_path)),
                {
                    "database_path": str(database_path),
                    "operator_name": api.operator_name,
                },
            )
            self.assertNotIn(
                "xlsx_path",
                json.loads(config_path.read_text(encoding="utf-8")),
            )

    def test_legacy_config_xlsx_path_is_ignored(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            config_path = root / "config.json"
            config_path.write_text(json.dumps({
                "database_path": str(root / "ledger.db"),
                "xlsx_path": str(root / "elsewhere" / "old.xlsx"),
            }), encoding="utf-8")

            config = app._load_config(str(config_path))

            self.assertEqual(config, {"database_path": str(root / "ledger.db")})

    def test_excel_roundtrip_is_not_exposed_by_desktop_api(self):
        with tempfile.TemporaryDirectory() as tmp:
            api = app.Api(str(Path(tmp) / "ledger.db"))

            self.assertFalse(hasattr(api, "export_editable"))
            self.assertFalse(hasattr(api, "preview_import"))
            self.assertFalse(hasattr(api, "apply_import"))

    def test_ui_exposes_quotation_workspace_without_excel_roundtrip_controls(self):
        html = (Path(app.APP_DIR) / "ui" / "index.html").read_text(encoding="utf-8")
        javascript = (Path(app.APP_DIR) / "ui" / "app.js").read_text(encoding="utf-8")

        self.assertIn('data-tab="报价单"', html)
        self.assertIn('id="btnQuoteNew"', html)
        self.assertIn('id="quoteHistoryModal"', html)
        self.assertIn('"quotation_history"', javascript)
        self.assertIn('call("quotation_defaults"', javascript)
        self.assertNotIn('id="btnExportEditable"', html)
        self.assertNotIn('id="btnImport"', html)

    def test_unwritable_packaged_windows_database_moves_without_xlsx_config(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            executable = root / "app" / "BKTC_Ledger.exe"
            packaged_db = executable.parent / "data" / "BKTC_Ledger.db"
            packaged_db.parent.mkdir(parents=True)
            packaged_db.write_bytes(b"database")
            user_data = root / "user-data"

            with mock.patch.object(app, "_path_is_writable", return_value=False):
                resolved = app._resolve_windows_database_path(
                    str(packaged_db), executable=str(executable),
                    platform_name="win32", frozen=True,
                    user_data_dir=str(user_data),
                )

            self.assertEqual(resolved, str(user_data / "BKTC_Ledger.db"))
            self.assertEqual(Path(resolved).read_bytes(), b"database")
            self.assertEqual(
                app._xlsx_path_for_database(resolved),
                str(user_data / "BKTC_Ledger.xlsx"),
            )

    def test_download_update_carries_current_database_and_workbook(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            current_dir = root / "current"
            current_dir.mkdir()
            current_db = current_dir / "BKTC_Ledger.db"
            current_xlsx = current_dir / "BKTC_Ledger.xlsx"
            current_db.write_bytes(b"current-db")
            current_xlsx.write_bytes(b"current-xlsx")
            api = app.Api(str(current_db))
            api.xlsx_path = str(current_xlsx)

            archive = io.BytesIO()
            import zipfile
            with zipfile.ZipFile(archive, "w") as zf:
                zf.writestr("BKTC_Ledger/BKTC_Ledger.exe", b"exe")
                zf.writestr("BKTC_Ledger/data/BKTC_Ledger.db", b"starter-db")
                zf.writestr("BKTC_Ledger/data/BKTC_Ledger.xlsx", b"starter-xlsx")
            archive_bytes = archive.getvalue()

            class Response(io.BytesIO):
                def __enter__(self):
                    return self

                def __exit__(self, *args):
                    self.close()

            with mock.patch.object(app.os.path, "expanduser", return_value=str(root)), \
                    mock.patch.object(app.urllib.request, "urlopen", return_value=Response(archive_bytes)):
                update_dir = Path(api.download_update("https://example.test/update.zip", "BKTC_Ledger-Windows.zip"))

            self.assertEqual(
                (update_dir / "BKTC_Ledger" / "data" / "BKTC_Ledger.db").read_bytes(),
                b"current-db",
            )
            self.assertEqual(
                (update_dir / "BKTC_Ledger" / "data" / "BKTC_Ledger.xlsx").read_bytes(),
                b"current-xlsx",
            )

    def test_operator_name_is_persisted_and_used_for_audit(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            config_path = root / "config.json"
            api = app.Api(
                str(root / "ledger.db"),
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
            first = app.Api(str(db))
            second = app.Api(str(db))
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

        rows = app.Api("/tmp/bktc-query-test.db").get_unpaid_rows(data)

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["客户"], "测试客户")
        self.assertEqual(rows[0]["JOB No"], "26BS001")
        self.assertEqual(rows[0]["未回收金额"], 113)
        self.assertNotIn("inv_dt", rows[0])
        json.dumps(rows, ensure_ascii=False)

    def test_quotation_defaults_reuse_order_customer_model_and_terms(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            api = app.Api(str(root / "ledger.db"), operator_name="报价员")
            state = api.load_state()
            data = state["data"]
            data["合同订单"] = [{
                "记录ID": "contract-1", "JOB No": "26BS001", "客户": "客户A",
                "担当者": "张三", "币种": "RMB", "送货地点": "上海",
            }]
            data["付款条件"] = [{
                "记录ID": "term-1", "JOB No": "26BS001", "款类": "预付款",
                "比例%": 100, "账期天数": 0, "说明": "全额预付",
            }]
            data["发货批次"] = [{
                "记录ID": "ship-1", "JOB No": "26BS001", "发货批次": "1",
            }]
            data["设备台账"] = [
                {"记录ID": "dev-1", "JOB No": "26BS001", "设备型号": "KT1000",
                 "未税单价": 100, "是否无偿": "否", "发货批次": "1"},
                {"记录ID": "dev-2", "JOB No": "26BS001", "设备型号": "KT1000",
                 "未税单价": 100, "是否无偿": "否", "发货批次": "1"},
            ]
            api.save_data(data, state["rules"])

            defaults = api.quotation_defaults("26BS001")
            saved = api.save_quotation({
                "quote_no": "BJ-001", "quote_date": "2026-09-03",
                "customer": defaults["customer"], "source_job_no": "26BS001",
                "currency": defaults["currency"], "tax_rate": 13,
                "items": defaults["items"],
            })
            history = api.quotation_history("客户A", "KT1000")

            self.assertEqual(defaults["salesperson"], "张三")
            self.assertEqual(defaults["payment_terms"], "全额预付")
            self.assertEqual(defaults["items"], [{
                "model": "KT1000", "description": "", "quantity": 2,
                "unit": "台", "unit_price": 100,
            }])
            self.assertEqual(saved["quotation"]["customer"], "客户A")
            self.assertEqual(history[0]["quote_no"], "BJ-001")
            self.assertEqual(api.database_revision, saved["revision"])

    def test_saved_quotation_exports_to_downloads_without_changing_database(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            api = app.Api(str(root / "ledger.db"), operator_name="报价员")
            api.load_state()
            saved = api.save_quotation({
                "quote_no": "BJ/2026:001", "quote_date": "2026-09-03",
                "customer": "客户A", "currency": "RMB", "tax_rate": 13,
                "payment_terms": "签约后 30 日内付款",
                "items": [{
                    "model": "KT1000EPS-C300", "quantity": 2,
                    "unit": "台", "unit_price": 1000000,
                }],
            })
            revision = api.database_revision

            with mock.patch.object(app.os.path, "expanduser", return_value=str(root)):
                result = api.export_quotation(saved["quotation"]["quote_id"])

            exported = Path(result["path"])
            self.assertTrue(exported.is_file())
            self.assertEqual(exported.parent, root / "Downloads" / "报价单")
            self.assertNotIn("/", exported.name)
            self.assertNotIn(":", exported.name)
            self.assertEqual(api.database_revision, revision)
            self.assertEqual(result["quote_no"], "BJ/2026:001")

    def test_quotation_export_uses_timestamped_fallback_when_target_is_open(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            api = app.Api(str(root / "ledger.db"))
            api.load_state()
            saved = api.save_quotation({
                "quote_no": "BJ-OPEN", "quote_date": "2026-09-04",
                "customer": "客户A", "items": [{
                    "model": "KT1000", "quantity": 1, "unit_price": 100,
                }],
            })
            real_export = app.quotation_export.export_quotation_xlsx
            calls = []

            def export_with_open_file_fallback(quotation, path):
                calls.append(Path(path))
                if len(calls) == 1:
                    raise PermissionError(13, "file is open", path)
                return real_export(quotation, path)

            with mock.patch.object(app.os.path, "expanduser", return_value=str(root)), \
                    mock.patch.object(
                        app.quotation_export, "export_quotation_xlsx",
                        side_effect=export_with_open_file_fallback,
                    ):
                result = api.export_quotation(saved["quotation"]["quote_id"])

            self.assertEqual(len(calls), 2)
            self.assertNotEqual(calls[0], calls[1])
            self.assertTrue(Path(result["path"]).is_file())
            self.assertIn("已改用带时间戳文件", result["notice"])

    def test_quotation_defaults_keep_contract_model_when_order_has_no_device_rows(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            api = app.Api(str(root / "ledger.db"))
            state = api.load_state()
            state["data"]["合同订单"] = [{
                "记录ID": "contract-1", "JOB No": "26BS009",
                "客户": "客户A", "设备型号": "KT1000FI", "总台数": 0,
            }]
            api.save_data(state["data"], state["rules"])

            defaults = api.quotation_defaults("26BS009")

            self.assertEqual(defaults["items"], [{
                "model": "KT1000FI", "description": "", "quantity": 1,
                "unit": "台", "unit_price": 0,
            }])

    def test_quotation_subject_uses_every_model_found_in_device_rows(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            api = app.Api(str(root / "ledger.db"))
            state = api.load_state()
            state["data"]["合同订单"] = [{
                "记录ID": "contract-1", "JOB No": "26BS008", "客户": "客户A",
                "设备型号": "KT1000FI",
            }]
            state["data"]["发货批次"] = [{
                "记录ID": "ship-1", "JOB No": "26BS008", "发货批次": "1",
            }]
            state["data"]["设备台账"] = [
                {"记录ID": "dev-1", "JOB No": "26BS008", "设备型号": "KT1000FI",
                 "未税单价": 100, "是否无偿": "否", "发货批次": "1"},
                {"记录ID": "dev-2", "JOB No": "26BS008", "设备型号": "KT1000EPS-C300",
                 "未税单价": 200, "是否无偿": "否", "发货批次": "1"},
            ]
            api.save_data(state["data"], state["rules"])

            defaults = api.quotation_defaults("26BS008")

            self.assertEqual(defaults["subject"], "KT1000FI / KT1000EPS-C300")


if __name__ == "__main__":
    unittest.main()
