import json
import tempfile
import unittest
from pathlib import Path

import app
import core


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


if __name__ == "__main__":
    unittest.main()
