import tempfile
import unittest
from pathlib import Path
from unittest import mock

from openpyxl import load_workbook

import app
import core
import move_in_request


def sample_data(device_count=2):
    data = core.empty_data()
    data["合同订单"] = [{
        "记录ID": "contract-1", "JOB No": "26BS006", "客户": "搬入客户",
        "送货地点": "上海市嘉定区测试路 1 号", "发货地点": "平湖康肯",
    }]
    data["发货批次"] = [
        {"记录ID": "ship-1", "JOB No": "26BS006", "发货批次": "1", "出荷日": "2026-08-28"},
        {"记录ID": "ship-2", "JOB No": "26BS006", "发货批次": "2", "出荷日": "2026-09-01"},
    ]
    data["设备台账"] = [{
        "记录ID": f"dev-{i}", "JOB No": "26BS006", "发货批次": "1",
        "设备型号": "KT1000FI" if i == 0 else "KT1000EPS-C300",
        "機番": f"2580{i}", "製造番号": f"26BS006-{i + 1:02d}",
        "PO No": "PO-1" if i == 0 else "PO-2",
    } for i in range(device_count)]
    return data


class MoveInRequestTests(unittest.TestCase):
    def test_defaults_use_exact_job_and_batch_and_do_not_confuse_ship_date(self):
        data = sample_data()
        defaults = move_in_request.shipment_defaults(data, shipment_id="ship-1")

        self.assertEqual(defaults["batch"], "1")
        self.assertEqual(defaults["address"], "上海市嘉定区测试路 1 号")
        self.assertEqual(defaults["ship_date"], "2026-08-28")
        self.assertEqual(defaults["move_in_date"], "")
        self.assertEqual(len(defaults["devices"]), 2)
        self.assertIn("PO-1", defaults["notes"])

    def test_missing_devices_is_rejected(self):
        data = sample_data()
        data["设备台账"] = []
        with self.assertRaisesRegex(ValueError, "没有关联设备"):
            move_in_request.shipment_defaults(data, shipment_id="ship-1")

    def test_workbook_is_a4_and_keeps_all_devices_in_appendix(self):
        data = sample_data(device_count=6)
        form = move_in_request.shipment_defaults(data, shipment_id="ship-1")
        form.update({"move_in_date": "2026-08-31", "contact": "测试联系人 13800000000"})

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "搬入依頼書.xlsx"
            move_in_request.export_xlsx(form, path)
            workbook = load_workbook(path, data_only=False)
            sheet = workbook["搬入依頼書"]
            self.assertEqual(sheet.print_area, "'搬入依頼書'!$A$1:$Q$35")
            self.assertEqual(sheet.page_setup.orientation, "portrait")
            self.assertEqual(sheet.page_setup.fitToWidth, 1)
            self.assertEqual(sheet.page_setup.fitToHeight, 1)
            self.assertEqual(workbook.sheetnames, ["搬入依頼書", "設備明細"])
            self.assertEqual(workbook["設備明細"].max_row, 7)
            self.assertEqual(sheet["B13"].value, "KT1000FI")
            self.assertEqual(sheet["C13"].value, "25800")

    def test_api_export_reads_current_database_and_writes_installation_folder(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            api = app.Api(str(root / "ledger.db"), operator_name="测试员")
            state = api.load_state()
            data = sample_data()
            api.save_data(data, state["rules"])
            defaults = api.get_move_in_request_defaults("ship-1")
            defaults.update({"move_in_date": "2026-08-31", "contact": "测试联系人"})

            output_dir = root / "搬入依頼書"
            with mock.patch.object(
                app, "_move_in_request_output_dir", return_value=(str(output_dir), "")
            ):
                result = api.export_move_in_request(defaults)

            output = Path(result["path"])
            self.assertTrue(output.is_file())
            self.assertEqual(output.parent, output_dir)
            self.assertEqual(result["device_count"], 2)
            self.assertEqual(load_workbook(output).active["B13"].value, "KT1000FI")

    def test_api_export_rejects_stale_database_revision(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            api = app.Api(str(root / "ledger.db"))
            state = api.load_state()
            api.save_data(sample_data(), state["rules"])
            form = api.get_move_in_request_defaults("ship-1")
            form.update({"move_in_date": "2026-08-31"})
            api.save_data(sample_data(), state["rules"])
            with self.assertRaises(app.database.StaleImportError):
                api.export_move_in_request(form)


if __name__ == "__main__":
    unittest.main()
