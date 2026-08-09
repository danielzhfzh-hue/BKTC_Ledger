import unittest

from openpyxl import Workbook

import build_ledger_main as ledger


class BuildEngineReliabilityTests(unittest.TestCase):
    def test_coverage_parser_accepts_chinese_semicolon_and_line_breaks(self):
        self.assertEqual(ledger.coverage_set("A；B\nC"), {"A", "B", "C"})

    def test_user_text_that_looks_like_formula_is_escaped(self):
        ws = Workbook().active

        ledger.set_cell(ws, 1, 1, '=HYPERLINK("https://bad.example")')
        ledger.set_cell(ws, 2, 1, "=SUM(A1:A1)", trusted_formula=True)

        self.assertEqual(ws.cell(1, 1).value, "'=HYPERLINK(\"https://bad.example\")")
        self.assertEqual(ws.cell(2, 1).value, "=SUM(A1:A1)")

    def test_no_device_job_aggregates_multiple_invoices_and_payments(self):
        wb = Workbook()
        ws = wb.active
        contract = {
            "担当者": "担当", "客户": "客户", "JOB No": "26BS001",
            "订单内容": "设备", "付款条件": "预付", "发货方式": "",
            "发货地点": "", "送货地点": "",
        }
        terms = [{"款类": "预付款"}]
        invoices = [
            {"记录ID": "i1", "款类": "预付款", "开票日": "2026-08-01",
             "含税金额": 100, "应收回款日": "2026-08-10"},
            {"记录ID": "i2", "款类": "预付款", "开票日": "2026-08-03",
             "含税金额": 200, "应收回款日": "2026-08-12"},
        ]
        payments = [
            {"记录ID": "p1", "款类": "预付款", "回款日": "2026-08-05", "含税金额": 30},
            {"记录ID": "p2", "款类": "预付款", "回款日": "2026-08-06", "含税金额": 50},
        ]

        ledger.write_job_sheet(
            ws, "26BS001", contract, [], terms, [], {}, invoices, payments
        )

        headers = {ws.cell(5, col).value: col for col in range(1, ws.max_column + 1)}
        self.assertEqual(ws.cell(6, headers["预付款·开票金额"]).value, 300)
        self.assertEqual(ws.cell(6, headers["预付款·回款金额"]).value, 80)
        self.assertEqual(ws.cell(6, headers["预付款·开票日"]).value, "2026/08/03")
        self.assertEqual(ws.cell(6, headers["预付款·回款日"]).value, "2026/08/06")

    def test_no_device_job_without_payment_ratio_is_visible_for_confirmation(self):
        contracts = [{"JOB No": "26BS001", "客户": "客户"}]
        invoices = [{
            "记录ID": "i1", "JOB No": "26BS001", "款类": "预付款",
            "含税金额": 100, "开票日": "2026-08-01", "应收回款日": "2026-08-10",
        }]
        payments = [{
            "记录ID": "p1", "JOB No": "26BS001", "款类": "预付款",
            "含税金额": 20, "回款日": "2026-08-05",
        }]

        rows = ledger.compute_unpaid_rows(contracts, [], [], invoices, payments, [])

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["预警等级"], "④待确认")
        self.assertEqual(rows[0]["未回收金额"], 80)
        self.assertIn("付款比例未配置", rows[0]["未回收原因"])


if __name__ == "__main__":
    unittest.main()
