import unittest

import core


class ReceivableDateTests(unittest.TestCase):
    def test_contract_and_receipt_trigger_require_explicit_event_dates(self):
        data = core.empty_data()
        data["合同订单"] = [{"JOB No": "26BS001", "客户": "客户", "合同生效日": "2026-09-10"}]
        data["付款条件"] = [{"JOB No": "26BS001", "款类": "预付款", "比例%": 100,
                           "账期天数": 14, "触发条件": "合同生效后"}]
        data["开票记录"] = [{"记录ID": "i1", "JOB No": "26BS001", "款类": "预付款",
                           "开票日": "2026-08-01", "含税金额": 113, "覆盖批次": "1"}]
        self.assertEqual(core.derive(data)["开票记录"][0]["应收回款日"], "2026-09-24")
        data["合同订单"][0]["合同生效日"] = ""
        self.assertEqual(core.derive(data)["开票记录"][0]["应收回款日"], "")
        data["付款条件"][0]["触发条件"] = "货到签收后"
        data["设备台账"] = [{"JOB No": "26BS001", "製造番号": "A", "发货批次": "1",
                           "未税单价": 100, "送货单回收": "已签收", "签收日": ""}]
        self.assertEqual(core.derive(data)["开票记录"][0]["应收回款日"], "")
        data["设备台账"][0]["签收日"] = "2026-09-12"
        self.assertEqual(core.derive(data)["开票记录"][0]["应收回款日"], "2026-09-26")

    def test_acceptance_trigger_uses_latest_covered_date_and_requires_all_dates(self):
        data = core.empty_data()
        data["合同订单"] = [{"JOB No": "26BS001", "客户": "客户"}]
        data["付款条件"] = [{"JOB No": "26BS001", "款类": "验收款", "比例%": 100,
                           "账期天数": 30, "触发条件": "验收合格后"}]
        data["设备台账"] = [
            {"JOB No": "26BS001", "製造番号": "A", "发货批次": "1", "未税单价": 100, "质保开始日": "2026-09-01"},
            {"JOB No": "26BS001", "製造番号": "B", "发货批次": "1", "未税单价": 100, "质保开始日": "2026-09-05"},
        ]
        data["开票记录"] = [{"记录ID": "i1", "JOB No": "26BS001", "款类": "验收款",
                           "开票日": "2026-08-01", "覆盖批次": "1", "含税金额": 226}]
        self.assertEqual(core.derive(data)["开票记录"][0]["应收回款日"], "2026-10-05")
        data["设备台账"][1]["质保开始日"] = ""
        self.assertEqual(core.derive(data)["开票记录"][0]["应收回款日"], "")

    def invoice_due(self, days, trigger="开票后", opened="2026-08-28"):
        data = core.empty_data()
        data["合同订单"] = [{"JOB No": "26BS003", "客户": "GTX"}]
        data["付款条件"] = [{"JOB No": "26BS003", "款类": "到货款",
                           "比例%": 100, "账期天数": 60, "触发条件": trigger}]
        data["开票记录"] = [{"JOB No": "26BS003", "款类": "到货款",
                           "开票日": opened, "账期天数": days}]
        return core.derive(data)["开票记录"][0]["应收回款日"]

    def test_blank_invoice_days_inherit_payment_term(self):
        for blank in (None, ""):
            with self.subTest(blank=blank):
                self.assertEqual(self.invoice_due(blank), "2026-10-27")

    def test_explicit_zero_days_are_not_replaced(self):
        self.assertEqual(self.invoice_due(0), "2026-08-28")

    def test_explicit_invoice_days_override_term(self):
        self.assertEqual(self.invoice_due(14), "2026-09-11")

    def test_next_month_end_handles_calendar_boundaries(self):
        for opened, expected in (("2026-08-28", "2026-09-30"),
                                 ("2026-12-05", "2027-01-31"),
                                 ("2028-01-10", "2028-02-29"),
                                 ("2027-01-10", "2027-02-28")):
            with self.subTest(opened=opened):
                self.assertEqual(self.invoice_due(0, "月结次月月底", opened), expected)
