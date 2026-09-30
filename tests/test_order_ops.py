import unittest
from unittest.mock import patch
from datetime import date

import core
import order_ops


def sample_domain_data():
    data = core.empty_data()
    data["合同订单"] = [{"记录ID":"c1","JOB No":"26BS001","客户":"GTX","担当者":"销售A","币种":"RMB","来源报价ID":"q1","来源报价单号":"Q26-001"}]
    data["付款条件"] = [{"记录ID":"t1","JOB No":"26BS001","款类":"发货款","比例%":100,"账期天数":30,"触发条件":"开票后","说明":"开票后30天"}]
    data["发货批次"] = [
        {"记录ID":"s1","JOB No":"26BS001","发货批次":"1","出荷日":"2026-09-20"},
        {"记录ID":"s2","JOB No":"26BS001","发货批次":"2","出荷日":"2026-09-30"},
    ]
    data["设备台账"] = [
        {"记录ID":"d1","JOB No":"26BS001","设备型号":"KPL-W13u-DP","製造番号":"SN001","未税单价":100,"发货批次":"1","质保开始日":"2026-09-22"},
        {"记录ID":"d2","JOB No":"26BS001","设备型号":"KPL-W13u-DP","製造番号":"","未税单价":100,"发货批次":"2"},
    ]
    data["开票记录"] = [{"记录ID":"i1","JOB No":"26BS001","款类":"发货款","开票日":"2026-08-01","状态":"已开","含税金额":226,"账期天数":30,"应收回款日":"2026-08-31","回款状态":"超期未回","覆盖批次":"1;2"}]
    return data


class OrderOpsTests(unittest.TestCase):
    def test_workspace_reporting_date_reaches_underlying_warning_calculation(self):
        workspace = order_ops.build_workspace(sample_domain_data(), today="2026-07-01")
        self.assertTrue(workspace["finance"]["rows"])
        self.assertTrue(all(row["预警等级"] == "③未到期" for row in workspace["finance"]["rows"]))
        self.assertTrue(all("逾期" not in row["未回收原因"] for row in workspace["finance"]["rows"]))

    def test_batch_and_serial_payments_both_count_when_they_are_distinct(self):
        data = sample_domain_data()
        data["设备台账"] = data["设备台账"][:1]
        data["发货批次"] = data["发货批次"][:1]
        data["开票记录"][0].update({"含税金额": 113, "覆盖批次": "1", "覆盖製造番号": "SN001"})
        data["回款记录"] = [
            {"记录ID": "p1", "JOB No": "26BS001", "款类": "发货款", "回款日": "2026-09-01", "含税金额": 40, "覆盖批次": "1"},
            {"记录ID": "p2", "JOB No": "26BS001", "款类": "发货款", "回款日": "2026-09-02", "含税金额": 73, "覆盖批次": "1", "覆盖製造番号": "SN001"},
        ]
        self.assertEqual(core.unpaid_report_rows(data), [])
        self.assertEqual(order_ops.build_order_snapshots(data)[0]["outstanding"], 0)

    def test_fully_prepaid_order_does_not_skip_fulfillment(self):
        data = sample_domain_data()
        for shipment in data["发货批次"]:
            shipment["出荷日"] = "2026-10-20"
        for device in data["设备台账"]:
            device["质保开始日"] = ""
        data["付款条件"][0]["款类"] = "预付款"
        data["开票记录"][0]["款类"] = "预付款"
        data["回款记录"] = [{"记录ID": "p1", "JOB No": "26BS001", "款类": "预付款",
                            "回款日": "2026-09-01", "含税金额": 226,
                            "覆盖批次": "1;2"}]
        snapshot = order_ops.build_order_snapshots(data, today="2026-09-27")[0]
        self.assertEqual(snapshot["stage"], "履约/出货")
        self.assertEqual(snapshot["paid_pct"], 100)
        self.assertEqual(snapshot["financial_status"], "已收清")

    def test_settlement_requires_exact_amount_not_rounded_percentage(self):
        data = sample_domain_data()
        for shipment in data["发货批次"]:
            shipment["出荷日"] = "2026-09-20"
        for device in data["设备台账"]:
            device["质保开始日"] = "2026-09-22"
        data["开票记录"][0]["含税金额"] = 224
        data["回款记录"] = [{"记录ID": "p1", "JOB No": "26BS001", "款类": "发货款",
                            "回款日": "2026-09-01", "含税金额": 226,
                            "覆盖批次": "1;2"}]
        snapshot = order_ops.build_order_snapshots(data, today="2026-09-27")[0]
        self.assertNotEqual(snapshot["stage"], "已结清")

    def test_missing_terms_are_visible_and_not_zero_or_settled(self):
        data = sample_domain_data()
        data["付款条件"] = []
        data["开票记录"] = []
        workspace = order_ops.build_workspace(data, today="2026-09-27")
        row = workspace["finance"]["rows"][0]
        self.assertIsNone(row["amount"])
        self.assertEqual(row["预警等级"], "④待确认")
        self.assertEqual(workspace["orders"][0]["risk_code"], "terms_missing")
        self.assertIsNone(workspace["orders"][0]["outstanding"])
        self.assertEqual(core.summary(data)["未回收金额待确认行数"], 1)

    def test_placeholder_and_incomplete_terms_do_not_appear_settled(self):
        for terms in ([{"JOB No": "26BS001", "款类": "", "比例%": None}],
                      [{"JOB No": "26BS001", "款类": "预付款", "比例%": None}]):
            with self.subTest(terms=terms):
                data = sample_domain_data()
                data["付款条件"] = terms
                data["开票记录"] = []
                workspace = order_ops.build_workspace(data, today="2026-09-27")
                self.assertIsNone(workspace["finance"]["rows"][0]["amount"])
                self.assertIsNone(workspace["orders"][0]["outstanding"])
                self.assertNotEqual(workspace["orders"][0]["stage"], "已结清")

    def test_workspace_derives_once_and_matches_independent_views(self):
        data = sample_domain_data()
        expected_orders = order_ops.build_order_snapshots(data, today="2026-09-27")
        expected_finance = order_ops.build_finance(data, today="2026-09-27")
        expected_actions = order_ops.build_actions(data, today="2026-09-27")
        with patch.object(core, "derive", wraps=core.derive) as derive:
            workspace = order_ops.build_workspace(data, today="2026-09-27")
            self.assertEqual(derive.call_count, 1)
        self.assertEqual(workspace["orders"], expected_orders)
        self.assertEqual(workspace["finance"], expected_finance)
        self.assertEqual(workspace["actions"], expected_actions)

    def test_pending_receivable_is_not_overdue_despite_provisional_date(self):
        pending = {"JOB No": "26BS001", "批次": "1", "款类": "验收款",
                   "预警等级": "④待确认", "预定回收日期": "2026-08-01",
                   "未回收金额": 113, "未回收原因": "未验收；待确认"}
        with patch.object(core, "unpaid_report_rows", return_value=[pending]):
            finance = order_ops.build_finance(sample_domain_data(), today="2026-09-27")
            self.assertIsNone(finance["rows"][0]["days"])
            self.assertEqual(finance["overdue_amounts"], {})
            bucket = next(x for x in finance["aging"] if x["bucket"] == "待确认")
            self.assertEqual(bucket["amounts"], {"RMB": 113})
            snapshot = order_ops.build_order_snapshots(sample_domain_data(), today="2026-09-27")[0]
            self.assertNotEqual(snapshot["risk_code"], "ar_overdue")
            actions = order_ops.build_actions(sample_domain_data(), today="2026-09-27")
            self.assertTrue(any(x["type"] == "ar_confirm" for x in actions))
            self.assertFalse(any(x["type"] == "ar_overdue" for x in actions))

    def test_snapshot_uses_business_date_and_document_flow(self):
        workspace = order_ops.build_workspace(sample_domain_data(), today=date(2026, 9, 27))
        snap = workspace["orders"][0]
        self.assertEqual(snap["shipped_count"], 1)
        self.assertEqual(snap["scheduled_count"], 1)
        self.assertEqual(snap["accepted_count"], 1)
        self.assertEqual(snap["pending_serial"], 1)
        self.assertEqual(snap["source_quote_no"], "Q26-001")
        node_types = {x["type"] for x in snap["document_flow"]["nodes"]}
        self.assertTrue({"quote", "order", "shipment", "invoice"}.issubset(node_types))

    def test_actions_prioritize_overdue_and_upcoming_shipment(self):
        actions = order_ops.build_actions(sample_domain_data(), today="2026-09-27")
        self.assertEqual(actions[0]["type"], "ar_overdue")
        shipment = next(x for x in actions if x["type"] == "shipment_due")
        self.assertIn("3 天内", shipment["title"])
        self.assertIn("1 台待编号", shipment["detail"])

    def test_finance_aging_and_forecast_are_currency_safe(self):
        finance = order_ops.build_finance(sample_domain_data(), today="2026-09-27")
        overdue = next(x for x in finance["aging"] if x["bucket"] == "逾期0-30天")
        self.assertEqual(overdue["amounts"]["RMB"], 226)
        self.assertEqual(finance["overdue_amounts"]["RMB"], 226)

    def test_manual_work_items_are_merged_into_actions(self):
        actions = order_ops.build_actions(sample_domain_data(), today="2026-09-27", work_items=[
            {"item_id":"w1","job_no":"26BS001","title":"确认客户搬入时间","due_date":"2026-09-28","priority":"high","status":"open","note":"微信确认"}
        ])
        manual = next(x for x in actions if x["manual"])
        self.assertEqual(manual["source_id"], "w1")


    def test_document_flow_links_payment_to_matching_invoice(self):
        data = sample_domain_data()
        data["回款记录"] = [{
            "记录ID": "p1", "JOB No": "26BS001", "款类": "发货款",
            "回款日": "2026-09-10", "含税金额": 50, "覆盖批次": "1;2",
        }]
        flow = order_ops.build_order_snapshots(data, today="2026-09-27")[0]["document_flow"]
        invoice = next(x for x in flow["nodes"] if x["type"] == "invoice")
        payment = next(x for x in flow["nodes"] if x["type"] == "payment")
        self.assertIn({"from": invoice["id"], "to": payment["id"]}, flow["edges"])


if __name__ == "__main__":
    unittest.main()
