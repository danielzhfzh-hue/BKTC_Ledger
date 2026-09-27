import unittest
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
