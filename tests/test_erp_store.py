import tempfile
import unittest
from pathlib import Path

import database
import erp_store
from tests.test_database import sample_data


class ErpStoreTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.db = Path(self.tmp.name) / "erp.db"
        self.initial = database.save_database(
            self.db, sample_data(), backup=False, reason="database_created"
        )

    def tearDown(self):
        self.tmp.cleanup()

    def test_master_data_is_progressive_and_audited(self):
        master = erp_store.list_master_data(self.db)
        customer = next(x for x in master["customers"] if x["name"] == "原客户")
        model = next(x for x in master["models"] if x["model"] == "KT1000")
        self.assertTrue(customer["inferred"])
        self.assertTrue(model["inferred"])

        saved = erp_store.save_master_record(
            self.db, "customer",
            {"name": "原客户", "currency": "RMB", "salesperson": "朱方周", "ship_to": "上海"},
            expected_revision=self.initial["revision"],
            backup=False,
            audit_context={"operator_name": "测试", "source": "master_data_save", "actions": [], "app_version": "test"},
        )
        self.assertGreater(saved["revision"], self.initial["revision"])
        persisted = next(x for x in erp_store.list_master_data(self.db)["customers"] if x["name"] == "原客户")
        self.assertNotIn("inferred", persisted)
        self.assertEqual(persisted["salesperson"], "朱方周")
        latest = database.get_audit_events(self.db)["events"][0]
        self.assertTrue(any(x["table_name"] == "客户主数据" for x in latest["changes"]))

    def test_work_item_lifecycle_uses_revision_and_audit(self):
        created = erp_store.save_work_item(
            self.db,
            {"job_no": "26BS001", "title": "确认搬入时间", "due_date": "2026-10-01", "priority": "high"},
            expected_revision=self.initial["revision"],
            backup=False,
            audit_context={"operator_name": "测试", "source": "work_item_save", "actions": [], "app_version": "test"},
        )
        item = created["work_item"]
        self.assertEqual(item["status"], "open")
        done = erp_store.complete_work_item(
            self.db, item["item_id"], expected_revision=created["revision"],
            audit_context={"operator_name": "测试", "source": "work_item_complete", "actions": [], "app_version": "test"},
        )
        self.assertEqual(done["work_item"]["status"], "done")
        self.assertTrue(done["work_item"]["completed_at"])
        self.assertEqual(erp_store.list_work_items(self.db, "26BS001", "open"), [])

    def test_quote_to_order_link_is_persisted_during_order_save(self):
        quote = database.save_quotation(
            self.db,
            {
                "quote_no": "Q26-001", "quote_date": "2026-09-27", "customer": "原客户",
                "currency": "RMB", "items": [{"model": "KT1000", "quantity": 1, "unit_price": 100}],
            },
            expected_revision=self.initial["revision"], backup=False,
        )
        data = database.load_database(self.db)
        data["合同订单"][0]["来源报价ID"] = quote["quotation"]["quote_id"]
        data["合同订单"][0]["来源报价单号"] = quote["quotation"]["quote_no"]
        saved = database.save_database(
            self.db, data, expected_revision=quote["revision"], backup=False
        )
        linked = database.get_quotation(self.db, quote["quotation"]["quote_id"])
        self.assertEqual(linked["source_job_no"], "")
        self.assertEqual(linked["converted_job_no"], "26BS001")
        self.assertEqual(linked["status"], "已接受")
        with self.assertRaises(database.QuotationValidationError):
            database.delete_quotation(
                self.db, quote["quotation"]["quote_id"],
                expected_revision=saved["revision"], backup=False,
            )

        unlinked_data = database.load_database(self.db)
        unlinked_data["合同订单"][0]["来源报价ID"] = ""
        unlinked_data["合同订单"][0]["来源报价单号"] = ""
        unlinked = database.save_database(
            self.db, unlinked_data, expected_revision=saved["revision"], backup=False
        )
        quote_after_unlink = database.get_quotation(
            self.db, quote["quotation"]["quote_id"]
        )
        self.assertEqual(quote_after_unlink["converted_job_no"], "")
        self.assertGreater(unlinked["revision"], saved["revision"])
        self.assertEqual(saved["data"]["合同订单"][0]["来源报价单号"], "Q26-001")


    def test_contact_and_structured_payment_template_round_trip(self):
        contact = erp_store.save_master_record(
            self.db, "contact",
            {"customer_name": "原客户", "name": "王工", "email": "wang@example.com", "is_primary": 1},
            expected_revision=self.initial["revision"], backup=False,
            audit_context={"operator_name": "测试", "source": "master_data_save", "actions": [], "app_version": "test"},
        )
        template = erp_store.save_master_record(
            self.db, "payment_template",
            {"name": "30-40-20-10", "items": [
                {"kind": "预付款", "ratio": 30, "days": 0, "trigger": "合同生效后", "description_zh": "预付款30%"},
                {"kind": "到货款", "ratio": 40, "days": 30, "trigger": "货到签收后", "description_zh": "到货款40%"},
            ]},
            expected_revision=contact["revision"], backup=False,
            audit_context={"operator_name": "测试", "source": "master_data_save", "actions": [], "app_version": "test"},
        )
        master = erp_store.list_master_data(self.db)
        self.assertEqual(master["contacts"][0]["name"], "王工")
        saved_template = next(x for x in master["payment_templates"] if x["name"] == "30-40-20-10")
        self.assertEqual(saved_template["items"][0]["kind"], "预付款")
        self.assertEqual(saved_template["items"][1]["ratio"], 40)
        self.assertGreater(template["revision"], contact["revision"])


if __name__ == "__main__":
    unittest.main()
