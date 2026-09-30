import tempfile
import sqlite3
from contextlib import closing
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

    def test_master_and_work_item_writes_reject_untracked_business_edit(self):
        with closing(sqlite3.connect(self.db)) as conn:
            conn.execute('UPDATE "设备台账" SET "未税单价"=999')
            conn.commit()
        with self.assertRaises(database.DatabaseValidationError):
            erp_store.save_master_record(self.db, "customer", {"name": "客户"}, backup=False)
        with self.assertRaises(database.DatabaseValidationError):
            erp_store.save_work_item(self.db, {"job_no": "26BS001", "title": "跟进"}, backup=False)
        self.assertEqual(erp_store.list_work_items(self.db), [])
        self.assertFalse(database.verify_audit_chain(self.db)["ok"])

    def test_order_rename_updates_work_item_and_audit_atomically(self):
        erp_store.save_work_item(self.db, {"job_no": "26BS001", "title": "确认搬入"}, backup=False)
        data = database.load_database(self.db)
        for table in database.core.TABLES:
            for record in data[table]:
                record["JOB No"] = "26BS099"
        database.save_database(self.db, data, backup=False)
        self.assertEqual(erp_store.list_work_items(self.db, "26BS001"), [])
        self.assertEqual(len(erp_store.list_work_items(self.db, "26BS099")), 1)
        contract = next(row for row in database.load_database(self.db)["合同订单"]
                        if row["JOB No"] == "26BS099")
        self.assertEqual(erp_store.list_work_items(self.db, "26BS099")[0]["order_id"],
                         contract["记录ID"])
        changes = database.get_audit_events(self.db)["events"][0]["changes"]
        self.assertTrue(any(x["table_name"] == "待办事项" and x["old_value"] == "26BS001"
                            and x["new_value"] == "26BS099" for x in changes))

    def test_deleting_order_preserves_task_without_dangling_job(self):
        erp_store.save_work_item(self.db, {"job_no": "26BS001", "title": "确认搬入"}, backup=False)
        database.save_database(self.db, database.core.empty_data(), backup=False)
        tasks = erp_store.list_work_items(self.db)
        self.assertEqual(len(tasks), 1)
        self.assertEqual(tasks[0]["job_no"], "")
        self.assertIsNone(tasks[0]["order_id"])
        self.assertIn("26BS001", tasks[0]["note"])

    def test_replacing_order_with_same_job_does_not_reassign_old_task(self):
        erp_store.save_work_item(self.db, {"job_no": "26BS001", "title": "旧订单跟进"}, backup=False)
        data = database.load_database(self.db)
        data["合同订单"][0]["记录ID"] = "replacement-contract"
        database.save_database(self.db, data, backup=False)
        task = erp_store.list_work_items(self.db)[0]
        self.assertEqual(task["job_no"], "")
        self.assertIsNone(task["order_id"])
        self.assertEqual(task["title"], "旧订单跟进")
        self.assertIn("原关联订单 26BS001 已删除", task["note"])
        self.assertTrue(database.verify_audit_chain(self.db)["ok"])

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

    def test_work_item_has_stable_foreign_key_to_order_record(self):
        created = erp_store.save_work_item(
            self.db, {"job_no": "26BS001", "title": "订单关联校验"}, backup=False
        )
        contract = next(
            row for row in database.load_database(self.db)["合同订单"]
            if row["JOB No"] == "26BS001"
        )
        self.assertEqual(created["work_item"]["order_id"], contract["记录ID"])
        with closing(sqlite3.connect(self.db)) as conn:
            conn.row_factory = sqlite3.Row
            conn.execute("PRAGMA foreign_keys=ON")
            foreign_keys = conn.execute('PRAGMA foreign_key_list("work_item")').fetchall()
            self.assertTrue(any(
                row["table"] == "合同订单" and row["from"] == "order_id"
                and row["to"] == "记录ID" and row["on_delete"] == "SET NULL"
                for row in foreign_keys
            ))
            with self.assertRaises(sqlite3.IntegrityError):
                conn.execute(
                    'UPDATE "work_item" SET order_id=? WHERE item_id=?',
                    ("missing-order", created["work_item"]["item_id"]),
                )

    def test_audit_verification_detects_missing_work_item_order_link(self):
        created = erp_store.save_work_item(
            self.db, {"job_no": "26BS001", "title": "订单关联篡改校验"}, backup=False
        )
        with closing(sqlite3.connect(self.db)) as conn:
            conn.execute(
                'UPDATE "work_item" SET order_id=NULL WHERE item_id=?',
                (created["work_item"]["item_id"],),
            )
            conn.commit()
        self.assertFalse(database.verify_audit_chain(self.db)["ok"])
        with self.assertRaises(database.DatabaseValidationError):
            erp_store.save_work_item(
                self.db, {"job_no": "26BS001", "title": "后续写入应被拒绝"}, backup=False
            )

    def test_schema_v9_work_item_link_migration_preserves_audit_chain(self):
        created = erp_store.save_work_item(
            self.db, {"job_no": "26BS001", "title": "迁移关联校验"}, backup=False
        )
        columns = (
            "item_id, job_no, type, title, due_date, status, priority, source_type, "
            "source_id, note, created_at, completed_at, updated_at"
        )
        with closing(sqlite3.connect(self.db)) as conn:
            conn.execute("PRAGMA foreign_keys=OFF")
            conn.execute('ALTER TABLE "work_item" RENAME TO "work_item_v10"')
            conn.execute(
                'CREATE TABLE "work_item" ('
                '"item_id" TEXT PRIMARY KEY, "job_no" TEXT NOT NULL DEFAULT \'\', '
                '"type" TEXT NOT NULL DEFAULT \'follow_up\', "title" TEXT NOT NULL, '
                '"due_date" TEXT NOT NULL DEFAULT \'\', "status" TEXT NOT NULL DEFAULT \'open\', '
                '"priority" TEXT NOT NULL DEFAULT \'medium\', "source_type" TEXT NOT NULL DEFAULT \'\', '
                '"source_id" TEXT NOT NULL DEFAULT \'\', "note" TEXT NOT NULL DEFAULT \'\', '
                '"created_at" TEXT NOT NULL, "completed_at" TEXT NOT NULL DEFAULT \'\', '
                '"updated_at" TEXT NOT NULL)'
            )
            conn.execute(
                f'INSERT INTO "work_item" ({columns}) SELECT {columns} FROM "work_item_v10"'
            )
            conn.execute('DROP TABLE "work_item_v10"')
            conn.execute("PRAGMA user_version=9")
            conn.execute("UPDATE _meta SET value='9' WHERE key='schema_version'")
            conn.commit()

        self.assertTrue(database.verify_audit_chain(self.db)["ok"])
        migrated = erp_store.list_work_items(self.db)[0]
        contract = next(
            row for row in database.load_database(self.db)["合同订单"]
            if row["JOB No"] == "26BS001"
        )
        self.assertEqual(migrated["order_id"], contract["记录ID"])
        self.assertTrue(database.verify_audit_chain(self.db)["ok"])

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
