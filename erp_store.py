#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Persistence helpers for ERP V2 master data and work items."""
from __future__ import annotations

import hashlib
import json
import os
import uuid
from datetime import datetime
from typing import Any

import database


MASTER_DEFS = {
    "customer": {
        "table": "customer_master", "id": "customer_id", "prefix": "customer",
        "required": ("name",),
        "fields": ("name", "name_en", "currency", "salesperson", "address", "ship_to", "incoterm", "active"),
        "label": "客户主数据",
    },
    "contact": {
        "table": "customer_contact", "id": "contact_id", "prefix": "contact",
        "required": ("name",),
        "fields": ("customer_id", "customer_name", "name", "title", "email", "phone", "is_primary"),
        "label": "客户联系人",
    },
    "model": {
        "table": "product_model", "id": "model_id", "prefix": "model",
        "required": ("model",),
        "fields": ("model", "description", "unit", "active"),
        "label": "设备型号主数据",
    },
    "payment_template": {
        "table": "payment_term_template", "id": "template_id", "prefix": "payment",
        "required": ("name",),
        "fields": ("name", "currency", "description", "items_json", "active"),
        "label": "付款条件模板",
    },
}


def _text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def _stable_id(prefix: str, value: str) -> str:
    digest = hashlib.sha1(value.encode("utf-8")).hexdigest()[:20]
    return f"{prefix}-{digest}"


def _revision_guard(path, expected_revision):
    current = database.get_revision(path) if os.path.exists(path) else 0
    if expected_revision is not None and current != expected_revision:
        raise database.StaleImportError(
            f"数据库已从修订 {expected_revision} 更新到 {current}，请重新加载后再操作"
        )
    return current


def _audit_change(table_name, record_id, job_no, old, new):
    if old is None:
        operation = "insert"
    elif new is None:
        operation = "delete"
    else:
        operation = "update"
    return {
        "table_name": table_name,
        "record_id": record_id,
        "job_no": job_no or "",
        "operation": operation,
        "field_name": "",
        "old_value": old,
        "new_value": new,
        "origin": "user",
    }


def _finish(conn, current_revision, reason, summary, changes, audit_context):
    revision = current_revision + 1
    now = datetime.now().isoformat(timespec="seconds")
    database._set_meta(conn, "revision", revision)
    database._set_meta(conn, "last_saved_at", now)
    conn.execute(
        "INSERT INTO _change_log(revision, created_at, reason, summary_json) VALUES(?, ?, ?, ?)",
        (revision, now, reason, json.dumps(summary, ensure_ascii=False)),
    )
    audit_event_id = None
    if changes:
        audit_event_id, _ = database._append_audit_event(
            conn, revision, now, reason, changes, audit_context
        )
    conn.commit()
    return {
        "ok": True, "revision": revision,
        "audit_event_id": audit_event_id,
        "audit_change_count": len(changes) if audit_event_id else 0,
    }


def _row_dict(row):
    return dict(row) if row is not None else None


def list_master_data(path):
    """Return persisted master data plus inferred legacy values without rewriting history."""
    conn = database._connect(path)
    try:
        customers = [dict(r) for r in conn.execute(
            'SELECT * FROM "customer_master" ORDER BY active DESC, name'
        )]
        models = [dict(r) for r in conn.execute(
            'SELECT * FROM "product_model" ORDER BY active DESC, model'
        )]
        contacts = [dict(r) for r in conn.execute(
            'SELECT * FROM "customer_contact" ORDER BY customer_name, is_primary DESC, name'
        )]
        templates = [dict(r) for r in conn.execute(
            'SELECT * FROM "payment_term_template" ORDER BY active DESC, name'
        )]

        customer_names = {r["name"].casefold() for r in customers}
        inferred_customers = []
        sql = '''
            SELECT DISTINCT TRIM(name) AS name FROM (
              SELECT "客户" AS name FROM "合同订单"
              UNION ALL SELECT customer AS name FROM quotation
            ) WHERE TRIM(COALESCE(name,'')) <> '' ORDER BY name
        '''
        for row in conn.execute(sql):
            name = row["name"]
            if name.casefold() in customer_names:
                continue
            inferred_customers.append({
                "customer_id": _stable_id("customer", name.casefold()), "name": name,
                "name_en": "", "currency": "RMB", "salesperson": "", "address": "",
                "ship_to": "", "incoterm": "", "active": 1, "updated_at": "",
                "inferred": True,
            })

        model_names = {r["model"].casefold() for r in models}
        inferred_models = []
        sql = '''
            SELECT DISTINCT TRIM(model) AS model FROM (
              SELECT "设备型号" AS model FROM "设备台账"
              UNION ALL SELECT model FROM quotation_item
            ) WHERE TRIM(COALESCE(model,'')) <> '' ORDER BY model
        '''
        for row in conn.execute(sql):
            model = row["model"]
            if model.casefold() in model_names:
                continue
            inferred_models.append({
                "model_id": _stable_id("model", model.casefold()), "model": model,
                "description": "", "unit": "台", "active": 1, "updated_at": "",
                "inferred": True,
            })

        for row in templates:
            try:
                row["items"] = json.loads(row.get("items_json") or "[]")
            except (TypeError, ValueError):
                row["items"] = []
        return {
            "customers": customers + inferred_customers,
            "contacts": contacts,
            "models": models + inferred_models,
            "payment_templates": templates,
        }
    finally:
        conn.close()


def save_master_record(path, kind, record, *, expected_revision=None,
                       audit_context=None, backup=True):
    spec = MASTER_DEFS.get(_text(kind))
    if not spec:
        raise ValueError(f"未知主数据类型：{kind}")
    if not isinstance(record, dict):
        raise ValueError("主数据格式无效")
    for field in spec["required"]:
        if not _text(record.get(field)):
            raise ValueError(f"{field} 不能为空")

    _revision_guard(path, expected_revision)
    backup_path = database._backup_database(path) if backup else None
    conn = database._connect(path)
    try:
        conn.execute("BEGIN IMMEDIATE")
        current = int(database._meta(conn, "revision", "0") or 0)
        if expected_revision is not None and current != expected_revision:
            raise database.StaleImportError(
                f"数据库已从修订 {expected_revision} 更新到 {current}，请重新加载后再操作"
            )

        pk = spec["id"]
        rid = _text(record.get(pk))
        natural = _text(record.get(spec["required"][0]))
        if not rid:
            rid = _stable_id(spec["prefix"], natural.casefold()) if natural else f"{spec['prefix']}-{uuid.uuid4().hex}"
        old = _row_dict(conn.execute(
            f'SELECT * FROM "{spec["table"]}" WHERE "{pk}"=?', (rid,)
        ).fetchone())

        now = datetime.now().isoformat(timespec="seconds")
        values = {field: ("" if record.get(field) is None else record.get(field)) for field in spec["fields"]}
        if kind == "customer" and not _text(values.get("currency")):
            values["currency"] = "RMB"
        if kind == "model" and not _text(values.get("unit")):
            values["unit"] = "台"
        if "active" in values:
            values["active"] = 0 if record.get("active") in (False, 0, "0") else 1
        if "is_primary" in values:
            values["is_primary"] = 1 if record.get("is_primary") in (True, 1, "1") else 0
        if "items_json" in values:
            raw_items = record.get("items", record.get("items_json", []))
            if isinstance(raw_items, str):
                try:
                    raw_items = json.loads(raw_items)
                except ValueError:
                    raw_items = []
            values["items_json"] = json.dumps(raw_items or [], ensure_ascii=False)

        columns = [pk, *spec["fields"], "updated_at"]
        payload = {pk: rid, **values, "updated_at": now}
        updates = [x for x in columns if x != pk]
        conn.execute(
            f'INSERT INTO "{spec["table"]}" ({", ".join(database._q(x) for x in columns)}) '
            f'VALUES ({", ".join("?" for _ in columns)}) '
            f'ON CONFLICT("{pk}") DO UPDATE SET '
            + ", ".join(f'{database._q(x)}=excluded.{database._q(x)}' for x in updates),
            [payload.get(x, "") for x in columns],
        )
        saved = _row_dict(conn.execute(
            f'SELECT * FROM "{spec["table"]}" WHERE "{pk}"=?', (rid,)
        ).fetchone())
        change = _audit_change(spec["label"], rid, "", old, saved)
        result = _finish(
            conn, current, "master_data_save", {spec["label"]: 1},
            [change] if old != saved else [], audit_context,
        )
        result.update({"record": saved, "backup": backup_path})
        return result
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def list_work_items(path, job_no="", status=""):
    conn = database._connect(path)
    try:
        where, params = [], []
        if _text(job_no):
            where.append("job_no=?"); params.append(_text(job_no))
        if _text(status):
            where.append("status=?"); params.append(_text(status))
        clause = (" WHERE " + " AND ".join(where)) if where else ""
        return [dict(r) for r in conn.execute(
            'SELECT * FROM "work_item"' + clause +
            " ORDER BY CASE priority WHEN 'high' THEN 0 WHEN 'medium' THEN 1 ELSE 2 END, "
            "CASE WHEN due_date='' THEN 1 ELSE 0 END, due_date, updated_at DESC",
            params,
        )]
    finally:
        conn.close()


def save_work_item(path, item, *, expected_revision=None, audit_context=None, backup=True):
    if not isinstance(item, dict) or not _text(item.get("title")):
        raise ValueError("待办标题不能为空")
    _revision_guard(path, expected_revision)
    backup_path = database._backup_database(path) if backup else None
    conn = database._connect(path)
    try:
        conn.execute("BEGIN IMMEDIATE")
        current = int(database._meta(conn, "revision", "0") or 0)
        if expected_revision is not None and current != expected_revision:
            raise database.StaleImportError(
                f"数据库已从修订 {expected_revision} 更新到 {current}，请重新加载后再操作"
            )
        job_no = _text(item.get("job_no"))
        if job_no and not conn.execute(
            'SELECT 1 FROM "合同订单" WHERE "JOB No"=?', (job_no,)
        ).fetchone():
            raise ValueError(f"JOB No 不存在：{job_no}")
        item_id = _text(item.get("item_id")) or f"work-{uuid.uuid4().hex}"
        old = _row_dict(conn.execute(
            'SELECT * FROM "work_item" WHERE item_id=?', (item_id,)
        ).fetchone())
        now = datetime.now().isoformat(timespec="seconds")
        created_at = _text(old.get("created_at")) if old else now
        status = _text(item.get("status")) or "open"
        completed_at = _text(item.get("completed_at"))
        if status == "done" and not completed_at:
            completed_at = now
        if status != "done":
            completed_at = ""
        payload = {
            "item_id": item_id, "job_no": job_no,
            "type": _text(item.get("type")) or "follow_up",
            "title": _text(item.get("title")),
            "due_date": _text(item.get("due_date")),
            "status": status,
            "priority": _text(item.get("priority")) or "medium",
            "source_type": _text(item.get("source_type")),
            "source_id": _text(item.get("source_id")),
            "note": _text(item.get("note")),
            "created_at": created_at, "completed_at": completed_at, "updated_at": now,
        }
        cols = tuple(payload)
        conn.execute(
            f'INSERT INTO "work_item" ({", ".join(database._q(x) for x in cols)}) '
            f'VALUES ({", ".join("?" for _ in cols)}) '
            'ON CONFLICT(item_id) DO UPDATE SET ' +
            ", ".join(f'{database._q(x)}=excluded.{database._q(x)}' for x in cols if x != "item_id"),
            [payload[x] for x in cols],
        )
        saved = _row_dict(conn.execute(
            'SELECT * FROM "work_item" WHERE item_id=?', (item_id,)
        ).fetchone())
        change = _audit_change("待办", item_id, job_no, old, saved)
        result = _finish(
            conn, current, "work_item_save", {"待办": 1},
            [change] if old != saved else [], audit_context,
        )
        result.update({"work_item": saved, "backup": backup_path})
        return result
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def complete_work_item(path, item_id, *, expected_revision=None, audit_context=None):
    conn = database._connect(path)
    try:
        row = conn.execute('SELECT * FROM "work_item" WHERE item_id=?', (_text(item_id),)).fetchone()
        if row is None:
            raise KeyError(f"待办不存在：{item_id}")
        item = dict(row)
    finally:
        conn.close()
    item["status"] = "done"
    return save_work_item(
        path, item, expected_revision=expected_revision,
        audit_context=audit_context, backup=True,
    )
