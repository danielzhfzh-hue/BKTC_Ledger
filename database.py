#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""SQLite persistence and safe round-trip Excel editing for BKTC Ledger."""

from __future__ import annotations

import hashlib
import json
import os
import re
import sqlite3
import uuid
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

import core


DATABASE_VERSION = 2
WORKBOOK_VERSION = 2
META_SHEET = "_BKTC_META"
HELP_SHEET = "使用说明"
ID_HEADER = "_记录ID"
DATABASE_EXTENSIONS = {".db", ".sqlite", ".sqlite3"}


class DatabaseValidationError(ValueError):
    def __init__(self, issues):
        self.issues = issues
        errors = [x["msg"] for x in issues if x.get("severity") == "error"]
        detail = "；".join(errors[:5])
        if len(errors) > 5:
            detail += f"；另有 {len(errors) - 5} 项"
        super().__init__("数据校验失败，未写入数据库：" + (detail or "存在未知错误"))


class SpreadsheetFormatError(ValueError):
    pass


class StaleImportError(RuntimeError):
    pass


@dataclass
class PreparedImport:
    database_path: str
    workbook_path: str
    database_id: str
    database_hash: str
    base_revision: int
    current_revision: int
    data: dict[str, list[dict[str, Any]]]
    rules: dict[str, Any]
    changes: list[dict[str, Any]]
    action_counts: dict[str, int]
    table_counts: dict[str, dict[str, int]]
    issues: list[dict[str, str]]
    file_signature: tuple[int, int]

    @property
    def stale(self):
        return self.base_revision != self.current_revision


def is_database_path(path):
    return Path(os.fspath(path)).suffix.lower() in DATABASE_EXTENSIONS


def database_path_for(path):
    """Return a sibling .db path for a legacy JSON path."""
    p = Path(os.fspath(path))
    if is_database_path(p):
        return str(p)
    name = p.name
    if name.endswith(".records.json"):
        name = name[:-len(".records.json")] + ".db"
    elif p.suffix.lower() == ".json":
        name = p.stem + ".db"
    else:
        name += ".db"
    return str(p.with_name(name))


def _q(name):
    return '"' + str(name).replace('"', '""') + '"'


def _sql_type(field_type):
    if field_type in (core.NUMBER,):
        return "REAL"
    if field_type in (core.INT,):
        return "INTEGER"
    return "TEXT"


def _connect(path):
    path = os.path.abspath(os.fspath(path))
    os.makedirs(os.path.dirname(path), exist_ok=True)
    conn = sqlite3.connect(path, timeout=15)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA journal_mode = WAL")
    conn.execute("PRAGMA synchronous = FULL")
    conn.execute("PRAGMA busy_timeout = 15000")
    _create_schema(conn)
    return conn


def _create_schema(conn):
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS "_meta" (
            "key" TEXT PRIMARY KEY,
            "value" TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS "_warning_rules" (
            "key" TEXT PRIMARY KEY,
            "value_json" TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS "_change_log" (
            "revision" INTEGER PRIMARY KEY,
            "created_at" TEXT NOT NULL,
            "reason" TEXT NOT NULL,
            "summary_json" TEXT NOT NULL
        );
        """
    )
    for table in core.TABLES:
        columns = [f'{_q("记录ID")} TEXT PRIMARY KEY']
        for field, field_type, *_ in core.SCHEMA[table]:
            constraints = []
            if field == "JOB No":
                constraints.append("NOT NULL")
                if table == "合同订单":
                    constraints.append("UNIQUE")
            if table == "发货批次" and field == "发货批次":
                constraints.append("NOT NULL")
            columns.append(
                f"{_q(field)} {_sql_type(field_type)} {' '.join(constraints)}".rstrip()
            )
        columns.extend([
            f'{_q("_extra_json")} TEXT NOT NULL DEFAULT "{{}}"',
            f'{_q("_row_order")} INTEGER NOT NULL',
        ])
        if table != "合同订单":
            columns.append(
                f'FOREIGN KEY ({_q("JOB No")}) REFERENCES {_q("合同订单")} '
                f'({_q("JOB No")}) ON UPDATE CASCADE ON DELETE CASCADE '
                "DEFERRABLE INITIALLY DEFERRED"
            )
        if table == "发货批次":
            columns.append(f'UNIQUE ({_q("JOB No")}, {_q("发货批次")})')
        if table == "设备台账":
            columns.append(
                f'FOREIGN KEY ({_q("JOB No")}, {_q("发货批次")}) '
                f'REFERENCES {_q("发货批次")} ({_q("JOB No")}, {_q("发货批次")}) '
                "ON UPDATE CASCADE ON DELETE RESTRICT DEFERRABLE INITIALLY DEFERRED"
            )
        conn.execute(f'CREATE TABLE IF NOT EXISTS {_q(table)} ({", ".join(columns)})')
        existing_columns = {
            row[1] for row in conn.execute(f'PRAGMA table_info({_q(table)})')
        }
        for field, field_type, *_ in core.SCHEMA[table]:
            if field not in existing_columns:
                conn.execute(
                    f'ALTER TABLE {_q(table)} ADD COLUMN {_q(field)} {_sql_type(field_type)}'
                )
        if table != "合同订单":
            conn.execute(
                f'CREATE INDEX IF NOT EXISTS {_q("idx_" + table + "_job")} '
                f'ON {_q(table)} ({_q("JOB No")})'
            )
    conn.execute(
        f'CREATE INDEX IF NOT EXISTS {_q("idx_设备台账_batch")} '
        f'ON {_q("设备台账")} ({_q("JOB No")}, {_q("发货批次")})'
    )
    conn.execute(
        "INSERT OR IGNORE INTO _meta(key, value) VALUES('database_id', ?)",
        (str(uuid.uuid4()),),
    )
    conn.execute("INSERT OR IGNORE INTO _meta(key, value) VALUES('revision', '0')")
    conn.execute(
        "INSERT INTO _meta(key, value) VALUES('schema_version', ?) "
        "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
        (str(DATABASE_VERSION),),
    )
    conn.execute(f"PRAGMA user_version = {DATABASE_VERSION}")
    conn.commit()


def _meta(conn, key, default=""):
    row = conn.execute("SELECT value FROM _meta WHERE key = ?", (key,)).fetchone()
    return row[0] if row else default


def _set_meta(conn, key, value):
    conn.execute(
        "INSERT INTO _meta(key, value) VALUES(?, ?) "
        "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
        (key, str(value)),
    )


def _new_record_id():
    return "local-" + uuid.uuid4().hex


def _prepare_data(data):
    prepared = core.derive(data)
    for table in core.TABLES:
        seen = set()
        for rec in prepared[table]:
            rid = str(rec.get("记录ID") or "").strip()
            if not rid or rid in seen:
                rid = _new_record_id()
                rec["记录ID"] = rid
            seen.add(rid)
    issues = core.validate(prepared)
    if any(x.get("severity") == "error" for x in issues):
        raise DatabaseValidationError(issues)
    return prepared, issues


def _backup_database(path, keep=10):
    path = os.path.abspath(os.fspath(path))
    if not os.path.exists(path) or os.path.getsize(path) == 0:
        return None
    backup_dir = os.path.join(os.path.dirname(path), "备份")
    os.makedirs(backup_dir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(path))[0]
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]
    target = os.path.join(backup_dir, f"{stem}_database_{stamp}.db")
    source = _connect(path)
    destination = sqlite3.connect(target)
    try:
        source.backup(destination)
    finally:
        destination.close()
        source.close()
    backups = sorted(
        (os.path.join(backup_dir, f) for f in os.listdir(backup_dir)
         if f.startswith(stem + "_database_") and f.endswith(".db")),
        key=os.path.getmtime,
    )
    while len(backups) > keep:
        os.remove(backups.pop(0))
    return target


def _insert_rows(conn, table, records):
    fields = [f[0] for f in core.SCHEMA[table]]
    known = set(fields) | {"记录ID"}
    columns = ["记录ID", *fields, "_extra_json", "_row_order"]
    sql = (
        f'INSERT INTO {_q(table)} ({", ".join(_q(x) for x in columns)}) '
        f'VALUES ({", ".join("?" for _ in columns)})'
    )
    for order, rec in enumerate(records):
        extra = {k: v for k, v in rec.items() if k not in known}
        values = [rec["记录ID"]]
        for field in fields:
            value = rec.get(field)
            if table == "设备台账" and field == "发货批次" and value == "":
                value = None
            values.append(value)
        values.extend([json.dumps(extra, ensure_ascii=False, separators=(",", ":")), order])
        conn.execute(sql, values)


def _rules_from_connection(conn):
    rules = {}
    for row in conn.execute("SELECT key, value_json FROM _warning_rules"):
        try:
            rules[row[0]] = json.loads(row[1])
        except (TypeError, ValueError):
            continue
    return {**core.DEFAULT_RULES, **rules}


def save_database(path, data, rules=None, *, reason="manual_save", backup=True,
                  expected_revision=None):
    path = os.path.abspath(os.fspath(path))
    prepared, issues = _prepare_data(data)
    if rules is None:
        rules = get_rules(path) if os.path.exists(path) else dict(core.DEFAULT_RULES)
    early_revision = get_revision(path) if os.path.exists(path) else 0
    if expected_revision is not None and early_revision != expected_revision:
        raise StaleImportError(
            f"数据库已从修订 {expected_revision} 更新到 {early_revision}，请重新加载后再操作"
        )
    backup_path = _backup_database(path) if backup else None
    conn = _connect(path)
    try:
        conn.execute("BEGIN IMMEDIATE")
        current_revision = int(_meta(conn, "revision", "0") or 0)
        if expected_revision is not None and current_revision != expected_revision:
            raise StaleImportError(
                f"数据库已从修订 {expected_revision} 更新到 {current_revision}，请重新加载后再操作"
            )
        for table in ("回款记录", "开票记录", "设备台账", "付款条件", "发货批次", "合同订单"):
            conn.execute(f'DELETE FROM {_q(table)}')
        for table in ("合同订单", "发货批次", "付款条件", "设备台账", "开票记录", "回款记录"):
            _insert_rows(conn, table, prepared[table])
        conn.execute("DELETE FROM _warning_rules")
        for key, value in (rules or {}).items():
            conn.execute(
                "INSERT INTO _warning_rules(key, value_json) VALUES(?, ?)",
                (key, json.dumps(value, ensure_ascii=False)),
            )
        foreign_key_issues = conn.execute("PRAGMA foreign_key_check").fetchall()
        if foreign_key_issues:
            raise sqlite3.IntegrityError(f"外键检查失败：{foreign_key_issues[:3]}")
        revision = current_revision + 1
        now = datetime.now().isoformat(timespec="seconds")
        _set_meta(conn, "revision", revision)
        _set_meta(conn, "last_saved_at", now)
        summary = {table: len(prepared[table]) for table in core.TABLES}
        conn.execute(
            "INSERT INTO _change_log(revision, created_at, reason, summary_json) "
            "VALUES(?, ?, ?, ?)",
            (revision, now, reason, json.dumps(summary, ensure_ascii=False)),
        )
        conn.commit()
    except sqlite3.IntegrityError as exc:
        conn.rollback()
        raise DatabaseValidationError([
            {"severity": "error", "msg": f"关系约束检查失败：{exc}"}
        ]) from exc
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
    return {
        "ok": True,
        "path": path,
        "revision": revision,
        "backup": backup_path,
        "data": load_database(path),
        "issues": issues,
    }


def load_database(path):
    path = os.path.abspath(os.fspath(path))
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    conn = _connect(path)
    try:
        integrity = conn.execute("PRAGMA integrity_check").fetchone()[0]
        if integrity != "ok":
            raise RuntimeError(f"数据库完整性检查失败：{integrity}")
        data = core.empty_data()
        for table in core.TABLES:
            fields = [f[0] for f in core.SCHEMA[table]]
            select = ["记录ID", *fields, "_extra_json"]
            rows = conn.execute(
                f'SELECT {", ".join(_q(x) for x in select)} FROM {_q(table)} '
                f'ORDER BY {_q("_row_order")}'
            )
            for row in rows:
                try:
                    rec = json.loads(row["_extra_json"] or "{}")
                except (TypeError, ValueError):
                    rec = {}
                rec["记录ID"] = row["记录ID"]
                for field in fields:
                    rec[field] = row[field]
                data[table].append(rec)
        return core.normalize(data)
    finally:
        conn.close()


def get_rules(path):
    if not os.path.exists(path):
        return dict(core.DEFAULT_RULES)
    conn = _connect(path)
    try:
        return _rules_from_connection(conn)
    finally:
        conn.close()


def get_revision(path):
    if not os.path.exists(path):
        return 0
    conn = _connect(path)
    try:
        return int(_meta(conn, "revision", "0") or 0)
    finally:
        conn.close()


def get_database_info(path):
    conn = _connect(path)
    try:
        return {
            "path": os.path.abspath(os.fspath(path)),
            "database_id": _meta(conn, "database_id"),
            "revision": int(_meta(conn, "revision", "0") or 0),
            "last_saved_at": _meta(conn, "last_saved_at"),
            "integrity": conn.execute("PRAGMA integrity_check").fetchone()[0],
        }
    finally:
        conn.close()


def migrate_json_to_database(json_path, database_path):
    json_path = os.path.abspath(os.fspath(json_path))
    database_path = os.path.abspath(os.fspath(database_path))
    if os.path.exists(database_path):
        return {
            "ok": True,
            "path": database_path,
            "revision": get_revision(database_path),
            "data": load_database(database_path),
            "migrated": False,
        }
    data = core.load_store(json_path)
    rules = core.get_rules(json_path)
    result = save_database(
        database_path, data, rules, reason="json_migration", backup=False
    )
    result["migrated"] = True
    result["source"] = json_path
    return result


def _canonical_hash(data):
    payload = json.dumps(data, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def _safe_excel_text(value):
    text = str(value)
    if text.startswith(("=", "+", "-", "@")):
        return "'" + text
    return text


def _excel_value(value, field_type):
    if value is None or value == "":
        return None
    if field_type == core.DATE:
        try:
            return datetime.strptime(str(value)[:10], "%Y-%m-%d").date()
        except ValueError:
            return _safe_excel_text(value)
    if field_type == core.NUMBER:
        return float(value)
    if field_type == core.INT:
        return int(float(value))
    return _safe_excel_text(value)


def export_editable_workbook(database_path, workbook_path):
    database_path = os.path.abspath(os.fspath(database_path))
    workbook_path = os.path.abspath(os.fspath(workbook_path))
    data = load_database(database_path)
    info = get_database_info(database_path)
    wb = Workbook()
    help_ws = wb.active
    help_ws.title = HELP_SHEET
    help_ws.sheet_view.showGridLines = False
    help_ws["A1"] = "BKTC 订单数据 · 外部编辑副本"
    help_ws["A1"].font = Font(name="Noto Sans SC", size=20, bold=True, color="004494")
    instructions = [
        "使用方法",
        "1. 在六个业务工作表中修改数据；可新增整行，也可删除整行。",
        "2. 不要重命名工作表或表头；隐藏的 _记录ID 列用于可靠识别原记录。",
        "3. 自动计算字段未导出，重新导入后由 BKTC Ledger 统一计算。",
        "4. 回到应用点击“导入变更”，先核对差异，再勾选二次确认后写入。",
        "5. 若数据库在导出后已被保存过，旧副本会被拒绝，请重新导出，避免覆盖新数据。",
        "安全保障：导入使用事务、外键校验和写入前数据库备份。",
        f"数据库：{os.path.basename(database_path)}",
        f"导出修订：{info['revision']}",
        f"导出时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
    ]
    for row, text in enumerate(instructions, 3):
        help_ws.cell(row, 1, text)
        help_ws.cell(row, 1).font = Font(
            name="Noto Sans SC", size=11, bold=(row == 3), color="1C1C1C"
        )
        help_ws.cell(row, 1).alignment = Alignment(wrap_text=True, vertical="top")
    help_ws.column_dimensions["A"].width = 96
    help_ws.row_dimensions[1].height = 34

    header_fill = PatternFill("solid", fgColor="004494")
    accent_fill = PatternFill("solid", fgColor="13BCBC")
    stripe_fill = PatternFill("solid", fgColor="F4F8FB")
    for table_index, table in enumerate(core.TABLES):
        ws = wb.create_sheet(table)
        ws.sheet_properties.tabColor = "13BCBC" if table_index % 2 else "004494"
        editable = [f for f in core.SCHEMA[table] if f[0] not in core.DERIVED.get(table, [])]
        headers = [ID_HEADER, *[f[0] for f in editable]]
        ws.append(headers)
        for cell in ws[1]:
            cell.fill = accent_fill if cell.column == 1 else header_fill
            cell.font = Font(name="Noto Sans SC", bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 25
        for rec in data[table]:
            ws.append([
                rec.get("记录ID"),
                *[_excel_value(rec.get(field), field_type)
                  for field, field_type, *_ in editable],
            ])
        ws.column_dimensions["A"].hidden = True
        ws.freeze_panes = "B2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(2, ws.max_row)}"
        for row in range(2, ws.max_row + 1):
            if row % 2 == 0:
                for cell in ws[row]:
                    cell.fill = stripe_fill
            for cell in ws[row]:
                cell.font = Font(name="Noto Sans SC", size=10, color="1C1C1C")
        for col, (field, field_type, *options) in enumerate(editable, 2):
            letter = get_column_letter(col)
            ws.column_dimensions[letter].width = min(34, max(12, len(field) * 2 + 4))
            if field_type == core.DATE:
                for cell in ws[letter][1:]:
                    cell.number_format = "yyyy-mm-dd"
            if field_type in (core.NUMBER, core.INT):
                for cell in ws[letter][1:]:
                    cell.number_format = "0.###"
            choices = options[0] if options else []
            if choices:
                formula = '"' + ",".join(str(x) for x in choices if x != "") + '"'
                if len(formula) <= 255:
                    dv = DataValidation(type="list", formula1=formula, allow_blank=True)
                    dv.error = "请从列表选择，或清空单元格"
                    dv.errorTitle = "值不在允许列表"
                    ws.add_data_validation(dv)
                    dv.add(f"{letter}2:{letter}1048576")
        ws.conditional_formatting.add(
            f"B2:{get_column_letter(len(headers))}{max(2, ws.max_row)}",
            FormulaRule(formula=["MOD(ROW(),2)=0"], fill=stripe_fill),
        )

    meta = wb.create_sheet(META_SHEET)
    metadata = {
        "format_version": WORKBOOK_VERSION,
        "database_id": info["database_id"],
        "base_revision": info["revision"],
        "base_hash": _canonical_hash(data),
        "exported_at": datetime.now().isoformat(timespec="seconds"),
    }
    for key, value in metadata.items():
        meta.append([key, value])
    meta.sheet_state = "veryHidden"
    os.makedirs(os.path.dirname(workbook_path), exist_ok=True)
    wb.save(workbook_path)
    return {
        "ok": True,
        "path": workbook_path,
        "revision": info["revision"],
        "rows": sum(len(data[t]) for t in core.TABLES),
    }


def _workbook_metadata(wb):
    if META_SHEET not in wb.sheetnames:
        raise SpreadsheetFormatError("这不是 BKTC Ledger 导出的可编辑工作簿（缺少元数据）")
    ws = wb[META_SHEET]
    return {str(row[0].value): row[1].value for row in ws.iter_rows(min_col=1, max_col=2)
            if row[0].value not in (None, "")}


def _cell_value(cell, field, field_type, table, row_number):
    if cell.data_type == "f":
        raise SpreadsheetFormatError(
            f"{table} 第 {row_number} 行“{field}”包含公式；导入只接受值，请粘贴为值后重试"
        )
    value = cell.value
    if value is None or value == "":
        return None if field_type in (core.NUMBER, core.INT) else ""
    if field_type == core.DATE:
        if isinstance(value, (datetime, date)):
            return value.strftime("%Y-%m-%d")
        text = str(value).strip()
        normalized = core._norm(text, core.DATE)
        try:
            datetime.strptime(normalized, "%Y-%m-%d")
        except ValueError as exc:
            raise SpreadsheetFormatError(
                f"{table} 第 {row_number} 行“{field}”不是有效日期：{value!r}"
            ) from exc
        return normalized
    if field_type in (core.NUMBER, core.INT):
        try:
            number = float(str(value).replace(",", ""))
        except (TypeError, ValueError) as exc:
            raise SpreadsheetFormatError(
                f"{table} 第 {row_number} 行“{field}”不是数字：{value!r}"
            ) from exc
        return int(round(number)) if field_type == core.INT else number
    text = str(value)
    if len(text) > 1 and text[0] == "'" and text[1] in "=+-@":
        text = text[1:]
    return text


def _record_label(table, rec):
    job = str(rec.get("JOB No") or "（无 JOB）")
    secondaries = {
        "合同订单": "客户",
        "付款条件": "款类",
        "设备台账": "製造番号",
        "发货批次": "发货批次",
        "开票记录": "款类",
        "回款记录": "款类",
    }
    secondary = str(rec.get(secondaries[table]) or "").strip()
    return job + (" / " + secondary if secondary else "")


def _diff_data(before, after):
    changes = []
    action_counts = {"新增": 0, "修改": 0, "删除": 0}
    table_counts = {t: {"新增": 0, "修改": 0, "删除": 0} for t in core.TABLES}
    for table in core.TABLES:
        old = {str(r.get("记录ID")): r for r in before[table]}
        new = {str(r.get("记录ID")): r for r in after[table]}
        field_names = [f[0] for f in core.SCHEMA[table]]
        for rid in new.keys() - old.keys():
            fields = [
                {"field": f, "before": "", "after": new[rid].get(f)}
                for f in field_names if new[rid].get(f) not in (None, "")
            ]
            changes.append({"table": table, "action": "新增", "record_id": rid,
                            "label": _record_label(table, new[rid]), "fields": fields})
            action_counts["新增"] += 1
            table_counts[table]["新增"] += 1
        for rid in old.keys() - new.keys():
            changes.append({"table": table, "action": "删除", "record_id": rid,
                            "label": _record_label(table, old[rid]), "fields": []})
            action_counts["删除"] += 1
            table_counts[table]["删除"] += 1
        for rid in old.keys() & new.keys():
            fields = []
            for field in field_names:
                before_value = old[rid].get(field)
                after_value = new[rid].get(field)
                if before_value != after_value:
                    fields.append({"field": field, "before": before_value,
                                   "after": after_value,
                                   "derived": field in core.DERIVED.get(table, [])})
            if fields:
                changes.append({"table": table, "action": "修改", "record_id": rid,
                                "label": _record_label(table, new[rid]), "fields": fields})
                action_counts["修改"] += 1
                table_counts[table]["修改"] += 1
    order = {"删除": 0, "修改": 1, "新增": 2}
    changes.sort(key=lambda x: (core.TABLES.index(x["table"]), order[x["action"]], x["label"]))
    return changes, action_counts, table_counts


def _replace_coverage_token(value, old, new):
    """替换覆盖串中的完整键，同时兼容调拨番号的箭头写法。"""
    if not old or old == new:
        return value
    changed = False
    result = []
    for item in core.coverage_values(value):
        pieces = re.split(r"([→⇒⟶➡➝])", item)
        for index in range(0, len(pieces), 2):
            if pieces[index] == old:
                pieces[index] = new
                changed = True
        result.append("".join(pieces))
    return ";".join(result) if changed else value


def _propagate_import_relational_edits(current, incoming):
    """用稳定记录 ID 识别 Excel 改名，同步仍保留旧键的关联行。"""
    current_by_id = {
        table: {core.s(rec.get("记录ID")): rec for rec in current[table]
                if core.s(rec.get("记录ID"))}
        for table in core.TABLES
    }
    incoming_by_id = {
        table: {core.s(rec.get("记录ID")): rec for rec in incoming[table]
                if core.s(rec.get("记录ID"))}
        for table in core.TABLES
    }

    job_renames = {}
    for rid, old in current_by_id["合同订单"].items():
        new = incoming_by_id["合同订单"].get(rid)
        if new and core.s(old.get("JOB No")) != core.s(new.get("JOB No")):
            job_renames[core.s(old.get("JOB No"))] = core.s(new.get("JOB No"))
    for table in ("付款条件", "设备台账", "发货批次", "开票记录", "回款记录"):
        for rec in incoming[table]:
            old_job = core.s(rec.get("JOB No"))
            if old_job in job_renames:
                rec["JOB No"] = job_renames[old_job]

    for rid, old in current_by_id["发货批次"].items():
        new = incoming_by_id["发货批次"].get(rid)
        if not new:
            continue
        old_batch = core.s(old.get("发货批次"))
        new_batch = core.s(new.get("发货批次"))
        if not old_batch or old_batch == new_batch:
            continue
        job = core.s(new.get("JOB No"))
        for dev in incoming["设备台账"]:
            if (core.s(dev.get("JOB No")) == job
                    and core.s(dev.get("发货批次")) == old_batch):
                dev["发货批次"] = new_batch
        for table in ("开票记录", "回款记录"):
            for rec in incoming[table]:
                if core.s(rec.get("JOB No")) == job:
                    rec["覆盖批次"] = _replace_coverage_token(
                        rec.get("覆盖批次"), old_batch, new_batch
                    )

    for rid, old in current_by_id["设备台账"].items():
        new = incoming_by_id["设备台账"].get(rid)
        if not new:
            continue
        old_serial = core.s(old.get("製造番号"))
        new_serial = core.s(new.get("製造番号"))
        if not old_serial or old_serial == new_serial:
            continue
        job = core.s(new.get("JOB No"))
        for table in ("开票记录", "回款记录"):
            for rec in incoming[table]:
                if core.s(rec.get("JOB No")) == job:
                    rec["覆盖製造番号"] = _replace_coverage_token(
                        rec.get("覆盖製造番号"), old_serial, new_serial
                    )

    for rid, old in current_by_id["付款条件"].items():
        new = incoming_by_id["付款条件"].get(rid)
        if not new:
            continue
        old_kind = core.s(old.get("款类"))
        new_kind = core.s(new.get("款类"))
        if not old_kind or old_kind == new_kind:
            continue
        job = core.s(new.get("JOB No"))
        if any(rec is not new and core.s(rec.get("JOB No")) == job
               and core.s(rec.get("款类")) == old_kind
               for rec in incoming["付款条件"]):
            continue
        for table in ("开票记录", "回款记录"):
            for rec in incoming[table]:
                if (core.s(rec.get("JOB No")) == job
                        and core.s(rec.get("款类")) == old_kind):
                    rec["款类"] = new_kind

    return incoming


def prepare_editable_import(database_path, workbook_path):
    database_path = os.path.abspath(os.fspath(database_path))
    workbook_path = os.path.abspath(os.fspath(workbook_path))
    if not os.path.exists(workbook_path):
        raise FileNotFoundError(workbook_path)
    current = load_database(database_path)
    info = get_database_info(database_path)
    wb = load_workbook(workbook_path, data_only=False)
    metadata = _workbook_metadata(wb)
    if int(metadata.get("format_version") or 0) != WORKBOOK_VERSION:
        raise SpreadsheetFormatError("工作簿格式版本不兼容，请从当前应用重新导出")
    if str(metadata.get("database_id") or "") != info["database_id"]:
        raise SpreadsheetFormatError("工作簿来自另一个数据库，不能导入当前数据库")
    try:
        base_revision = int(metadata.get("base_revision"))
    except (TypeError, ValueError) as exc:
        raise SpreadsheetFormatError("工作簿缺少有效的数据库修订号") from exc
    base_hash = str(metadata.get("base_hash") or "")
    current_hash = _canonical_hash(current)
    if len(base_hash) != 64:
        raise SpreadsheetFormatError("工作簿缺少有效的导出数据摘要，请重新导出")
    if base_revision == info["revision"] and base_hash != current_hash:
        raise StaleImportError("数据库内容已被外部修改，请重新导出编辑副本")

    incoming = core.empty_data()
    for table in core.TABLES:
        if table not in wb.sheetnames:
            raise SpreadsheetFormatError(f"缺少工作表：{table}")
        ws = wb[table]
        editable = [f for f in core.SCHEMA[table] if f[0] not in core.DERIVED.get(table, [])]
        expected_headers = [ID_HEADER, *[f[0] for f in editable]]
        headers = [cell.value for cell in ws[1]]
        actual = [x for x in headers if x not in (None, "")]
        if len(actual) != len(set(actual)):
            raise SpreadsheetFormatError(f"{table} 表头存在重复列")
        if set(actual) != set(expected_headers):
            missing = [x for x in expected_headers if x not in actual]
            extra = [x for x in actual if x not in expected_headers]
            details = []
            if missing:
                details.append("缺少 " + "、".join(missing))
            if extra:
                details.append("多出 " + "、".join(str(x) for x in extra))
            raise SpreadsheetFormatError(f"{table} 表头已改变：" + "；".join(details))
        col_by_header = {value: index + 1 for index, value in enumerate(headers)
                         if value not in (None, "")}
        current_by_id = {str(r.get("记录ID") or ""): r for r in current[table]}
        seen_ids = set()
        for row_number in range(2, ws.max_row + 1):
            cells = [ws.cell(row_number, col_by_header[h]) for h in expected_headers]
            if not any(cell.value not in (None, "") for cell in cells):
                continue
            id_cell = cells[0]
            if id_cell.data_type == "f":
                raise SpreadsheetFormatError(f"{table} 第 {row_number} 行记录ID不能是公式")
            rid = str(id_cell.value or "").strip()
            if rid in seen_ids:
                rid = ""  # 复制整行时隐藏 ID 也会被复制；后续重复项按新增记录处理。
            elif rid:
                seen_ids.add(rid)
            if rid and rid not in current_by_id:
                raise SpreadsheetFormatError(
                    f"{table} 第 {row_number} 行记录ID不属于当前数据库，请勿修改隐藏的 _记录ID 列"
                )
            rec = dict(current_by_id[rid]) if rid else {}
            rec["记录ID"] = rid
            for field, field_type, *_ in editable:
                cell = ws.cell(row_number, col_by_header[field])
                rec[field] = _cell_value(cell, field, field_type, table, row_number)
            incoming[table].append(rec)

    incoming = _propagate_import_relational_edits(current, incoming)
    prepared_data, issues = _prepare_data(incoming)
    changes, action_counts, table_counts = _diff_data(current, prepared_data)
    stat = os.stat(workbook_path)
    return PreparedImport(
        database_path=database_path,
        workbook_path=workbook_path,
        database_id=info["database_id"],
        database_hash=current_hash,
        base_revision=base_revision,
        current_revision=info["revision"],
        data=prepared_data,
        rules=get_rules(database_path),
        changes=changes,
        action_counts=action_counts,
        table_counts=table_counts,
        issues=issues,
        file_signature=(stat.st_size, stat.st_mtime_ns),
    )


def apply_editable_import(database_path, prepared):
    database_path = os.path.abspath(os.fspath(database_path))
    if database_path != prepared.database_path:
        raise StaleImportError("预览对应的数据库已切换，请重新选择文件并预览")
    current_info = get_database_info(database_path)
    if prepared.stale or current_info["revision"] != prepared.current_revision:
        raise StaleImportError("数据库在工作簿导出或差异预览后已发生变化，请重新导出/预览")
    if current_info["database_id"] != prepared.database_id:
        raise StaleImportError("当前数据库与差异预览不匹配")
    if _canonical_hash(load_database(database_path)) != prepared.database_hash:
        raise StaleImportError("数据库内容在差异预览后又被修改，请重新预览")
    stat = os.stat(prepared.workbook_path)
    if (stat.st_size, stat.st_mtime_ns) != prepared.file_signature:
        raise StaleImportError("Excel 文件在差异预览后又被修改，请重新预览")
    return save_database(
        database_path,
        prepared.data,
        prepared.rules,
        reason="xlsx_import",
        backup=True,
        expected_revision=prepared.current_revision,
    )
