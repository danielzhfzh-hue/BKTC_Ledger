#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""台账维护工具核心层（纯 Python，macOS / Windows 通用）。

负责六表数据模型、JSON/SQLite 读写、飞书导出导入、字段规范化/自动派生、
业务校验，以及调用 build_ledger_main 生成完整台账 Excel。
"""
import hashlib
import json
import os
import re
import shutil
import sys
import time
from datetime import datetime, timedelta

from openpyxl import load_workbook

from build_ledger_main import (  # noqa: E402
    DEFAULT_RULES,
    build_from_data,
    compute_unpaid_rows,
    cov_match,
    get_column_letter,
    load_table,
    matching_records,
    num,
    s,
)

TABLES = ["合同订单", "付款条件", "设备台账", "发货批次", "开票记录", "回款记录"]

SELECT = "select"
TEXT = "text"
NUMBER = "number"
INT = "int"
DATE = "date"

SCHEMA = {
    "合同订单": [
        ("JOB No", TEXT), ("客户", TEXT), ("担当者", TEXT), ("订单内容", TEXT),
        ("付款条件", TEXT), ("发货方式", TEXT), ("发货地点", TEXT), ("送货地点", TEXT),
        ("币种", TEXT), ("设备型号", TEXT), ("总台数", INT), ("备注", TEXT),
    ],
    "付款条件": [
        ("JOB No", TEXT), ("客户", TEXT), ("款类", SELECT,
                           ["预付款", "发货款", "到货款", "验收款", "质保款"]),
        ("比例%", NUMBER), ("账期天数", INT), ("触发条件", SELECT,
         ["开票后", "货到签收后", "验收合格后", "验收后", "质保期满后",
          "合同生效后", "月结次月月底"]),
        ("说明", TEXT),
    ],
    "设备台账": [
        ("JOB No", TEXT), ("客户", TEXT), ("PO No", TEXT), ("设备型号", TEXT), ("製造番号", TEXT),
        ("機番", TEXT), ("未税单价", NUMBER), ("是否无偿", SELECT, ["否", "是"]),
        ("发货批次", TEXT), ("送货单回收", SELECT, ["", "已签收"]),
        ("验收状态", SELECT, ["未验收", "已验收"]), ("质保开始日", DATE),
        ("质保结束日", DATE), ("质保期", TEXT), ("备注", TEXT),
    ],
    "发货批次": [
        ("JOB No", TEXT), ("客户", TEXT), ("发货批次", TEXT), ("出荷日", DATE), ("台数", INT),
        ("未税合计", NUMBER), ("含税合计", NUMBER), ("覆盖製造番号", TEXT),
    ],
    "开票记录": [
        ("JOB No", TEXT), ("客户", TEXT), ("款类", SELECT,
                           ["预付款", "发货款", "到货款", "验收款", "质保款", "全额"]),
        ("开票日", DATE), ("状态", SELECT, ["已开"]), ("含税金额", NUMBER),
        ("覆盖批次", TEXT), ("覆盖製造番号", TEXT), ("覆盖台数", INT),
        ("账期天数", INT), ("应收回款日", DATE),
        ("回款状态", SELECT, ["未回款", "部分回款", "超期未回", "已回款"]),
    ],
    "回款记录": [
        ("JOB No", TEXT), ("客户", TEXT), ("款类", SELECT,
                           ["预付款", "发货款", "到货款", "验收款", "质保款"]),
        ("回款日", DATE), ("含税金额", NUMBER), ("覆盖批次", TEXT),
        ("覆盖製造番号", TEXT), ("覆盖台数", INT), ("对应应收回款日", DATE),
        ("是否超期", SELECT, ["", "否", "是"]), ("超期天数", INT),
    ],
}

DERIVED = {
    "合同订单": ["设备型号", "总台数", "付款条件"],
    "付款条件": ["客户"],
    "发货批次": ["客户", "台数", "未税合计", "含税合计", "覆盖製造番号"],
    "设备台账": ["客户", "验收状态"],
    "开票记录": ["客户", "覆盖台数", "应收回款日", "回款状态"],
    "回款记录": ["客户", "覆盖台数", "对应应收回款日", "是否超期", "超期天数"],
}


def empty_data():
    return {t: [] for t in TABLES}


def _norm(v, typ):
    if v is None:
        return None if typ in (NUMBER, INT) else ""  # 数字字段保留 null=未填(区别于 0),下游一律 `or 0` 兜底
    if typ == TEXT:
        return s(v)
    if typ == SELECT:
        if isinstance(v, list):
            v = v[0] if v else ""
        return s(v)
    if typ == DATE:
        st = s(v)
        m = re.match(r"(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})", st)
        if m:
            return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
        return st  # 非标准格式原样保留(不再静默清空),交给 validate 报 warning
    if typ == NUMBER:
        try:
            return float(v)
        except (TypeError, ValueError):
            return 0.0
    if typ == INT:
        try:
            return int(float(v))
        except (TypeError, ValueError):
            return 0
    return s(v)


def normalize(data):
    """规范化六表已知字段类型；未知字段原样保留（如关联、公式等）。"""
    out = empty_data()
    for table in TABLES:
        for rec in data.get(table, []):
            rec = dict(rec or {})
            for field, typ, *_opt in SCHEMA[table]:
                rec[field] = _norm(rec.get(field), typ)
            out[table].append(rec)
    return out


def coverage_count(text):
    return len(set(coverage_values(text)))


def coverage_values(text):
    """拆分覆盖批次/製造番号，同时接受中英文分号和换行。"""
    return [x.strip() for x in re.split(r"[;；\r\n]+", s(text)) if x.strip()]


def derive(data):
    """自动派生：发货批次合计、合同总台数/设备型号、覆盖台数等。"""
    data = normalize(data)
    customer_by_job = {s(c["JOB No"]): s(c["客户"]) for c in data["合同订单"]}
    for table in ("付款条件", "设备台账", "发货批次", "开票记录", "回款记录"):
        for rec in data[table]:
            rec["客户"] = customer_by_job.get(s(rec["JOB No"]), "")

    dev_by = {}
    for d in data["设备台账"]:
        dev_by.setdefault((s(d["JOB No"]), s(d["发货批次"])), []).append(d)

    for x in data["发货批次"]:
        devs = dev_by.get((s(x["JOB No"]), s(x["发货批次"])), [])
        x["台数"] = len(devs)
        x["未税合计"] = round(sum(float(d.get("未税单价") or 0) for d in devs), 3)
        x["含税合计"] = round(x["未税合计"] * 1.13, 2)
        x["覆盖製造番号"] = ";".join(d["製造番号"] for d in devs if s(d["製造番号"]))

    # 付款条件文本 = 条款说明按"；"拼接（跨表：改条款→合同文本更新）；无条款保留原值
    terms_by_job = {}
    for t in data["付款条件"]:
        terms_by_job.setdefault(s(t["JOB No"]), []).append(t)

    dev_count = {}
    dev_models = {}
    for d in data["设备台账"]:
        job = s(d["JOB No"])
        dev_count[job] = dev_count.get(job, 0) + 1
        dev_models.setdefault(job, set()).add(s(d["设备型号"]))
        d["验收状态"] = "已验收" if s(d.get("质保开始日")) else "未验收"
    for c in data["合同订单"]:
        job = s(c["JOB No"])
        c["总台数"] = dev_count.get(job, 0)
        if not s(c.get("设备型号")):
            models = sorted(m for m in dev_models.get(job, set()) if m)
            c["设备型号"] = ";".join(models)
        if not s(c.get("币种")):
            c["币种"] = "RMB"
        terms = terms_by_job.get(job, [])
        if terms:
            desc = "；".join(s(t.get("说明")).strip() for t in terms if s(t.get("说明")).strip())
            if desc:
                c["付款条件"] = desc

    for table in ("开票记录", "回款记录"):
        for x in data[table]:
            job_devices = [d for d in data["设备台账"]
                           if s(d["JOB No"]) == s(x["JOB No"])]
            has_coverage = (coverage_values(x.get("覆盖批次"))
                            or coverage_values(x.get("覆盖製造番号")))
            x["覆盖台数"] = (sum(
                1 for d in job_devices
                if cov_match(x, s(d.get("製造番号")), s(d.get("发货批次")))
            ) if has_coverage else 0)
            coverage_tokens = coverage_values(x.get("覆盖製造番号"))
            if coverage_tokens:
                batches = []
                safe_to_rebuild = True
                for token in coverage_tokens:
                    parts = {p for p in re.split(r"[→⇒⟶➡➝]", token) if p}
                    matches = [d for d in job_devices
                               if s(d["製造番号"]) == token
                               or s(d["製造番号"]) in parts]
                    if len(matches) != 1:
                        safe_to_rebuild = False
                        break
                    batch = s(matches[0]["发货批次"])
                    if batch and batch not in batches:
                        batches.append(batch)
                blank_serial_batches = {s(d["发货批次"]) for d in job_devices
                                        if not s(d["製造番号"])}
                removed_batches = set(coverage_values(x.get("覆盖批次"))) - set(batches)
                if removed_batches & blank_serial_batches:
                    safe_to_rebuild = False
                if safe_to_rebuild and batches:
                    x["覆盖批次"] = ";".join(batches)
            if table == "开票记录":
                invd = s(x.get("开票日"))
                x["应收回款日"] = ""
                if invd:
                    try:
                        days = int(float(x.get("账期天数") or 0))
                        x["应收回款日"] = (datetime.strptime(invd[:10], "%Y-%m-%d")
                                            + timedelta(days=days)).strftime("%Y-%m-%d")
                    except ValueError:
                        pass

    # 状态字段同金额事实保持一致，不再依赖手工回填。
    devs_by_job = {}
    for dev in data["设备台账"]:
        devs_by_job.setdefault(s(dev["JOB No"]), []).append(dev)
    invoices_by_job = {}
    for inv in data["开票记录"]:
        invoices_by_job.setdefault(s(inv["JOB No"]), []).append(inv)
    payments_by_job = {}
    for payment in data["回款记录"]:
        payments_by_job.setdefault(s(payment["JOB No"]), []).append(payment)

    def covered_devices(record, job_devices):
        return [d for d in job_devices
                if s(d.get("是否无偿")) != "是"
                and cov_match(record, s(d.get("製造番号")), s(d.get("发货批次")))]

    def allocated_amount(record, selected_devices, job_devices):
        all_covered = covered_devices(record, job_devices)
        total_price = sum(num(d.get("未税单价")) or 0 for d in all_covered)
        selected_price = sum(num(d.get("未税单价")) or 0 for d in selected_devices
                             if d in all_covered)
        return ((num(record.get("含税金额")) or 0) * selected_price / total_price
                if total_price else 0.0)

    today = datetime.now().date()
    for inv in data["开票记录"]:
        job = s(inv["JOB No"])
        kind = s(inv["款类"])
        meaningful = (kind or s(inv.get("开票日"))
                      or inv.get("含税金额") not in (None, "", 0, 0.0)
                      or s(inv.get("覆盖批次")) or s(inv.get("覆盖製造番号")))
        if not meaningful:
            inv["回款状态"] = ""
            continue
        job_devices = devs_by_job.get(job, [])
        candidates = [p for p in payments_by_job.get(job, [])
                      if kind == "全额" or s(p["款类"]) == kind]
        related_invoices = [other for other in invoices_by_job.get(job, [])
                            if s(other["款类"]) == kind]
        invoice_devices = covered_devices(inv, job_devices)
        if job_devices:
            paid = sum(allocated_amount(p, invoice_devices, job_devices) for p in candidates)
            invoiced = sum(allocated_amount(other, invoice_devices, job_devices)
                           for other in related_invoices)
        else:
            paid = sum(num(p.get("含税金额")) or 0 for p in candidates)
            invoiced = sum(num(other.get("含税金额")) or 0
                           for other in related_invoices)
        amount = invoiced or (num(inv.get("含税金额")) or 0)
        due = None
        try:
            due = datetime.strptime(s(inv.get("应收回款日"))[:10], "%Y-%m-%d").date()
        except ValueError:
            pass
        if amount > 0 and paid >= amount - 0.01:
            inv["回款状态"] = "已回款"
        elif paid > 0.01:
            inv["回款状态"] = "部分回款"
        elif due and due < today:
            inv["回款状态"] = "超期未回"
        else:
            inv["回款状态"] = "未回款"

    for payment in data["回款记录"]:
        job = s(payment["JOB No"])
        kind = s(payment["款类"])
        meaningful = (kind or s(payment.get("回款日"))
                      or payment.get("含税金额") not in (None, "", 0, 0.0)
                      or s(payment.get("覆盖批次"))
                      or s(payment.get("覆盖製造番号")))
        if not meaningful:
            payment["对应应收回款日"] = ""
            payment["是否超期"] = ""
            payment["超期天数"] = None
            continue
        job_devices = devs_by_job.get(job, [])
        candidates = [inv for inv in invoices_by_job.get(job, [])
                      if s(inv["款类"]) in (kind, "全额")]
        if job_devices:
            relevant = []
            for dev in covered_devices(payment, job_devices):
                for inv in matching_records(candidates, s(dev["製造番号"]),
                                            s(dev["发货批次"])):
                    if inv not in relevant:
                        relevant.append(inv)
        else:
            relevant = candidates
        due_dates = []
        for inv in relevant:
            try:
                due_dates.append(datetime.strptime(
                    s(inv.get("应收回款日"))[:10], "%Y-%m-%d"
                ).date())
            except ValueError:
                continue
        due = min(due_dates) if due_dates else None
        payment["对应应收回款日"] = due.strftime("%Y-%m-%d") if due else ""
        paid_on = None
        try:
            paid_on = datetime.strptime(s(payment.get("回款日"))[:10], "%Y-%m-%d").date()
        except ValueError:
            pass
        if due and paid_on:
            late_days = max(0, (paid_on - due).days)
            payment["是否超期"] = "是" if late_days else "否"
            payment["超期天数"] = late_days
        else:
            payment["是否超期"] = ""
            payment["超期天数"] = None
    return data


def _job_exists(data, job):
    return any(s(c["JOB No"]) == job for c in data["合同订单"])


def validate(data):
    """业务校验，返回 [{severity, msg}]；severity: error / warning / info。"""
    issues = []
    contract_jobs = [s(c.get("JOB No")) for c in data["合同订单"]]
    jobs = set(contract_jobs)
    seen_jobs = set()
    for c in data["合同订单"]:
        job = s(c["JOB No"])
        if not re.fullmatch(r"\d{2}(BS|DS)\d{3}", job):
            issues.append({"severity": "error", "msg": f"合同 JOB No 格式不正确：{job!r}"})
        if job in seen_jobs:
            issues.append({"severity": "error", "msg": f"合同 JOB No 重复：{job!r}"})
        seen_jobs.add(job)
    for table in TABLES:
        seen_ids = set()
        for rec in data[table]:
            rid = s(rec.get("记录ID"))
            if rid and rid in seen_ids:
                issues.append({"severity": "error",
                               "msg": f"{table} 记录ID重复：{rid!r}"})
            if rid:
                seen_ids.add(rid)
        for field, field_type, *options in SCHEMA[table]:
            if field_type != SELECT or not options:
                continue
            allowed = set(options[0])
            for rec in data[table]:
                value = s(rec.get(field))
                if value and value not in allowed:
                    issues.append({"severity": "error",
                                   "msg": f"{table} {field}值不在允许列表中：{value!r}"})
    for table in ("付款条件", "设备台账", "发货批次", "开票记录", "回款记录"):
        for rec in data[table]:
            job = s(rec.get("JOB No"))
            if not job:
                issues.append({"severity": "error", "msg": f"{table} 存在缺少 JOB No 的记录"})
            elif job not in jobs:
                issues.append({"severity": "error",
                               "msg": f"{table} 引用了不存在的 JOB No：{job!r}"})
    sums = {}
    for t in data["付款条件"]:
        if not s(t.get("款类")):
            continue  # 新建订单的空占位行不构成一组付款比例
        key = s(t.get("JOB No"))
        sums[key] = sums.get(key, 0.0) + float(t.get("比例%") or 0)
    for job, total in sorted(sums.items()):
        if abs(total - 100) > 0.01:
            issues.append({"severity": "error",
                           "msg": f"{job} 付款条件比例合计 {total:g}% ≠ 100%"})
    term_kinds = {}
    for term in data["付款条件"]:
        if s(term.get("款类")):
            term_kinds.setdefault(s(term.get("JOB No")), set()).add(s(term.get("款类")))
    for table in ("开票记录", "回款记录"):
        for rec in data[table]:
            job, kind = s(rec.get("JOB No")), s(rec.get("款类"))
            if not kind or (table == "开票记录" and kind == "全额"):
                continue
            if kind not in term_kinds.get(job, set()):
                issues.append({
                    "severity": "error",
                    "msg": f"{table} {job} 款类 {kind!r} 不在该 JOB 的付款条件中",
                })
    seen = {}
    seen_full = {}
    KNOWN_DUP = {("23BS004", "23BS004-058"), ("23BS004", "23BS004-061")}
    for d in data["设备台账"]:
        job, ser, kiki = s(d.get("JOB No")), s(d.get("製造番号")), s(d.get("機番"))
        if not ser:
            issues.append({"severity": "warning",
                           "msg": f"{job} 设备台账存在空製造番号（将按批次匹配覆盖）"})
            continue
        if (job, ser) in seen:
            if (job, ser) not in KNOWN_DUP:
                issues.append({"severity": "warning",
                               "msg": f"製造番号重复：{job} {ser}"})
        else:
            seen[(job, ser)] = 1
        if (job, ser, kiki) in seen_full:
            issues.append({"severity": "error",
                           "msg": f"设备业务键重复 (JOB, 製造番号, 機番)：{job} {ser} {kiki}"})
        else:
            seen_full[(job, ser, kiki)] = 1
        if d.get("质保开始日") and d.get("质保结束日") and d["质保结束日"] < d["质保开始日"]:
            issues.append({"severity": "warning",
                           "msg": f"{job} {ser} 质保结束日早于开始日"})
    for x in data["开票记录"]:
        if not s(x.get("开票日")):
            issues.append({"severity": "warning",
                           "msg": f"{s(x.get('JOB No'))} 开票记录缺少开票日"})
        if float(x.get("含税金额") or 0) <= 0:
            issues.append({"severity": "warning",
                           "msg": f"{s(x.get('JOB No'))} 开票金额 ≤ 0"})
        if not s(x.get("覆盖批次")) and not s(x.get("覆盖製造番号")):
            issues.append({"severity": "info",
                           "msg": f"{s(x.get('JOB No'))} 开票未填覆盖批次/覆盖製造番号，"
                                  "可能无法参与未回收计算"})
    for p in data["回款记录"]:
        if not s(p.get("回款日")):
            issues.append({"severity": "warning",
                           "msg": f"{s(p.get('JOB No'))} 回款记录缺少回款日"})
        if float(p.get("含税金额") or 0) <= 0:
            issues.append({"severity": "warning",
                           "msg": f"{s(p.get('JOB No'))} 回款金额 ≤ 0"})
        meaningful = (s(p.get("款类")) or s(p.get("回款日"))
                      or p.get("含税金额") not in (None, "", 0, 0.0))
        job_has_devices = any(s(d.get("JOB No")) == s(p.get("JOB No"))
                              for d in data["设备台账"])
        if (meaningful and job_has_devices and not s(p.get("覆盖批次"))
                and not s(p.get("覆盖製造番号"))):
            issues.append({"severity": "error",
                           "msg": f"{s(p.get('JOB No'))} 回款必须选择覆盖批次或设备，"
                                  "否则该笔金额不会参与未回收计算"})
    shipment_seen = set()
    for x in data["发货批次"]:
        key = (s(x.get("JOB No")), s(x.get("发货批次")))
        if not key[1]:
            issues.append({"severity": "error", "msg": f"{key[0]} 发货批次名称不能为空"})
        if key in shipment_seen:
            issues.append({"severity": "error",
                           "msg": f"发货批次业务键重复：{key[0]} {key[1]}"})
        shipment_seen.add(key)
    dev_batches = {(s(d["JOB No"]), s(d["发货批次"])) for d in data["设备台账"]}
    for d in data["设备台账"]:
        job, batch = s(d.get("JOB No")), s(d.get("发货批次"))
        if batch and (job, batch) not in shipment_seen:
            issues.append({"severity": "error",
                           "msg": f"{job} 设备引用了不存在的发货批次：{batch!r}"})
    device_serials = {(s(d.get("JOB No")), s(d.get("製造番号")))
                      for d in data["设备台账"] if s(d.get("製造番号"))}
    for table in ("开票记录", "回款记录"):
        for rec in data[table]:
            job = s(rec.get("JOB No"))
            for batch in coverage_values(rec.get("覆盖批次")):
                if (job, batch) not in shipment_seen:
                    issues.append({"severity": "error",
                                   "msg": f"{table} {job} 覆盖了不存在的批次：{batch!r}"})
            for serial in coverage_values(rec.get("覆盖製造番号")):
                if (job, serial) not in device_serials:
                    issues.append({"severity": "error",
                                   "msg": f"{table} {job} 覆盖了不存在的製造番号：{serial!r}"})
    for x in data["发货批次"]:
        if (s(x["JOB No"]), s(x["发货批次"])) not in dev_batches:
            issues.append({"severity": "warning",
                           "msg": f"{s(x['JOB No'])} 批次 {s(x['发货批次'])} 无设备，"
                                  "生成后不会显示在台账页（请补设备或删除该批次）"})
    date_fields = {t: [f for f, typ, *_ in SCHEMA[t] if typ == DATE] for t in TABLES}
    for t in TABLES:
        for rec in data[t]:
            job = s(rec.get("JOB No"))
            for f in date_fields[t]:
                v = s(rec.get(f))
                if v:
                    try:
                        datetime.strptime(v, "%Y-%m-%d")
                    except ValueError:
                        issues.append({"severity": "error",
                                       "msg": f"{job} {f}={v!r} 不是有效日期 YYYY-MM-DD"})
    return issues


def load_store(path):
    if str(path).lower().endswith((".db", ".sqlite", ".sqlite3")):
        import database
        return database.load_database(path)
    with open(path, encoding="utf-8") as f:
        raw = json.load(f)
    data = empty_data()
    for t in TABLES:
        data[t] = raw.get(t) or []
    return data


def get_rules(path):
    if str(path).lower().endswith((".db", ".sqlite", ".sqlite3")):
        import database
        return database.get_rules(path)
    rules = {}
    if os.path.exists(path):
        try:
            with open(path, encoding="utf-8") as f:
                rules = json.load(f).get("预警规则") or {}
        except (OSError, ValueError):
            pass
    return {**DEFAULT_RULES, **rules}


def save_store(path, data, rules=None):
    if str(path).lower().endswith((".db", ".sqlite", ".sqlite3")):
        import database
        return database.save_database(path, data, rules)["path"]
    data = derive(data)
    if rules is None:
        rules = get_rules(path)
    payload = {"tables": TABLES, "version": 1, "预警规则": rules}
    for t in TABLES:
        payload[t] = data[t]
    if os.path.exists(path):
        shutil.copy2(path, path + ".bak")  # 覆盖前留一份最近备份,可撤销上次保存/同步
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=1)
    os.replace(tmp, path)
    return path


def import_from_export_dir(dirpath):
    """从飞书导出文件夹导入（contracts/terms/shipments/invoices/payments.json + devices/*.json）。"""
    missing = []
    for name in ("contracts", "terms", "shipments", "invoices", "payments"):
        if not os.path.exists(os.path.join(dirpath, name + ".json")):
            missing.append(name + ".json")
    dev_dir = os.path.join(dirpath, "devices")
    if not os.path.isdir(dev_dir) or not any(f.endswith(".json") for f in os.listdir(dev_dir)):
        missing.append("devices/*.json")
    if missing:
        raise FileNotFoundError("缺少导出文件: " + ", ".join(missing))
    contracts = load_table([os.path.join(dirpath, "contracts.json")])
    terms = load_table([os.path.join(dirpath, "terms.json")])
    shipments = load_table([os.path.join(dirpath, "shipments.json")])
    invoices = load_table([os.path.join(dirpath, "invoices.json")])
    payments = load_table([os.path.join(dirpath, "payments.json")])
    devices = []
    for p in sorted(os.listdir(dev_dir)):
        if p.endswith(".json"):
            devices += load_table([os.path.join(dev_dir, p)])
    data = {
        "合同订单": contracts, "付款条件": terms, "设备台账": devices,
        "发货批次": shipments, "开票记录": invoices, "回款记录": payments,
    }
    return derive(data)


def summary(data):
    counts = {t: len(data.get(t, [])) for t in TABLES}
    contracts = data.get("合同订单", [])
    terms = data.get("付款条件", [])
    shipments = data.get("发货批次", [])
    invoices = data.get("开票记录", [])
    payments = data.get("回款记录", [])
    devices = data.get("设备台账", [])
    unpaid = compute_unpaid_rows(contracts, terms, shipments, invoices, payments, devices)
    counts["未回收行数"] = len(unpaid)
    counts["未回收合计"] = round(sum(r["未回收金额"] for r in unpaid), 2)
    return counts


def unpaid_report_rows(data, rules=None):
    """Return JSON-safe rows using the same calculation as the generated unpaid sheet."""
    data = derive(data)
    rows = compute_unpaid_rows(
        data["合同订单"], data["付款条件"], data["发货批次"],
        data["开票记录"], data["回款记录"], data["设备台账"], rules,
    )
    public_fields = [
        "客户", "JOB No", "批次", "款类", "预警等级", "开票日期",
        "预定回收日期", "未回收金额", "未回收原因",
    ]
    return [{field: row.get(field) for field in public_fields} for row in rows]


def generate_xlsx(data, out, rules=None):
    """规范化/派生后生成台账 Excel；返回 (out, issues, counts)。"""
    data = derive(data)
    issues = validate(data)
    for table, recs in data.items():
        for i, rec in enumerate(recs):
            if not rec.get("记录ID"):
                raw = json.dumps(rec, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
                digest = hashlib.sha256(f"{table}:{i}:{raw}".encode("utf-8")).hexdigest()[:20]
                rec["记录ID"] = f"rec-app-{digest}"
    build_from_data(
        data["合同订单"], data["付款条件"], data["设备台账"],
        data["发货批次"], data["开票记录"], data["回款记录"], out, rules=rules,
    )
    return out, issues, summary(data)


def backup_xlsx(path, keep=10):
    """把当前 xlsx 备份到同目录 备份/ 下，按修改时间保留最近 keep 份。返回备份路径。"""
    if not os.path.exists(path):
        return None
    bak_dir = os.path.join(os.path.dirname(os.path.abspath(path)), "备份")
    os.makedirs(bak_dir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(path))[0]
    dst = os.path.join(bak_dir, f"{stem}_{time.strftime('%Y%m%d_%H%M%S')}.xlsx")
    shutil.copy2(path, dst)
    files = [os.path.join(bak_dir, f) for f in os.listdir(bak_dir) if f.endswith(".xlsx")]
    files.sort(key=lambda p: os.path.getmtime(p))
    while len(files) > keep:
        old = files.pop(0)
        try:
            os.remove(old)
        except OSError:
            pass
    return dst


def _exp_val(v, typ=None):
    """导出单元格取值：None→空；列表→; 拼接；日期/数字按类型转真值（Excel 可排序求和）。"""
    if v is None or v == "":
        return None
    if isinstance(v, (list, tuple)):
        return ";".join(s(x) for x in v)
    if typ == DATE:
        try:
            from datetime import datetime
            return datetime.strptime(s(v)[:10], "%Y-%m-%d").date()
        except ValueError:
            v = s(v)
    if typ == NUMBER:
        try:
            return float(v)
        except (TypeError, ValueError):
            v = s(v)
    if typ == INT:
        try:
            return int(float(v))
        except (TypeError, ValueError):
            v = s(v)
    if isinstance(v, str) and v.startswith(("=", "+", "-", "@")):
        return "'" + v
    return v


def export_filtered(path, table, rows, fields):
    """把筛选后的行 + 选定字段写成 xlsx（日期/数字按类型转真值）。
    fields 可为字段名字符串列表（类型查 SCHEMA）或 [{name,type}]（含虚拟关联列类型）。"""
    from openpyxl import Workbook
    types = {f: t for f, t, *_ in SCHEMA.get(table, [])}
    norm = []
    for f in fields:
        if isinstance(f, dict):
            norm.append((f.get("name"), f.get("type")))
        else:
            norm.append((f, types.get(f)))
    wb = Workbook()
    ws = wb.active
    ws.title = (table or "导出")[:31]
    ws.append([n for n, _ in norm])
    for r in rows:
        ws.append([_exp_val(r.get(n), t) for n, t in norm])
    wb.save(path)
    return path
