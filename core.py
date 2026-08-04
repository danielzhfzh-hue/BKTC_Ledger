#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""台账维护工具核心层（纯 Python，macOS / Windows 通用）。

负责六表数据模型、records.json 读写、飞书导出导入、字段规范化/自动派生、
业务校验，以及调用 build_ledger_main 生成完整台账 Excel。
"""
import json
import os
import re
import shutil
import sys

from openpyxl import load_workbook

from build_ledger_main import (  # noqa: E402
    DEFAULT_RULES,
    build_from_data,
    compute_unpaid_rows,
    get_column_letter,
    load_table,
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
        ("JOB No", TEXT), ("款类", SELECT,
                           ["预付款", "发货款", "到货款", "验收款", "质保款"]),
        ("比例%", NUMBER), ("账期天数", INT), ("触发条件", SELECT,
         ["开票后", "货到签收后", "验收合格后", "验收后", "质保期满后",
          "合同生效后", "月结次月月底"]),
        ("说明", TEXT),
    ],
    "设备台账": [
        ("JOB No", TEXT), ("PO No", TEXT), ("设备型号", TEXT), ("製造番号", TEXT),
        ("機番", TEXT), ("未税单价", NUMBER), ("是否无偿", SELECT, ["否", "是"]),
        ("发货批次", TEXT), ("送货单回收", SELECT, ["", "已签收"]),
        ("验收状态", SELECT, ["未验收", "已验收"]), ("质保开始日", DATE),
        ("质保结束日", DATE), ("质保期", TEXT), ("备注", TEXT),
    ],
    "发货批次": [
        ("JOB No", TEXT), ("发货批次", TEXT), ("出荷日", DATE), ("台数", INT),
        ("未税合计", NUMBER), ("含税合计", NUMBER), ("覆盖製造番号", TEXT),
    ],
    "开票记录": [
        ("JOB No", TEXT), ("款类", SELECT,
                           ["预付款", "发货款", "到货款", "验收款", "质保款", "全额"]),
        ("开票日", DATE), ("状态", SELECT, ["已开"]), ("含税金额", NUMBER),
        ("覆盖批次", TEXT), ("覆盖製造番号", TEXT), ("覆盖台数", INT),
        ("账期天数", INT), ("应收回款日", DATE),
        ("回款状态", SELECT, ["未回款", "部分回款", "超期未回", "已回款"]),
    ],
    "回款记录": [
        ("JOB No", TEXT), ("款类", SELECT,
                           ["预付款", "发货款", "到货款", "验收款", "质保款"]),
        ("回款日", DATE), ("含税金额", NUMBER), ("覆盖批次", TEXT),
        ("覆盖製造番号", TEXT), ("覆盖台数", INT), ("对应应收回款日", DATE),
        ("是否超期", SELECT, ["", "否", "是"]), ("超期天数", INT),
    ],
}

DERIVED = {
    "合同订单": ["设备型号", "总台数"],
    "发货批次": ["台数", "未税合计", "含税合计", "覆盖製造番号"],
    "开票记录": ["覆盖台数"],
    "回款记录": ["覆盖台数"],
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
                if field in rec:
                    rec[field] = _norm(rec[field], typ)
            out[table].append(rec)
    return out


def coverage_count(text):
    return len({x.strip() for x in s(text).split(";") if x.strip()})


def derive(data):
    """自动派生：发货批次合计、合同总台数/设备型号、覆盖台数等。"""
    data = normalize(data)
    dev_by = {}
    for d in data["设备台账"]:
        dev_by.setdefault((s(d["JOB No"]), s(d["发货批次"])), []).append(d)

    for x in data["发货批次"]:
        devs = dev_by.get((s(x["JOB No"]), s(x["发货批次"])), [])
        x["台数"] = len(devs)
        x["未税合计"] = round(sum(float(d.get("未税单价") or 0) for d in devs), 3)
        x["含税合计"] = round(x["未税合计"] * 1.13, 2)
        x["覆盖製造番号"] = ";".join(d["製造番号"] for d in devs if s(d["製造番号"]))

    dev_count = {}
    dev_models = {}
    for d in data["设备台账"]:
        job = s(d["JOB No"])
        dev_count[job] = dev_count.get(job, 0) + 1
        dev_models.setdefault(job, set()).add(s(d["设备型号"]))
    for c in data["合同订单"]:
        job = s(c["JOB No"])
        c["总台数"] = dev_count.get(job, 0)
        if not s(c.get("设备型号")):
            models = sorted(m for m in dev_models.get(job, set()) if m)
            c["设备型号"] = ";".join(models)
        if not s(c.get("币种")):
            c["币种"] = "RMB"

    for table in ("开票记录", "回款记录"):
        for x in data[table]:
            x["覆盖台数"] = coverage_count(x.get("覆盖製造番号"))
            if not s(x.get("覆盖批次")):
                batches = {s(d["发货批次"]) for d in data["设备台账"]
                           if s(d["JOB No"]) == s(x["JOB No"])
                           and (not s(x.get("覆盖製造番号"))
                               or s(d["製造番号"]) in s(x.get("覆盖製造番号")).split(";"))}
                if len(batches) == 1:
                    x["覆盖批次"] = next(iter(batches))
    return data


def _job_exists(data, job):
    return any(s(c["JOB No"]) == job for c in data["合同订单"])


def validate(data):
    """业务校验，返回 [{severity, msg}]；severity: error / warning / info。"""
    issues = []
    jobs = {s(c["JOB No"]) for c in data["合同订单"]}
    for c in data["合同订单"]:
        job = s(c["JOB No"])
        if not re.fullmatch(r"\d{2}(BS|DS)\d{3}", job):
            issues.append({"severity": "error", "msg": f"合同 JOB No 格式不正确：{job!r}"})
    for table in ("付款条件", "设备台账", "发货批次", "开票记录", "回款记录"):
        for rec in data[table]:
            job = s(rec.get("JOB No"))
            if job and job not in jobs:
                issues.append({"severity": "error",
                               "msg": f"{table} 引用了不存在的 JOB No：{job!r}"})
    sums = {}
    for t in data["付款条件"]:
        key = s(t.get("JOB No"))
        sums[key] = sums.get(key, 0.0) + float(t.get("比例%") or 0)
    for job, total in sorted(sums.items()):
        if abs(total - 100) > 0.01:
            issues.append({"severity": "error",
                           "msg": f"{job} 付款条件比例合计 {total:g}% ≠ 100%"})
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
    dev_batches = {(s(d["JOB No"]), s(d["发货批次"])) for d in data["设备台账"]}
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
                if v and not re.fullmatch(r"\d{4}-\d{2}-\d{2}", v):
                    issues.append({"severity": "warning",
                                   "msg": f"{job} {f}={v!r} 不是 YYYY-MM-DD 格式"})
    return issues


def load_store(path):
    with open(path, encoding="utf-8") as f:
        raw = json.load(f)
    data = empty_data()
    for t in TABLES:
        data[t] = raw.get(t) or []
    return data


def get_rules(path):
    rules = {}
    if os.path.exists(path):
        try:
            with open(path, encoding="utf-8") as f:
                rules = json.load(f).get("预警规则") or {}
        except (OSError, ValueError):
            pass
    return {**DEFAULT_RULES, **rules}


def save_store(path, data, rules=None):
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


def generate_xlsx(data, out, rules=None):
    """规范化/派生后生成台账 Excel；返回 (out, issues, counts)。"""
    data = derive(data)
    issues = validate(data)
    for recs in data.values():
        for i, rec in enumerate(recs):
            if not rec.get("记录ID"):
                rec["记录ID"] = f"rec-app-{i}-{abs(hash(json.dumps(rec, ensure_ascii=False))):x}"
    build_from_data(
        data["合同订单"], data["付款条件"], data["设备台账"],
        data["发货批次"], data["开票记录"], data["回款记录"], out, rules=rules,
    )
    return out, issues, summary(data)


# ---------- 从当前 Excel 读回可见数据（用于同步手改内容） ----------

_DATE_RE = re.compile(r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})")


def _to_date(v):
    m = _DATE_RE.search(s(v))
    return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}" if m else ""


def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _merged_fill(ws):
    fill = {}
    for mr in ws.merged_cells.ranges:
        v = ws.cell(mr.min_row, mr.min_col).value
        if v is None or v == "":
            continue
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                fill[(r, c)] = v
    return fill


def _headers(ws, row=5):
    heads = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row, c).value
        if v:
            heads[str(v).strip()] = c
    return heads


def load_from_workbook(path):
    """解析当前台账 Excel 的可见业务数据（六表；付款条件表不解析，保留原数据源）。"""
    wb = load_workbook(path, data_only=False)
    nav = wb["导航页"] if "导航页" in wb.sheetnames else None
    remark_by_job = {}
    if nav is not None:
        for r in range(3, nav.max_row + 1):
            b = str(nav.cell(r, 2).value or "")
            m = re.search(r'"(\d{2}(?:BS|DS)\d{3})"', b)
            if m:
                remark_by_job[m.group(1)] = s(nav.cell(r, 13).value)

    contracts, devices, shipments, invoices, payments = [], [], [], [], []
    for sheet in wb.sheetnames:
        if sheet in ("导航页", "未回收管理表"):
            continue
        ws = wb[sheet]
        heads = _headers(ws, 5)

        def col(name):
            return heads.get(name)

        fill = _merged_fill(ws)
        job = s(ws.cell(3, col("JOB NO") or 4).value)
        if not job:
            continue
        contracts.append({
            "JOB No": job,
            "客户": s(ws.cell(3, col("取引先") or 2).value),
            "担当者": s(ws.cell(3, col("担当者") or 1).value),
            "订单内容": s(ws.cell(3, col("注文内容") or 5).value).replace("\n", "；"),
            "付款条件": s(ws.cell(3, col("支払条件") or 7).value),
            "发货方式": s(ws.cell(3, col("発送方式") or 11).value),
            "发货地点": s(ws.cell(3, col("出荷先") or 12).value),
            "送货地点": s(ws.cell(3, col("納入先") or 13).value),
            "备注": remark_by_job.get(job, ""),
            "币种": "RMB",
        })

        inv_cols = []   # (款类, 开票日c, 状况c, 金额c)
        pay_cols = []   # (款类, 回款日c, 金额c)
        inv_hidden = {}
        pay_hidden = {}
        for name, c in heads.items():
            if name.endswith("·开票日"):
                t = name[:-4]
                inv_cols.append((t, c, heads.get(t + "·开票状况"), heads.get(t + "·开票金额")))
            elif name.endswith("·回款日"):
                t = name[:-4]
                pay_cols.append((t, c, heads.get(t + "·回款金额")))
            elif name.startswith("行开票金额·"):
                inv_hidden[name[len("行开票金额·"):]] = c
            elif name.startswith("行回款金额·"):
                pay_hidden[name[len("行回款金额·"):]] = c
        due_col = heads.get("行应收回款日")

        def rowval(r, c):
            return fill.get((r, c), ws.cell(r, c).value if c else None)

        rows = list(range(6, ws.max_row))
        has_serial_rows = any(s(rowval(r, col("製造番号"))) for r in rows)
        for r in rows:
            serial = s(rowval(r, col("製造番号")))
            deviceish = serial or s(rowval(r, col("发货批次"))) or (
                has_serial_rows and (
                    s(rowval(r, col("设备型号")))
                    or _num(rowval(r, col("未税单价")))
                    or rowval(r, col("出荷日"))))
            if not deviceish:
                continue
            devices.append({
                "JOB No": job,
                "PO No": s(rowval(r, col("PO No"))),
                "设备型号": s(rowval(r, col("设备型号"))),
                "製造番号": serial,
                "機番": s(rowval(r, col("機番"))),
                "未税单价": _num(rowval(r, col("未税单价"))),
                "是否无偿": s(rowval(r, col("是否无偿"))) or "否",
                "发货批次": s(rowval(r, col("发货批次"))),
                "送货单回收": s(rowval(r, col("送货单回收"))),
                "验收状态": s(rowval(r, col("验收状态"))) or "未验收",
                "质保开始日": _to_date(rowval(r, col("质保开始日"))),
                "质保结束日": _to_date(rowval(r, col("质保结束日"))),
                "质保期": s(rowval(r, col("质保期"))),
                "备注": s(rowval(r, col("备注"))),
            })

        def _anchor_rows(t, hidden_col, amt_col, rows):
            """从隐藏公式提取每条开票/回款记录的锚点行 -> 覆盖行列表。"""
            letter = get_column_letter(amt_col)
            groups = {}
            for r in rows:
                f = s(ws.cell(r, hidden_col).value)
                m = re.search(rf"{letter}(\d+)", f)
                if m and int(m.group(1)) >= 6:
                    groups.setdefault(int(m.group(1)), []).append(r)
            return groups

        def collect_blocks(blocks):
            out = []
            for t, dc, sc, ac in blocks:
                hid = inv_hidden.get(t)
                groups = _anchor_rows(t, hid, ac, rows) if hid else {}
                if not groups:
                    # 无设备/汇总行（如 26BS009）：以可见值本身作为一条记录
                    for r in rows:
                        d = _to_date(rowval(r, dc))
                        if d:
                            groups.setdefault(r, [r])
                for anchor, group in groups.items():
                    ar = anchor
                    d = _to_date(ws.cell(ar, dc).value)
                    if not d:
                        d = _to_date(rowval(ar, dc))
                    amt = round(_num(ws.cell(ar, ac).value or rowval(ar, ac)), 2)
                    status = s(ws.cell(ar, sc).value or rowval(ar, sc))
                    serials = [s(rowval(rr, col("製造番号"))) for rr in group]
                    batches = [s(rowval(rr, col("发货批次"))) for rr in group]
                    dues = [_to_date(rowval(rr, due_col)) for rr in group if due_col]
                    out.append({
                        "JOB No": job, "款类": t, "开票日": d,
                        "状态": "已开" if status == "已开票" else "",
                        "含税金额": amt,
                        "覆盖批次": ";".join(dict.fromkeys(b for b in batches if b)),
                        "覆盖製造番号": ";".join(dict.fromkeys(x for x in serials if x)),
                        "覆盖台数": len(group),
                        "应收回款日": dues[0] if dues and len(set(dues)) == 1 else "",
                        "账期天数": 0,
                        "回款状态": "",
                    })
            return out

        def collect_pays(blocks):
            out = []
            for t, dc, ac in blocks:
                hid = pay_hidden.get(t)
                groups = _anchor_rows(t, hid, ac, rows) if hid else {}
                if not groups:
                    for r in rows:
                        d = _to_date(rowval(r, dc))
                        if d:
                            groups.setdefault(r, [r])
                for anchor, group in groups.items():
                    ar = anchor
                    d = _to_date(ws.cell(ar, dc).value or rowval(ar, dc))
                    amt = round(_num(ws.cell(ar, ac).value or rowval(ar, ac)), 2)
                    serials = [s(rowval(rr, col("製造番号"))) for rr in group]
                    batches = [s(rowval(rr, col("发货批次"))) for rr in group]
                    out.append({
                        "JOB No": job, "款类": t, "回款日": d,
                        "含税金额": amt,
                        "覆盖批次": ";".join(dict.fromkeys(b for b in batches if b)),
                        "覆盖製造番号": ";".join(dict.fromkeys(x for x in serials if x)),
                        "覆盖台数": len(group),
                        "对应应收回款日": "", "是否超期": "", "超期天数": 0,
                    })
            return out

        invoices += collect_blocks(inv_cols)
        payments += collect_pays(pay_cols)

        # 发货批次：按批次聚合设备行（无设备行仍保留批次，但台数/覆盖只算有设备的行）
        batch_rows = {}
        for r in rows:
            b = s(rowval(r, col("发货批次")))
            if b:
                batch_rows.setdefault(b, []).append(r)
        for b, brs in batch_rows.items():
            dev_rows = [rr for rr in brs if s(rowval(rr, col("製造番号")))]
            serials = [s(rowval(rr, col("製造番号"))) for rr in dev_rows]
            shipments.append({
                "JOB No": job, "发货批次": b,
                "出荷日": _to_date(rowval(brs[0], col("出荷日"))),
                "台数": len(dev_rows),
                "未税合计": round(sum(_num(rowval(rr, col("未税单价"))) for rr in dev_rows), 3),
                "含税合计": 0.0,
                "覆盖製造番号": ";".join(dict.fromkeys(x for x in serials if x)),
            })
    for x in shipments:
        x["含税合计"] = round(x["未税合计"] * 1.13, 2)
    return {"合同订单": contracts, "付款条件": [], "设备台账": devices,
            "发货批次": shipments, "开票记录": invoices, "回款记录": payments}


def merge_workbook(store, parsed):
    """把 Excel 解析出的可见改动并入 store；保留 store 独有的字段（条款、应收日等）。
    返回 (新store, 变更摘要)。"""
    store = derive(store)
    changed = {"合同订单": {"新增": 0, "更新": 0, "删除": 0},
               "付款条件": {"新增": 0, "更新": 0, "删除": 0},
               "设备台账": {"新增": 0, "更新": 0, "删除": 0},
               "发货批次": {"新增": 0, "更新": 0, "删除": 0},
               "开票记录": {"新增": 0, "更新": 0, "删除": 0},
               "回款记录": {"新增": 0, "更新": 0, "删除": 0}}

    # 合同订单
    by_job = {s(c["JOB No"]): c for c in store["合同订单"]}
    parsed_jobs = {s(c["JOB No"]) for c in parsed["合同订单"]}
    for pc in parsed["合同订单"]:
        job = s(pc["JOB No"])
        if job in by_job:
            c = by_job[job]
            changed["合同订单"]["更新"] += 1
        else:
            c = dict(pc)
            c["记录ID"] = f"rec-xlsx-{job}"
            store["合同订单"].append(c)
            changed["合同订单"]["新增"] += 1
        for f in ("客户", "担当者", "订单内容", "付款条件", "发货方式", "发货地点", "送货地点"):
            c[f] = pc.get(f, c.get(f, ""))
        if pc.get("备注") is not None:
            c["备注"] = pc["备注"]
        c["币种"] = c.get("币种") or "RMB"
    removed = [c for c in store["合同订单"] if s(c["JOB No"]) not in parsed_jobs]
    changed["合同订单"]["删除"] = len(removed)
    store["合同订单"] = [c for c in store["合同订单"] if s(c["JOB No"]) in parsed_jobs]

    def merge_list(store_recs, parsed_recs, keyf, visible_fields, table):
        index = {}
        for i, rec in enumerate(store_recs):
            index.setdefault(keyf(rec), []).append(i)
        matched = set()
        result = []
        for pr in parsed_recs:
            cands = [i for i in index.get(keyf(pr), []) if i not in matched]
            if cands:
                i = cands[0]
                matched.add(i)
                rec = store_recs[i]
                changed[table]["更新"] += 1
            else:
                rec = dict(pr)
                rec["记录ID"] = f"rec-xlsx-{table}-{len(store_recs)}"
                store_recs.append(rec)
                i = len(store_recs) - 1
                matched.add(i)
                changed[table]["新增"] += 1
            for f in visible_fields:
                rec[f] = pr.get(f, rec.get(f, ""))
            result.append(rec)
        changed[table]["删除"] = len(store_recs) - len(result)
        return result

    # 设备台账：键 (JOB, 製造番号)，机番变更也能被更新
    store["设备台账"] = merge_list(
        store["设备台账"], parsed["设备台账"],
        lambda d: (s(d.get("JOB No")), s(d.get("製造番号"))),
        ["PO No", "设备型号", "機番", "未税单价", "是否无偿", "发货批次", "送货单回收",
         "验收状态", "质保开始日", "质保结束日", "质保期", "备注"],
        "设备台账")
    # 发货批次：键 (JOB, 发货批次)
    store["发货批次"] = merge_list(
        store["发货批次"], parsed["发货批次"],
        lambda x: (s(x.get("JOB No")), s(x.get("发货批次"))),
        ["出荷日", "台数", "未税合计", "含税合计", "覆盖製造番号"],
        "发货批次")
    # 开票记录：键 (JOB, 款类, 开票日)，金额变化视为更新
    store["开票记录"] = merge_list(
        store["开票记录"], parsed["开票记录"],
        lambda x: (s(x.get("JOB No")), s(x.get("款类")), s(x.get("开票日"))),
        # 覆盖串保留 store 的权威值；仅同步用户可见可改的金额与状态
        ["含税金额", "状态"],
        "开票记录")
    # 回款记录：键 (JOB, 款类, 回款日)
    store["回款记录"] = merge_list(
        store["回款记录"], parsed["回款记录"],
        lambda x: (s(x.get("JOB No")), s(x.get("款类")), s(x.get("回款日"))),
        ["含税金额"],
        "回款记录")
    return derive(store), changed


def reconcile_store_with_excel(store_path, xlsx_path):
    """读取 store 与 Excel，把 Excel 手改内容并入 store 并保存。"""
    store = load_store(store_path)
    parsed = load_from_workbook(xlsx_path)
    store, changed = merge_workbook(store, parsed)
    save_store(store_path, store)
    return store, changed
