#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ERP V2 domain calculations.

This module is intentionally pure: it turns the six authoritative business tables
into read-only order snapshots, action signals, document flow and finance views.
SQLite persistence stays in database.py; UI code should consume these derived
structures instead of re-implementing business rules.
"""
from __future__ import annotations

import copy
import hashlib
from collections import defaultdict
from datetime import date
from typing import Any

import core


STAGE_LABELS = (
    "已受注", "待生产", "履约/出货", "验收", "开票", "回款", "已结清",
)
PRIORITY_ORDER = {"high": 0, "medium": 1, "low": 2}


def _text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def _num(value: Any) -> float:
    try:
        return float(value or 0)
    except (TypeError, ValueError):
        return 0.0


def _date(value: Any) -> date | None:
    raw = _text(value)
    if not raw:
        return None
    try:
        return date.fromisoformat(raw[:10].replace("/", "-"))
    except (TypeError, ValueError):
        return None


def _today(value: date | str | None = None) -> date:
    if isinstance(value, date):
        return value
    if value:
        parsed = _date(value)
        if parsed:
            return parsed
    return date.today()


def _pct(numerator: float, denominator: float) -> int:
    if denominator <= 0:
        return 0
    return max(0, min(100, round(numerator / denominator * 100)))


def _stable_id(*parts: Any) -> str:
    raw = "|".join(_text(x) for x in parts)
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:16]


def _group_amount(rows: list[dict[str, Any]], amount_field: str = "amount") -> dict[str, float]:
    totals: dict[str, float] = defaultdict(float)
    for row in rows:
        totals[_text(row.get("currency")) or "RMB"] += _num(row.get(amount_field))
    return {currency: round(amount, 2) for currency, amount in sorted(totals.items())}


def _document_flow(job: str, quote_id: str, quote_no: str,
                   shipments: list[dict[str, Any]], invoices: list[dict[str, Any]],
                   payments: list[dict[str, Any]]) -> dict[str, Any]:
    nodes: list[dict[str, Any]] = []
    edges: list[dict[str, str]] = []
    order_id = f"order:{job}"
    if quote_id or quote_no:
        quote_key = quote_id or quote_no
        nodes.append({"id": f"quote:{quote_key}", "type": "quote", "label": quote_no or "来源报价"})
        edges.append({"from": f"quote:{quote_key}", "to": order_id})
    nodes.append({"id": order_id, "type": "order", "label": job})

    shipment_nodes = {}
    for row in shipments:
        batch = _text(row.get("发货批次")) or "未命名批次"
        node_id = f"shipment:{job}:{batch}"
        shipment_nodes[batch] = node_id
        nodes.append({
            "id": node_id, "type": "shipment", "label": batch,
            "date": _text(row.get("出荷日")),
            "count": int(_num(row.get("台数"))),
        })
        edges.append({"from": order_id, "to": node_id})

    invoice_by_kind_batch: dict[tuple[str, str], list[str]] = defaultdict(list)
    invoice_by_kind: dict[str, list[str]] = defaultdict(list)
    for index, row in enumerate(invoices):
        rid = _text(row.get("记录ID")) or f"{index}"
        node_id = f"invoice:{rid}"
        kind = _text(row.get("款类")) or "开票"
        batches = [x.strip() for x in _text(row.get("覆盖批次")).replace("；", ";").split(";") if x.strip()]
        nodes.append({
            "id": node_id, "type": "invoice", "label": kind,
            "date": _text(row.get("开票日")), "amount": _num(row.get("含税金额")),
        })
        parents = [shipment_nodes[b] for b in batches if b in shipment_nodes] or [order_id]
        for parent in parents:
            edges.append({"from": parent, "to": node_id})
        invoice_by_kind[kind].append(node_id)
        for batch in batches:
            invoice_by_kind_batch[(kind, batch)].append(node_id)

    for index, row in enumerate(payments):
        rid = _text(row.get("记录ID")) or f"{index}"
        node_id = f"payment:{rid}"
        kind = _text(row.get("款类")) or "回款"
        batches = [x.strip() for x in _text(row.get("覆盖批次")).replace("；", ";").split(";") if x.strip()]
        nodes.append({
            "id": node_id, "type": "payment", "label": kind,
            "date": _text(row.get("回款日")), "amount": _num(row.get("含税金额")),
        })
        parents: list[str] = []
        for batch in batches:
            parents.extend(invoice_by_kind_batch.get((kind, batch), []))
        if not parents:
            parents = list(invoice_by_kind.get(kind, []))
        if not parents:
            parents = [shipment_nodes[b] for b in batches if b in shipment_nodes]
        if not parents:
            parents = [order_id]
        for parent in dict.fromkeys(parents):
            edges.append({"from": parent, "to": node_id})

    return {"nodes": nodes, "edges": edges}


def build_order_snapshots(data: dict[str, list[dict[str, Any]]], rules=None,
                          today: date | str | None = None) -> list[dict[str, Any]]:
    """Return compact order-level ERP snapshots from the six authoritative tables."""
    current = _today(today)
    derived = core.derive(copy.deepcopy(data))
    unpaid_rows = core.unpaid_report_rows(derived, rules)
    unpaid_by_job: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in unpaid_rows:
        unpaid_by_job[_text(row.get("JOB No"))].append(row)

    by_table: dict[str, dict[str, list[dict[str, Any]]]] = {}
    for table in ("付款条件", "设备台账", "发货批次", "开票记录", "回款记录"):
        grouped: dict[str, list[dict[str, Any]]] = defaultdict(list)
        for row in derived.get(table, []):
            grouped[_text(row.get("JOB No"))].append(row)
        by_table[table] = grouped

    result = []
    for order in derived.get("合同订单", []):
        job = _text(order.get("JOB No"))
        devices = by_table["设备台账"].get(job, [])
        shipments = by_table["发货批次"].get(job, [])
        invoices = [x for x in by_table["开票记录"].get(job, []) if _text(x.get("款类")) or _num(x.get("含税金额"))]
        payments = [x for x in by_table["回款记录"].get(job, []) if _text(x.get("款类")) or _num(x.get("含税金额"))]
        currency = _text(order.get("币种")) or "RMB"
        total_untaxed = round(sum(_num(x.get("未税单价")) for x in devices), 2)
        expected_gross = round(total_untaxed * 1.13, 2)

        shipment_dates = {
            _text(x.get("发货批次")): _date(x.get("出荷日")) for x in shipments
            if _text(x.get("发货批次"))
        }
        shipped_devices = [
            x for x in devices
            if shipment_dates.get(_text(x.get("发货批次")))
            and shipment_dates[_text(x.get("发货批次"))] <= current
        ]
        scheduled_devices = [
            x for x in devices
            if shipment_dates.get(_text(x.get("发货批次")))
            and shipment_dates[_text(x.get("发货批次"))] > current
        ]
        accepted_devices = [
            x for x in devices
            if _text(x.get("验收状态")) == "已验收" or _text(x.get("质保开始日"))
        ]
        pending_serial = sum(1 for x in devices if not _text(x.get("製造番号")))
        invoice_amount = round(sum(_num(x.get("含税金额")) for x in invoices), 2)
        paid_amount = round(sum(_num(x.get("含税金额")) for x in payments), 2)
        outstanding = round(sum(_num(x.get("未回收金额")) for x in unpaid_by_job.get(job, [])), 2)

        ship_pct = _pct(len(shipped_devices), len(devices))
        accept_pct = _pct(len(accepted_devices), len(devices))
        invoice_pct = _pct(invoice_amount, expected_gross) if expected_gross else (100 if invoices else 0)
        paid_pct = _pct(paid_amount, invoice_amount) if invoice_amount else 0

        stage, stage_index = "已受注", 0
        if devices and pending_serial:
            stage, stage_index = "待生产", 1
        if shipped_devices or scheduled_devices:
            stage, stage_index = "履约/出货", 2
        if devices and len(shipped_devices) == len(devices) and len(accepted_devices) < len(devices):
            stage, stage_index = "验收", 3
        if invoice_amount > 0:
            stage, stage_index = "开票", 4
        if paid_amount > 0:
            stage, stage_index = "回款", 5
        if invoice_pct >= 99 and invoice_amount > 0 and outstanding <= 0.01 and paid_pct >= 99:
            stage, stage_index = "已结清", 6

        overdue_rows, due_soon_rows = [], []
        for row in unpaid_by_job.get(job, []):
            due = _date(row.get("预定回收日期"))
            if not due:
                continue
            delta = (due - current).days
            if delta < 0:
                overdue_rows.append(row)
            elif delta <= 30:
                due_soon_rows.append(row)

        risk_level, risk_code, risk_label = "low", "normal", "正常"
        upcoming_ship = any(
            0 <= ((_date(x.get("出荷日")) or current) - current).days <= 7
            for x in shipments if _date(x.get("出荷日"))
        )
        if overdue_rows:
            risk_level, risk_code, risk_label = "high", "ar_overdue", "应收逾期"
        elif upcoming_ship and pending_serial:
            risk_level, risk_code, risk_label = "high", "serial_before_ship", "临近发货仍待编号"
        elif due_soon_rows:
            risk_level, risk_code, risk_label = "medium", "ar_due_soon", "应收临近"
        elif pending_serial:
            risk_level, risk_code, risk_label = "medium", "missing_serial", f"{pending_serial}台待编号"
        elif outstanding > 0.01:
            risk_level, risk_code, risk_label = "medium", "ar_open", "应收跟进"

        quote_id = _text(order.get("来源报价ID"))
        quote_no = _text(order.get("来源报价单号"))
        result.append({
            "job": job, "customer": _text(order.get("客户")),
            "salesperson": _text(order.get("担当者")),
            "content": _text(order.get("订单内容")) or _text(order.get("设备型号")),
            "currency": currency, "source_quote_id": quote_id, "source_quote_no": quote_no,
            "device_count": len(devices), "shipped_count": len(shipped_devices),
            "scheduled_count": len(scheduled_devices), "accepted_count": len(accepted_devices),
            "pending_serial": pending_serial, "total_untaxed": total_untaxed,
            "expected_gross": expected_gross, "invoice_amount": invoice_amount,
            "paid_amount": paid_amount, "outstanding": outstanding,
            "ship_pct": ship_pct, "accept_pct": accept_pct, "invoice_pct": invoice_pct,
            "paid_pct": paid_pct, "stage": stage, "stage_index": stage_index,
            "risk_level": risk_level, "risk_code": risk_code, "risk_label": risk_label,
            "next_ship_date": min(
                (_text(x.get("出荷日")) for x in shipments
                 if _date(x.get("出荷日")) and _date(x.get("出荷日")) >= current),
                default="",
            ),
            "document_flow": _document_flow(job, quote_id, quote_no, shipments, invoices, payments),
        })
    return sorted(result, key=lambda row: row["job"], reverse=True)


def build_finance(data: dict[str, list[dict[str, Any]]], rules=None,
                  today: date | str | None = None) -> dict[str, Any]:
    current = _today(today)
    derived = core.derive(copy.deepcopy(data))
    customer_by_job = {_text(x.get("JOB No")): _text(x.get("客户")) for x in derived["合同订单"]}
    currency_by_job = {_text(x.get("JOB No")): _text(x.get("币种")) or "RMB" for x in derived["合同订单"]}
    rows, aging, forecast = [], defaultdict(list), defaultdict(list)

    for raw in core.unpaid_report_rows(derived, rules):
        job = _text(raw.get("JOB No"))
        due = _date(raw.get("预定回收日期"))
        row = {
            **raw, "job": job,
            "customer": customer_by_job.get(job, _text(raw.get("客户"))),
            "currency": currency_by_job.get(job, "RMB"),
            "amount": round(_num(raw.get("未回收金额")), 2),
            "due_date": due.isoformat() if due else "",
            "days": (due - current).days if due else None,
        }
        rows.append(row)
        if due is None:
            aging["待确认"].append(row); forecast["待确认"].append(row); continue
        days = (due - current).days
        if days >= 0:
            aging["未到期"].append(row)
            forecast["0-7天" if days <= 7 else "8-30天" if days <= 30 else "31-60天" if days <= 60 else "61-90天" if days <= 90 else "90天以上"].append(row)
        else:
            overdue = abs(days)
            aging["逾期0-30天" if overdue <= 30 else "逾期31-60天" if overdue <= 60 else "逾期61-90天" if overdue <= 90 else "逾期90天以上"].append(row)
            forecast["已逾期"].append(row)

    def bucket_payload(source, names):
        return [{"bucket": name, "count": len(source.get(name, [])),
                 "amounts": _group_amount(source.get(name, [])), "rows": source.get(name, [])}
                for name in names]

    return {
        "rows": rows,
        "aging": bucket_payload(aging, ("未到期", "逾期0-30天", "逾期31-60天", "逾期61-90天", "逾期90天以上", "待确认")),
        "forecast": bucket_payload(forecast, ("已逾期", "0-7天", "8-30天", "31-60天", "61-90天", "90天以上", "待确认")),
        "open_amounts": _group_amount(rows),
        "overdue_amounts": _group_amount(forecast.get("已逾期", [])),
    }


def build_actions(data: dict[str, list[dict[str, Any]]], rules=None,
                  today: date | str | None = None,
                  work_items: list[dict[str, Any]] | None = None) -> list[dict[str, Any]]:
    current = _today(today)
    snapshots = {row["job"]: row for row in build_order_snapshots(data, rules, current)}
    finance = build_finance(data, rules, current)
    actions = []

    def add(job, kind, priority, title, detail="", due_date="", amount=0, currency="RMB", source_id=""):
        actions.append({
            "action_id": _stable_id(job, kind, due_date, source_id, title),
            "job": job, "customer": snapshots.get(job, {}).get("customer", ""),
            "type": kind, "priority": priority, "title": title, "detail": detail,
            "due_date": due_date, "amount": round(_num(amount), 2),
            "currency": currency or snapshots.get(job, {}).get("currency", "RMB"),
            "source_id": source_id, "manual": False,
        })

    for row in finance["rows"]:
        days, job = row["days"], row["job"]
        source = _text(row.get("批次")) + ":" + _text(row.get("款类"))
        if days is None:
            add(job, "ar_confirm", "medium", "应收日期待确认", _text(row.get("未回收原因")), "", row["amount"], row["currency"], source)
        elif days < 0:
            add(job, "ar_overdue", "high", f"应收已逾期 {abs(days)} 天", _text(row.get("款类")), row["due_date"], row["amount"], row["currency"], source)
        elif days <= 7:
            add(job, "ar_due", "high", f"应收将在 {days} 天内到期", _text(row.get("款类")), row["due_date"], row["amount"], row["currency"], source)
        elif days <= 30:
            add(job, "ar_due", "medium", f"应收将在 {days} 天内到期", _text(row.get("款类")), row["due_date"], row["amount"], row["currency"], source)

    derived = core.derive(copy.deepcopy(data))
    devices_by_job, shipments_by_job = defaultdict(list), defaultdict(list)
    for row in derived["设备台账"]: devices_by_job[_text(row.get("JOB No"))].append(row)
    for row in derived["发货批次"]: shipments_by_job[_text(row.get("JOB No"))].append(row)

    for job, snap in snapshots.items():
        for shipment in shipments_by_job.get(job, []):
            ship_date = _date(shipment.get("出荷日"))
            if not ship_date or ship_date < current: continue
            days = (ship_date - current).days
            if days <= 7:
                batch = _text(shipment.get("发货批次"))
                missing = sum(1 for d in devices_by_job.get(job, [])
                              if _text(d.get("发货批次")) == batch and not _text(d.get("製造番号")))
                add(job, "shipment_due", "high" if days <= 2 else "medium",
                    f"发货批次 {batch} 将在 {days} 天内出货",
                    f"{missing} 台待编号" if missing else "制造番号已齐",
                    ship_date.isoformat(), source_id=batch)
        shipped_dates = [_date(x.get("出荷日")) for x in shipments_by_job.get(job, [])
                         if _date(x.get("出荷日")) and _date(x.get("出荷日")) <= current]
        if shipped_dates and snap["accepted_count"] < snap["shipped_count"]:
            days_since = (current - max(shipped_dates)).days
            if days_since >= 14:
                add(job, "acceptance_wait", "high" if days_since >= 30 else "medium",
                    f"已发货 {days_since} 天仍待验收",
                    f"{snap['shipped_count'] - snap['accepted_count']} 台待验收")
        if snap["pending_serial"] and not any(a["job"] == job and a["type"] == "shipment_due" for a in actions):
            add(job, "missing_serial", "low", f"{snap['pending_serial']} 台设备待补制造番号")

    for item in work_items or []:
        if _text(item.get("status")) == "done": continue
        due = _date(item.get("due_date"))
        priority = _text(item.get("priority")) or "medium"
        actions.append({
            "action_id": f"work:{_text(item.get('item_id'))}",
            "job": _text(item.get("job_no")),
            "customer": snapshots.get(_text(item.get("job_no")), {}).get("customer", ""),
            "type": "work_item", "priority": priority if priority in PRIORITY_ORDER else "medium",
            "title": _text(item.get("title")), "detail": _text(item.get("note")),
            "due_date": due.isoformat() if due else "", "amount": 0,
            "currency": snapshots.get(_text(item.get("job_no")), {}).get("currency", "RMB"),
            "source_id": _text(item.get("item_id")), "manual": True, "work_item": item,
        })

    actions.sort(key=lambda row: (PRIORITY_ORDER.get(row["priority"], 9), row["due_date"] or "9999-12-31", row["job"], row["title"]))
    return actions


def build_workspace(data: dict[str, list[dict[str, Any]]], rules=None,
                    today: date | str | None = None,
                    work_items: list[dict[str, Any]] | None = None) -> dict[str, Any]:
    current = _today(today)
    orders = build_order_snapshots(data, rules, current)
    finance = build_finance(data, rules, current)
    actions = build_actions(data, rules, current, work_items)
    return {
        "as_of": current.isoformat(), "orders": orders, "actions": actions, "finance": finance,
        "stage_counts": [{"stage": label, "count": sum(1 for row in orders if row["stage"] == label)} for label in STAGE_LABELS],
        "kpis": {
            "orders": len(orders), "action_count": len(actions),
            "high_action_count": sum(1 for x in actions if x["priority"] == "high"),
            "pending_ship_orders": sum(1 for x in orders if x["device_count"] and x["ship_pct"] < 100),
            "pending_accept_devices": sum(max(0, x["shipped_count"] - x["accepted_count"]) for x in orders),
            "pending_serial_devices": sum(x["pending_serial"] for x in orders),
            "open_amounts": finance["open_amounts"], "overdue_amounts": finance["overdue_amounts"],
        },
    }
