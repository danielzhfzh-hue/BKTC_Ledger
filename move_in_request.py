#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Generate the Japanese 装置搬入依頼書 used by the shipment workflow.

The supplied workbook is an old BIFF ``.xls`` file with embedded stamp objects
and unreliable pagination.  This module deliberately rebuilds the printable
form as a clean, one-page A4 ``.xlsx`` and keeps the business data mapping in a
small, testable function.
"""

from __future__ import annotations

import os
import tempfile
from datetime import date, datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.page import PageMargins


THIN = Side(style="thin", color="000000")
DOUBLE = Side(style="double", color="000000")
GRID = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
LABEL_FILL = PatternFill("solid", fgColor="E0E0E0")
FONT = Font(name="ＭＳ Ｐ明朝", size=12, color="000000")
SMALL_FONT = Font(name="ＭＳ Ｐ明朝", size=10, color="000000")
TITLE_FONT = Font(name="ＭＳ Ｐ明朝", size=18, color="000000")


def _text(value):
    raw = "" if value is None else str(value)
    return "'" + raw if raw.startswith(("=", "+", "-", "@")) else raw


def _date_value(value):
    if isinstance(value, (datetime, date)):
        return value.date() if isinstance(value, datetime) else value
    try:
        return date.fromisoformat(str(value))
    except (TypeError, ValueError):
        return _text(value)


def _set(ws, cell, value="", *, font=None, alignment=None, border=None,
         fill=None, number_format=None):
    target = ws[cell]
    target.value = _date_value(value) if number_format == "date" else (
        _text(value) if isinstance(value, str) else value
    )
    target.font = font or FONT
    target.alignment = alignment or Alignment(vertical="center")
    if border:
        target.border = border
    if fill:
        target.fill = fill
    if number_format == "date":
        target.number_format = "[$-804]yyyy\\年m\\月d\\日"
    elif number_format:
        target.number_format = number_format
    return target


def _merge(ws, start_row, start_col, end_row, end_col, value="", **style):
    ws.merge_cells(start_row=start_row, start_column=start_col,
                   end_row=end_row, end_column=end_col)
    border = style.get("border")
    fill = style.get("fill")
    for row in ws.iter_rows(min_row=start_row, max_row=end_row,
                            min_col=start_col, max_col=end_col):
        for cell in row:
            if border:
                cell.border = border
            if fill:
                cell.fill = fill
            cell.font = style.get("font") or FONT
            cell.alignment = style.get("alignment") or Alignment(vertical="center")
    return _set(ws, f"{chr(64 + start_col)}{start_row}", value, **style)


def _row_style(ws, row, start=1, end=17, *, height=21, border=GRID,
               fill=None, font=None):
    ws.row_dimensions[row].height = height
    for col in range(start, end + 1):
        cell = ws.cell(row, col)
        cell.font = font or FONT
        cell.alignment = Alignment(vertical="center")
        if border:
            cell.border = border
        if fill:
            cell.fill = fill


def _field_label(ws, row, number, label, end_row=None):
    end_row = end_row or row
    _merge(ws, row, 1, end_row, 1, number, font=FONT, fill=LABEL_FILL,
           border=GRID, alignment=Alignment(horizontal="center", vertical="center"))
    _merge(ws, row, 2, end_row, 2, label, font=FONT, fill=LABEL_FILL,
           border=GRID, alignment=Alignment(horizontal="center", vertical="center"))


def shipment_defaults(data, shipment_id=None, job_no=None, batch=None,
                      issue_date=None):
    """Build form defaults from an authoritative database-shaped data object."""
    shipments = data.get("发货批次", [])
    shipment = None
    if shipment_id:
        shipment = next((r for r in shipments
                         if str(r.get("记录ID", "")) == str(shipment_id)), None)
    if shipment is None and job_no is not None and batch is not None:
        shipment = next((r for r in shipments
                         if str(r.get("JOB No", "")) == str(job_no)
                         and str(r.get("发货批次", "")) == str(batch)), None)
    if shipment is None:
        raise ValueError("找不到所选发货批次，请刷新数据后重试")

    job = str(shipment.get("JOB No") or "")
    batch_name = str(shipment.get("发货批次") or "")
    devices = [r for r in data.get("设备台账", [])
               if str(r.get("JOB No") or "") == job
               and str(r.get("发货批次") or "") == batch_name]
    if not devices:
        raise ValueError(f"批次 {job} / {batch_name} 没有关联设备，无法生成搬入依頼書")
    contract = next((r for r in data.get("合同订单", [])
                     if str(r.get("JOB No") or "") == job), {})
    customer = str(contract.get("客户") or "")
    address = str(contract.get("送货地点") or "")
    ship_from = str(contract.get("发货地点") or "")
    models = sorted({str(r.get("设备型号") or "") for r in devices if r.get("设备型号")})
    machine_numbers = [str(r.get("機番") or "") for r in devices]
    serials = [str(r.get("製造番号") or "") for r in devices]
    po_groups = {}
    for row in devices:
        po = str(row.get("PO No") or "")
        machine = str(row.get("機番") or "")
        if po:
            po_groups.setdefault(po, []).append(machine or str(row.get("製造番号") or ""))
    note_lines = []
    if ship_from:
        note_lines.append(f"※出荷元：{ship_from}")
    if shipment.get("出荷日"):
        note_lines.append(f"※出荷日：{shipment.get('出荷日')}")
    for po, machines in po_groups.items():
        values = "，".join(x for x in machines if x)
        note_lines.append(f"機番{values}はPO:{po}表記お願いいたします")
    return {
        "shipment_id": str(shipment.get("记录ID") or ""),
        "job_no": job,
        "batch": batch_name,
        "customer": customer,
        "address": address,
        "ship_from": ship_from,
        "ship_date": str(shipment.get("出荷日") or ""),
        "devices": [
            {"model": str(r.get("设备型号") or ""),
             "machine_no": str(r.get("機番") or ""),
             "serial_no": str(r.get("製造番号") or ""),
             "po_no": str(r.get("PO No") or "")}
            for r in devices
        ],
        "models": models,
        "has_customer_id": all(machine_numbers),
        "issue_date": str(issue_date or date.today().isoformat()),
        "move_in_date": "",
        "move_in_time": "",
        "contact": "",
        "vehicle_type": "",
        "vehicle_tonnage": "",
        "vehicle_note": "",
        "department": "本社工場　/　物流購買部　/　技術設計部　/　技術サービス部",
        "other_department": "平湖康肯",
        "notes": "\n".join(note_lines),
        "manager": "",
        "issuer": "",
    }


def _device_rows(devices, limit=4):
    rows = []
    for d in devices[:limit]:
        rows.append((d.get("model", ""), d.get("machine_no", ""), d.get("serial_no", "")))
    return rows


def build_workbook(form):
    """Build the normalized KE-403-style workbook from a validated form."""
    wb = Workbook()
    ws = wb.active
    ws.title = "搬入依頼書"
    ws.sheet_view.showGridLines = False
    widths = {"A": 5.44, "B": 22.81, "C": 4.08, "D": 7.9, "E": 2.81,
              "F": 4.08, "G": 4.35, "H": 3.17, "I": 3.99, "J": 3.08,
              "K": 4.99, "L": 2.9, "M": 1.9, "N": 6.99, "O": 16.62,
              "P": 9.35, "Q": 8.99}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    for row, height in {1: 20.25, 2: 20.25, 3: 20.25, 4: 13.5, 5: 25.5,
                        6: 14.25, 7: 20.25, 8: 20.25, 9: 20.25,
                        10: 9, 11: 19.5, 12: 15.75}.items():
        ws.row_dimensions[row].height = height

    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")
    wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)
    center_wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)

    _merge(ws, 1, 1, 1, 2, form.get("issue_date", ""), font=FONT,
           alignment=center, number_format="date")
    _merge(ws, 1, 14, 1, 17, "[KE-403]", font=SMALL_FONT, alignment=right_align())
    _merge(ws, 2, 1, 2, 17, form.get("department", ""), font=FONT, alignment=left)
    other = form.get("other_department", "")
    _merge(ws, 3, 1, 3, 17, f"他（　　　{other}　　　）" if other else "他（　　　　　　　　　）",
           font=FONT, alignment=left)
    _merge(ws, 5, 1, 5, 17, "装 置 搬 入 依 頼 書", font=TITLE_FONT, alignment=center)
    for cell, label, value in (("P6", "所属長", form.get("manager", "")),
                               ("Q6", "発行者", form.get("issuer", ""))):
        _set(ws, cell, label, font=FONT, border=GRID, alignment=center)
        col = cell[0]
        _merge(ws, 7, ord(col) - 64, 9, ord(col) - 64, value, font=FONT,
               border=GRID, alignment=center)

    _row_style(ws, 11, fill=LABEL_FILL, border=GRID)
    _merge(ws, 11, 2, 11, 2, "項　　　目", font=FONT, fill=LABEL_FILL,
           border=Border(left=DOUBLE, right=THIN, top=THIN, bottom=DOUBLE), alignment=center)
    _merge(ws, 11, 3, 11, 17, "内　　　容", font=FONT, border=GRID, alignment=center)
    _field_label(ws, 12, 1, "機　　　種")
    _merge(ws, 12, 3, 12, 11, "機　　　番", font=FONT, fill=LABEL_FILL, border=GRID, alignment=center)
    _merge(ws, 12, 12, 12, 17, "制      番", font=FONT, fill=LABEL_FILL, border=GRID, alignment=center)

    devices = form.get("devices") or []
    rows = _device_rows(devices)
    while len(rows) < 4:
        rows.append(("", "", ""))
    _merge(ws, 13, 1, 16, 1, 2, font=FONT, fill=LABEL_FILL, border=GRID, alignment=center)
    for index, (model, machine, serial) in enumerate(rows, 13):
        _row_style(ws, index, height=(46 if index == 13 else 43), border=GRID)
        _set(ws, f"B{index}", model, font=FONT, border=GRID, alignment=center_wrap)
        _merge(ws, index, 3, index, 11, machine, font=FONT, border=GRID, alignment=center_wrap)
        _merge(ws, index, 12, index, 17, serial, font=FONT, border=GRID, alignment=center_wrap)

    row = 17
    _field_label(ws, row, 3, "客先装置ID No.")
    _merge(ws, row, 3, row, 4, ("☑ 有り" if form.get("has_customer_id") else "☐ 有り"),
           font=FONT, border=GRID, alignment=center)
    _merge(ws, row, 5, row, 15, ("☐ 無し" if form.get("has_customer_id") else "☑ 無し"),
           font=FONT, border=GRID, alignment=center)
    _merge(ws, row, 16, row, 17, "", font=FONT, border=GRID, alignment=center)

    _field_label(ws, 18, 4, "搬入先名")
    _merge(ws, 18, 3, 18, 17, form.get("customer", ""), font=FONT, border=GRID, alignment=left)
    _field_label(ws, 19, 5, "搬入先住所", 20)
    _merge(ws, 19, 3, 20, 17, form.get("address", ""), font=FONT, border=GRID, alignment=left)
    _field_label(ws, 21, 6, "搬入連絡先")
    _merge(ws, 21, 3, 21, 17, form.get("contact", ""), font=FONT, border=GRID, alignment=left)
    _field_label(ws, 22, 7, "搬入日時", 23)
    _merge(ws, 22, 3, 22, 8, form.get("move_in_date", ""), font=FONT, border=GRID,
           alignment=center, number_format="date")
    _merge(ws, 22, 9, 22, 10, "", font=FONT, border=GRID, alignment=center)
    _merge(ws, 22, 11, 22, 14, form.get("move_in_time", ""), font=FONT, border=GRID, alignment=center)
    _merge(ws, 23, 3, 23, 6, "AM　・　PM", font=FONT, border=GRID, alignment=center)
    _merge(ws, 23, 7, 23, 10, "", font=FONT, border=GRID, alignment=center)
    _merge(ws, 23, 11, 23, 17, "", font=FONT, border=GRID, alignment=center)
    _field_label(ws, 24, 8, "搬入車両", 27)
    vehicles = ["平ボディー", "ユニック車", "パワーゲート車", "その他"]
    chosen = str(form.get("vehicle_type") or "")
    for vehicle_row, label in enumerate(vehicles, 24):
        _merge(ws, vehicle_row, 4, vehicle_row, 7, ("☑ " if chosen == label else "☐ ") + label,
               font=FONT, border=GRID, alignment=left)
        if vehicle_row < 27:
            _merge(ws, vehicle_row, 8, vehicle_row, 9, form.get("vehicle_tonnage", ""),
                   font=FONT, border=GRID, alignment=center)
            _set(ws, f"J{vehicle_row}", "t", font=FONT, border=GRID, alignment=center)
        else:
            _merge(ws, vehicle_row, 8, vehicle_row, 9, "※１", font=SMALL_FONT,
                   border=GRID, alignment=center)
            _merge(ws, vehicle_row, 11, vehicle_row, 17, form.get("vehicle_note", ""),
                   font=SMALL_FONT, border=GRID, alignment=left)
    _field_label(ws, 28, 9, "特記事項", 35)
    notes = str(form.get("notes") or "").splitlines()
    if len(devices) > 4:
        notes.insert(0, f"※設備明細：全{len(devices)}台。4台を超える設備は別紙「設備明細」を参照")
    notes_overflow = len(notes) > 8
    if notes_overflow:
        notes = notes[:7] + ["\n".join(notes[7:])]
    for note_row in range(28, 36):
        _row_style(ws, note_row, start=3, end=17,
                   height=(54 if note_row == 35 and notes_overflow else 21), border=GRID)
        _merge(ws, note_row, 3, note_row, 17,
               notes[note_row - 28] if note_row - 28 < len(notes) else "",
               font=FONT, border=GRID, alignment=wrap)

    ws.print_area = "A1:Q35"
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.25, right=0.25, top=0.25, bottom=0.25,
                                  header=0.1, footer=0.1)
    ws.sheet_properties.pageSetUpPr.autoPageBreaks = False
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = "auto"
    wb.properties.title = f"搬入依頼書 {form.get('job_no', '')} {form.get('batch', '')}"
    wb.properties.creator = "上海康肯销售订单管理系统"

    if len(devices) > 4:
        appendix = wb.create_sheet("設備明細")
        appendix.sheet_view.showGridLines = False
        appendix.append(["JOB No", "发货批次", "機種", "機番", "制番", "PO No"])
        for cell in appendix[1]:
            cell.font = Font(name="ＭＳ Ｐ明朝", size=11, bold=True)
            cell.fill = LABEL_FILL
            cell.border = GRID
        for d in devices:
            appendix.append([form.get("job_no", ""), form.get("batch", ""), d.get("model", ""),
                             d.get("machine_no", ""), d.get("serial_no", ""), d.get("po_no", "")])
        for row_cells in appendix.iter_rows():
            for cell in row_cells:
                cell.font = SMALL_FONT
                cell.border = GRID
                cell.alignment = Alignment(vertical="center", wrap_text=True)
        for col, width in {"A": 16, "B": 12, "C": 28, "D": 18, "E": 20, "F": 18}.items():
            appendix.column_dimensions[col].width = width
        appendix.freeze_panes = "A2"
        appendix.auto_filter.ref = f"A1:F{appendix.max_row}"
        appendix.print_area = f"A1:F{appendix.max_row}"
        appendix.page_setup.orientation = appendix.ORIENTATION_LANDSCAPE
        appendix.page_setup.paperSize = appendix.PAPERSIZE_A4
        appendix.page_setup.fitToWidth = 1
        appendix.sheet_properties.pageSetUpPr.fitToPage = True

    return wb


def right_align():
    return Alignment(horizontal="right", vertical="center")


def validate_form(form):
    errors = []
    if not str(form.get("job_no") or "").strip():
        errors.append("缺少 JOB No")
    if not str(form.get("batch") or "").strip():
        errors.append("缺少发货批次")
    if not form.get("devices"):
        errors.append("批次至少需要一台设备")
    if not str(form.get("customer") or "").strip():
        errors.append("搬入先名不能为空")
    if not str(form.get("address") or "").strip():
        errors.append("搬入先地址不能为空")
    if not str(form.get("move_in_date") or "").strip():
        errors.append("搬入日期不能为空")
    return errors


def export_xlsx(form, path):
    errors = validate_form(form)
    if errors:
        raise ValueError("；".join(errors))
    path = os.path.abspath(os.fspath(path))
    os.makedirs(os.path.dirname(path), exist_ok=True)
    workbook = build_workbook(form)
    fd, temporary = tempfile.mkstemp(prefix=".move-in-", suffix=".xlsx",
                                      dir=os.path.dirname(path))
    os.close(fd)
    try:
        workbook.save(temporary)
        os.replace(temporary, path)
    finally:
        if os.path.exists(temporary):
            os.remove(temporary)
    return path
