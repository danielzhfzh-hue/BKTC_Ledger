#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Generate the Japanese 装置搬入依頼書 used by the shipment workflow.

The supplied workbook is an old BIFF ``.xls`` file with embedded stamp objects
and unreliable pagination.  This module rebuilds the printable form as a clean
A4 ``.xlsx``.  The first worksheet contains one complete equipment list, with
the request details placed below it so readers do not have to reconcile a
summary table with a second list.
"""

from __future__ import annotations

import os
import tempfile
from datetime import date, datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.page import PageMargins


THIN = Side(style="thin", color="000000")
GRID = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
LABEL_FILL = PatternFill("solid", fgColor="E0E0E0")
FONT = Font(name="ＭＳ Ｐ明朝", size=12, color="000000")
SMALL_FONT = Font(name="ＭＳ Ｐ明朝", size=10, color="000000")
TITLE_FONT = Font(name="ＭＳ Ｐ明朝", size=18, color="000000")
DETAIL_FONT = Font(name="ＭＳ Ｐ明朝", size=10, color="000000")

NOTE_LINE_WIDTH = 40


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
    note_lines = []
    if ship_from:
        note_lines.append(f"※出荷元：{ship_from}")
    if shipment.get("出荷日"):
        note_lines.append(f"※出荷日：{shipment.get('出荷日')}")
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
        "system_notes": "\n".join(note_lines),
        "notes": "",
        "manager": "",
        "issuer": "",
    }


def _wrap_note_lines(values, width=NOTE_LINE_WIDTH):
    """Wrap note lines to a conservative width for the merged A4 note cells."""
    lines = []
    for value in values:
        text = "" if value is None else str(value)
        if not text:
            lines.append("")
            continue
        while len(text) > width:
            lines.append(text[:width])
            text = text[width:]
        lines.append(text)
    return lines


def _write_device_table(ws, header_row, devices):
    """Write the single complete equipment table on the first worksheet."""
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    headers = ((1, 2, "No."), (3, 5, "機種"), (6, 9, "機番"),
               (10, 13, "制番"), (14, 17, "PO No"))
    _row_style(ws, header_row, height=24, border=GRID, fill=LABEL_FILL,
               font=DETAIL_FONT)
    for start_col, end_col, label in headers:
        _merge(ws, header_row, start_col, header_row, end_col, label,
               font=DETAIL_FONT, fill=LABEL_FILL, border=GRID, alignment=center)

    for offset, device in enumerate(devices):
        row = header_row + 1 + offset
        _row_style(ws, row, height=22, border=GRID, font=DETAIL_FONT)
        _merge(ws, row, 1, row, 2, offset + 1,
               font=DETAIL_FONT, border=GRID, alignment=center)
        _merge(ws, row, 6, row, 9, device.get("machine_no", ""),
               font=DETAIL_FONT, border=GRID, alignment=center)
        _merge(ws, row, 10, row, 13, device.get("serial_no", ""),
               font=DETAIL_FONT, border=GRID, alignment=center)

    for start_col, end_col, key in ((3, 5, "model"), (14, 17, "po_no")):
        group_start = 0
        for group_end in range(1, len(devices) + 1):
            at_end = group_end == len(devices)
            same_as_next = (not at_end
                            and devices[group_end].get(key, "") == devices[group_start].get(key, ""))
            if at_end or not same_as_next:
                first_row = header_row + 1 + group_start
                last_row = header_row + group_end
                _merge(ws, first_row, start_col, last_row, end_col,
                       devices[group_start].get(key, ""),
                       font=DETAIL_FONT, border=GRID, alignment=center)
                group_start = group_end
    return header_row + len(devices)


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

    devices = form.get("devices") or []
    _field_label(ws, 11, 1, "設備明細")
    _merge(ws, 11, 3, 11, 17, f"全{len(devices)}台", font=FONT,
           border=GRID, alignment=left)
    device_header_row = 12
    device_end_row = _write_device_table(ws, device_header_row, devices)

    # Keep the administrative part of the request directly below the one
    # authoritative equipment list.  There is no summary table or appendix.
    row = device_end_row + 1
    _field_label(ws, row, 2, "客先装置ID No.")
    _merge(ws, row, 3, row, 4, ("☑ 有り" if form.get("has_customer_id") else "☐ 有り"),
           font=FONT, border=GRID, alignment=center)
    _merge(ws, row, 5, row, 15, ("☐ 無し" if form.get("has_customer_id") else "☑ 無し"),
           font=FONT, border=GRID, alignment=center)
    _merge(ws, row, 16, row, 17, "", font=FONT, border=GRID, alignment=center)

    row += 1
    _field_label(ws, row, 3, "搬入先名")
    _merge(ws, row, 3, row, 17, form.get("customer", ""), font=FONT, border=GRID, alignment=left)
    row += 1
    _field_label(ws, row, 4, "搬入先住所", row + 1)
    _merge(ws, row, 3, row + 1, 17, form.get("address", ""), font=FONT, border=GRID, alignment=left)
    row += 2
    _field_label(ws, row, 5, "搬入連絡先")
    _merge(ws, row, 3, row, 17, form.get("contact", ""), font=FONT, border=GRID, alignment=left)
    row += 1
    _field_label(ws, row, 6, "搬入日時", row + 1)
    _merge(ws, row, 3, row, 8, form.get("move_in_date", ""), font=FONT, border=GRID,
           alignment=center, number_format="date")
    _merge(ws, row, 9, row, 10, "", font=FONT, border=GRID, alignment=center)
    _merge(ws, row, 11, row, 14, form.get("move_in_time", ""), font=FONT, border=GRID, alignment=center)
    _merge(ws, row + 1, 3, row + 1, 6, "AM　・　PM", font=FONT, border=GRID, alignment=center)
    _merge(ws, row + 1, 7, row + 1, 10, "", font=FONT, border=GRID, alignment=center)
    _merge(ws, row + 1, 11, row + 1, 17, "", font=FONT, border=GRID, alignment=center)
    row += 2
    _field_label(ws, row, 7, "搬入車両", row + 3)
    vehicles = ["平ボディー", "ユニック車", "パワーゲート車", "その他"]
    chosen = str(form.get("vehicle_type") or "")
    for vehicle_row, label in enumerate(vehicles, row):
        _merge(ws, vehicle_row, 4, vehicle_row, 7, ("☑ " if chosen == label else "☐ ") + label,
               font=FONT, border=GRID, alignment=left)
        if vehicle_row < row + 3:
            _merge(ws, vehicle_row, 8, vehicle_row, 9, form.get("vehicle_tonnage", ""),
                   font=FONT, border=GRID, alignment=center)
            _set(ws, f"J{vehicle_row}", "t", font=FONT, border=GRID, alignment=center)
        else:
            _merge(ws, vehicle_row, 8, vehicle_row, 9, "※１", font=SMALL_FONT,
                   border=GRID, alignment=center)
            _merge(ws, vehicle_row, 11, vehicle_row, 17, form.get("vehicle_note", ""),
                   font=SMALL_FONT, border=GRID, alignment=left)
    row += 4

    system_notes = str(form.get("system_notes") or "").splitlines()
    if not system_notes:
        # Keep direct callers using the pre-v1.9 form shape readable.
        if form.get("ship_from"):
            system_notes.append(f"※出荷元：{form.get('ship_from')}")
        if form.get("ship_date"):
            system_notes.append(f"※出荷日：{form.get('ship_date')}")
    notes = _wrap_note_lines(system_notes + str(form.get("notes") or "").splitlines())
    note_start = row
    note_rows = max(8, len(notes), 1)
    note_end = note_start + note_rows - 1
    _field_label(ws, note_start, 8, "特記事項", note_end)
    for note_row in range(note_start, note_end + 1):
        _row_style(ws, note_row, start=3, end=17, height=21, border=GRID)
        _merge(ws, note_row, 3, note_row, 17,
               notes[note_row - note_start] if note_row - note_start < len(notes) else "",
               font=FONT, border=GRID, alignment=wrap)

    last_row = note_end
    ws.print_area = f"A1:Q{last_row}"
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.25, right=0.25, top=0.25, bottom=0.25,
                                  header=0.1, footer=0.1)
    ws.sheet_properties.pageSetUpPr.autoPageBreaks = True
    # Repeat the equipment header only when the equipment table itself spills
    # past the first printed page.  If the complete list ends on page one,
    # later pages contain only request details and should not show a duplicate
    # device header.
    if device_end_row > 35:
        ws.print_title_rows = "1:12"
    ws.freeze_panes = "A13"
    ws.oddFooter.center.text = "搬入依頼書　第 &P 页"
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = "auto"
    wb.properties.title = f"搬入依頼書 {form.get('job_no', '')} {form.get('batch', '')}"
    wb.properties.creator = "上海康肯销售订单管理系统"

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
