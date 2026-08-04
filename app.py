#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""BKTC 台账维护工具（pywebview 桌面壳，macOS / Windows 通用）。"""
import argparse
import os
import shutil
import subprocess
import sys
import time

APP_DIR = sys._MEIPASS if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, APP_DIR)

import core  # noqa: E402
import webview  # noqa: E402

DEFAULT_XLSX = r"/Users/danielzhu/projects/订单整理/BKTC上海POU营业管理表.xlsx"
DEFAULT_STORE = r"/Users/danielzhu/projects/订单整理/BKTC上海POU营业管理表.records.json"
FALLBACK_EXPORT = "/tmp/export_20260803_fixed"


def schema_for_js():
    out = {}
    for t in core.TABLES:
        derived = set(core.DERIVED.get(t, []))
        out[t] = [list(f) + [f[0] in derived] for f in core.SCHEMA[t]]
    return out


class Api:
    def __init__(self, xlsx_path, store_path):
        self.xlsx_path = xlsx_path
        self.store_path = store_path

    def _init_store(self):
        if os.path.isdir(FALLBACK_EXPORT):
            core.save_store(self.store_path, core.import_from_export_dir(FALLBACK_EXPORT))
            return
        core.save_store(self.store_path, core.empty_data())

    def load_state(self, store=None, xlsx=None):
        if store:
            self.store_path = store
        if xlsx:
            self.xlsx_path = xlsx
        if not os.path.exists(self.store_path):
            self._init_store()
        data = core.load_store(self.store_path)
        return {"store_path": self.store_path, "xlsx_path": self.xlsx_path,
                "data": data, "schema": schema_for_js(),
                "rules": core.get_rules(self.store_path)}

    def save_data(self, data, rules=None):
        core.save_store(self.store_path, data, rules)
        return {"ok": True, "path": self.store_path}

    def generate(self, data, rules=None):
        if not self.xlsx_path or not os.path.isdir(os.path.dirname(self.xlsx_path)):
            raise RuntimeError("请先选择台账 Excel 文件")
        if os.path.exists(self.xlsx_path):
            root, ext = os.path.splitext(self.xlsx_path)
            bak = f"{root}_备份_{time.strftime('%Y%m%d_%H%M%S')}{ext}"
            shutil.copy2(self.xlsx_path, bak)
        core.save_store(self.store_path, data, rules)
        out, issues, counts = core.generate_xlsx(data, self.xlsx_path, rules)
        return {"ok": True, "out": out, "issues": issues, "counts": counts}

    def pick_store(self):
        w = webview.windows[0]
        res = w.create_file_dialog(webview.OPEN_DIALOG,
                                   file_types=("JSON (*.json)", "All files (*.*)"))
        return res[0] if res else None

    def pick_xlsx(self):
        w = webview.windows[0]
        res = w.create_file_dialog(webview.OPEN_DIALOG,
                                   file_types=("Excel (*.xlsx)", "All files (*.*)"))
        return res[0] if res else None

    def pick_dir(self):
        w = webview.windows[0]
        res = w.create_file_dialog(webview.FOLDER_DIALOG)
        return res[0] if res else None

    def import_from_dir(self, path):
        data = core.import_from_export_dir(path)
        core.save_store(self.store_path, data, core.get_rules(self.store_path))
        return {"ok": True, "data": data, "path": self.store_path}

    def sync_from_excel(self):
        if not os.path.exists(self.xlsx_path):
            raise RuntimeError("请先选择台账 Excel 文件")
        store, changed = core.reconcile_store_with_excel(self.store_path, self.xlsx_path)
        return {"ok": True, "data": store, "changed": changed}

    def open_path(self, path):
        if sys.platform == "darwin":
            subprocess.Popen(["open", path])
        elif os.name == "nt":
            os.startfile(path)  # noqa: S606
        else:
            subprocess.Popen(["xdg-open", path])
        return True


def main():
    ap = argparse.ArgumentParser(description="BKTC 台账维护工具")
    ap.add_argument("--xlsx", default=os.environ.get("BKTC_XLSX", DEFAULT_XLSX),
                    help="台账 Excel 路径（可用 --xlsx 或界面选择）")
    ap.add_argument("--store", default=os.environ.get("BKTC_STORE", DEFAULT_STORE),
                    help="records.json 数据文件路径")
    ap.add_argument("--debug", action="store_true")
    args = ap.parse_args()

    api = Api(args.xlsx, args.store)
    webview.create_window(
        "BKTC 台账维护工具",
        os.path.join(APP_DIR, "ui", "index.html"),
        js_api=api,
        width=1440,
        height=900,
        min_size=(1100, 680),
        text_select=True,
    )
    webview.start(debug=args.debug)


if __name__ == "__main__":
    main()
