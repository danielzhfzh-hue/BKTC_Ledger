#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""BKTC 台账维护工具（pywebview 桌面壳，macOS / Windows 通用）。"""
import argparse
import os
import re
import shutil
import subprocess
import sys
import tarfile
import time
import urllib.error
import urllib.request
import zipfile

__version__ = "1.1.6"
REPO = "danielzhfzh-hue/BKTC_Ledger"

APP_DIR = sys._MEIPASS if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, APP_DIR)

import core  # noqa: E402
import webview  # noqa: E402


def _parse_version(tag):
    """'v1.2.3' -> (1, 2, 3)；解析不出的部分按 0。"""
    parts = []
    for p in str(tag).strip().lstrip("vV").split("."):
        m = re.match(r"\d+", p)
        parts.append(int(m.group()) if m else 0)
    return tuple(parts)


def _tag_from_location(url):
    """从 .../releases/tag/v1.2.3 提取版本号 1.2.3。"""
    m = re.search(r"/releases/tag/(?:v|V)?([0-9][0-9.]*)", url or "")
    return m.group(1) if m else ""

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
                "rules": core.get_rules(self.store_path),
                "version": __version__}

    def save_data(self, data, rules=None):
        core.save_store(self.store_path, data, rules)
        return {"ok": True, "path": self.store_path}

    def generate(self, data, rules=None):
        if not self.xlsx_path or not os.path.isdir(os.path.dirname(self.xlsx_path)):
            raise RuntimeError("请先选择台账 Excel 文件")
        core.backup_xlsx(self.xlsx_path)
        core.save_store(self.store_path, data, rules)
        out, issues, counts = core.generate_xlsx(data, self.xlsx_path, rules)
        return {"ok": True, "out": out, "issues": issues, "counts": counts}

    def export_xlsx(self, table, rows, fields):
        """把筛选后的行 + 选定字段导出到 ~/Downloads/<表>_导出_<时间>.xlsx。"""
        if not fields:
            raise RuntimeError("未选择导出字段")
        if not isinstance(rows, list):
            raise RuntimeError("无可导出的行")
        dl = os.path.join(os.path.expanduser("~"), "Downloads")
        os.makedirs(dl, exist_ok=True)
        safe = (table or "导出").replace("/", "_").replace("\\", "_")
        path = os.path.join(dl, f"{safe}_导出_{time.strftime('%Y%m%d_%H%M%S')}.xlsx")
        core.export_filtered(path, table, rows, fields)
        return {"ok": True, "path": path, "rows": len(rows), "fields": len(fields)}

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

    def check_update(self):
        """查 GitHub 最新 Release：走网页 releases/latest 302 重定向读 tag，
        避开 API 匿名 60 次/时限流；资产下载 URL 直接构造（公开库免鉴权）。"""
        latest_url = f"https://github.com/{REPO}/releases/latest"
        want = "macOS" if sys.platform == "darwin" else "Windows"
        asset_name = f"BKTC_Ledger-{want}.tar.gz"
        try:
            req = urllib.request.Request(latest_url, headers={"User-Agent": "BKTC_Ledger"})
            with urllib.request.urlopen(req, timeout=15) as resp:
                tag = _tag_from_location(resp.geturl())
        except (urllib.error.URLError, OSError) as e:
            return {"current": __version__, "latest": None, "has_update": False,
                    "error": f"无法访问 GitHub：{e}"}
        if not tag:
            return {"current": __version__, "latest": None, "has_update": False,
                    "error": "未找到最新版本（可能尚未发布 Release）"}
        asset_url = f"https://github.com/{REPO}/releases/download/v{tag}/{asset_name}"
        return {
            "current": __version__, "latest": tag,
            "has_update": _parse_version(tag) > _parse_version(__version__),
            "release_url": f"https://github.com/{REPO}/releases/tag/v{tag}",
            "asset_name": asset_name, "asset_url": asset_url,
        }

    def download_update(self, asset_url, asset_name):
        """下载更新资产到 ~/Downloads/BKTC_Ledger_update/ 并解压；返回解压目录。"""
        if not asset_url or not asset_name:
            raise RuntimeError("没有可下载的更新资产")
        dl_root = os.path.join(os.path.expanduser("~"), "Downloads")
        os.makedirs(dl_root, exist_ok=True)
        dest_dir = os.path.join(dl_root, "BKTC_Ledger_update")
        if os.path.isdir(dest_dir):
            shutil.rmtree(dest_dir, ignore_errors=True)
        os.makedirs(dest_dir, exist_ok=True)
        archive = os.path.join(dest_dir, asset_name)
        req = urllib.request.Request(asset_url, headers={"User-Agent": "BKTC_Ledger"})
        with urllib.request.urlopen(req, timeout=180) as resp, open(archive, "wb") as f:
            shutil.copyfileobj(resp, f)
        if asset_name.endswith((".tar.gz", ".tgz")):
            with tarfile.open(archive, "r:gz") as tf:
                try:
                    tf.extractall(dest_dir, filter="data")
                except TypeError:
                    tf.extractall(dest_dir)
        elif asset_name.endswith(".zip"):
            with zipfile.ZipFile(archive) as zf:
                zf.extractall(dest_dir)
        else:
            raise RuntimeError(f"未知压缩格式：{asset_name}")
        os.remove(archive)
        return dest_dir

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
