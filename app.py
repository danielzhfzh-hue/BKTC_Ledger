#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""BKTC 台账维护工具（pywebview 桌面壳，macOS / Windows 通用）。"""
import argparse
import getpass
import json
import os
import platform
import re
import shutil
import subprocess
import sys
import tarfile
import time
import urllib.error
import urllib.request
import uuid
import zipfile

__version__ = "1.5.1"
REPO = "danielzhfzh-hue/BKTC_Ledger"
CANONICAL_PROJECT_ROOT = "/Users/danielzhu/projects/订单整理/BKTC_Ledger"

IS_FROZEN = bool(getattr(sys, "frozen", False))
APP_DIR = sys._MEIPASS if IS_FROZEN else os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, APP_DIR)

import core  # noqa: E402
import database  # noqa: E402
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


def _default_config_path():
    if os.name == "nt":
        root = os.environ.get("APPDATA") or os.path.join(
            os.path.expanduser("~"), "AppData", "Roaming"
        )
    elif sys.platform == "darwin":
        root = os.path.join(os.path.expanduser("~"), "Library", "Application Support")
    else:
        root = os.environ.get("XDG_CONFIG_HOME") or os.path.join(
            os.path.expanduser("~"), ".config"
        )
    return os.path.join(root, "BKTC_Ledger", "config.json")


def _load_config(path=None):
    try:
        with open(path or _default_config_path(), encoding="utf-8") as f:
            raw = json.load(f)
    except (OSError, json.JSONDecodeError, TypeError):
        return {}
    if not isinstance(raw, dict):
        return {}
    config = {
        key: os.path.abspath(os.path.expanduser(value))
        for key in ("database_path", "xlsx_path")
        if isinstance((value := raw.get(key)), str) and value.strip()
    }
    if isinstance(raw.get("operator_name"), str) and raw["operator_name"].strip():
        config["operator_name"] = raw["operator_name"].strip()
    return config


def _save_config(database_path, xlsx_path, path=None, operator_name=None):
    path = os.path.abspath(os.path.expanduser(path or _default_config_path()))
    os.makedirs(os.path.dirname(path), exist_ok=True)
    temporary = path + ".tmp"
    with open(temporary, "w", encoding="utf-8") as f:
        json.dump(
            {
                "database_path": os.path.abspath(database_path),
                "xlsx_path": os.path.abspath(xlsx_path),
                "operator_name": str(operator_name or getpass.getuser()).strip(),
            },
            f,
            ensure_ascii=False,
            indent=2,
        )
    os.replace(temporary, path)

def _portable_default(filename, executable=None):
    """Find portable data in a sibling data directory or beside the executable."""
    executable_dir = os.path.dirname(os.path.abspath(executable or sys.executable))
    candidates = []
    directory = executable_dir
    for _ in range(4):
        data_candidate = os.path.join(directory, "data", filename)
        if data_candidate not in candidates:
            candidates.append(data_candidate)
        candidate = os.path.join(directory, filename)
        if candidate not in candidates:
            candidates.append(candidate)
        parent = os.path.dirname(directory)
        if parent == directory:
            break
        directory = parent
    for candidate in candidates:
        if os.path.exists(candidate):
            return candidate
    canonical = os.path.join(
        "/Users/danielzhu/projects/订单整理/BKTC_Ledger", "data", filename
    )
    if sys.platform == "darwin" and os.path.exists(canonical):
        return canonical
    for directory in [executable_dir, *[os.path.dirname(x) for x in candidates]]:
        if directory.lower().endswith(".app"):
            return os.path.join(os.path.dirname(directory), filename)
    return candidates[0]


def _is_ephemeral_path(path):
    """旧审计/解压流程可能留下 /tmp 路径，不能作为长期工作库。"""
    raw = str(path).replace("\\", "/")
    if raw == "/tmp" or raw.startswith(("/tmp/", "/private/tmp/")):
        return True
    try:
        normalized = os.path.abspath(os.path.expanduser(path))
    except (TypeError, ValueError):
        return False
    return normalized == "/tmp" or normalized.startswith(("/tmp/", "/private/tmp/"))


def _prefer_canonical_mac_data(path, filename, platform_name=None, project_root=None):
    """Keep this Mac's working data in the project's single canonical data directory."""
    platform_name = sys.platform if platform_name is None else platform_name
    if platform_name != "darwin":
        return path
    project_root = project_root or CANONICAL_PROJECT_ROOT
    canonical = os.path.join(project_root, "data", filename)
    if not os.path.exists(canonical):
        return path
    normalized = os.path.abspath(os.path.expanduser(path))
    release_root = os.path.join(project_root, "releases") + os.sep
    if _is_ephemeral_path(normalized) or normalized.startswith(release_root):
        return canonical
    return path


def _source_default(filename, legacy_path, platform=None):
    """Use the repository's persistent data directory in source mode."""
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", filename)


if IS_FROZEN:
    DEFAULT_XLSX = _portable_default("BKTC_Ledger.xlsx")
    DEFAULT_DATABASE = _portable_default("BKTC_Ledger.db")
    DEFAULT_LEGACY_JSON = _portable_default("BKTC_Ledger.records.json")
else:
    DEFAULT_XLSX = _source_default(
        "BKTC_Ledger.xlsx",
        r"/Users/danielzhu/projects/订单整理/BKTC上海POU营业管理表.xlsx",
    )
    DEFAULT_DATABASE = _source_default(
        "BKTC_Ledger.db",
        r"/Users/danielzhu/projects/订单整理/BKTC上海POU营业管理表.db",
    )
    DEFAULT_LEGACY_JSON = _source_default(
        "BKTC_Ledger.records.json",
        r"/Users/danielzhu/projects/订单整理/BKTC上海POU营业管理表.records.json",
    )


def schema_for_js():
    out = {}
    for t in core.TABLES:
        derived = set(core.DERIVED.get(t, []))
        out[t] = [
            [field, field_type, options[0] if options else [], field in derived]
            for field, field_type, *options in core.SCHEMA[t]
        ]
    return out


class Api:
    def __init__(self, xlsx_path, data_path, legacy_json_path=None, config_path=None,
                 operator_name=None):
        self.xlsx_path = xlsx_path
        self.database_path = ""
        self.database_revision = None
        self.legacy_json_path = None
        self.config_path = config_path
        self.operator_name = str(operator_name or getpass.getuser()).strip() or "unknown"
        self._pending_imports = {}
        self._set_data_path(data_path)
        if legacy_json_path and not str(data_path).lower().endswith(".json"):
            self.legacy_json_path = legacy_json_path

    def _set_data_path(self, path):
        path = os.path.abspath(os.path.expanduser(path))
        if path.lower().endswith(".json"):
            self.legacy_json_path = path
        else:
            self.legacy_json_path = None
        self.database_path = database.database_path_for(path)
        self.database_revision = None

    def _init_database(self):
        if self.legacy_json_path and os.path.exists(self.legacy_json_path):
            return database.migrate_json_to_database(
                self.legacy_json_path, self.database_path
            )
        return database.save_database(
            self.database_path, core.empty_data(), reason="database_created", backup=False
        )

    def _audit_context(self, requested=None, source="manual_save"):
        requested = requested if isinstance(requested, dict) else {}
        actions = requested.get("actions") if isinstance(requested.get("actions"), list) else []
        return {
            "operator_name": self.operator_name,
            "source": str(requested.get("source") or source),
            "actions": [str(x) for x in actions],
            "app_version": __version__,
            "platform": platform.platform(),
            "hostname": platform.node(),
        }

    def load_state(self, store=None, xlsx=None):
        if store:
            self._set_data_path(store)
            self._pending_imports.clear()
        if xlsx:
            self.xlsx_path = os.path.abspath(os.path.expanduser(xlsx))
        migration = None
        if not os.path.exists(self.database_path):
            migration = self._init_database()
        data = database.load_database(self.database_path)
        info = database.get_database_info(self.database_path)
        self.database_revision = info["revision"]
        config_warning = None
        if self.config_path:
            try:
                _save_config(
                    self.database_path, self.xlsx_path, self.config_path,
                    self.operator_name,
                )
            except OSError as exc:
                config_warning = f"路径已切换，但无法保存下次启动设置：{exc}"
        return {"store_path": self.database_path, "database_path": self.database_path,
                "xlsx_path": self.xlsx_path,
                "data": data, "schema": schema_for_js(),
                "rules": database.get_rules(self.database_path),
                "database_info": info,
                "migration": bool(migration and migration.get("migrated")),
                "migration_source": migration.get("source") if migration else None,
                "config_warning": config_warning,
                "operator_name": self.operator_name,
                "version": __version__}

    def save_data(self, data, rules=None, audit_context=None):
        result = database.save_database(
            self.database_path, data, rules, expected_revision=self.database_revision,
            audit_context=self._audit_context(audit_context),
        )
        self.database_revision = result["revision"]
        return {"ok": True, "path": self.database_path,
                "data": result["data"], "revision": result["revision"],
                "backup": result["backup"], "issues": result["issues"],
                "audit_event_id": result["audit_event_id"],
                "audit_change_count": result["audit_change_count"]}

    def generate(self, data, rules=None, audit_context=None):
        if not self.xlsx_path or not os.path.isdir(os.path.dirname(self.xlsx_path)):
            raise RuntimeError("请先选择台账 Excel 文件")
        saved = database.save_database(
            self.database_path, data, rules, expected_revision=self.database_revision,
            audit_context=self._audit_context(audit_context, "generate_ledger"),
        )
        self.database_revision = saved["revision"]
        core.backup_xlsx(self.xlsx_path)
        out, issues, counts = core.generate_xlsx(saved["data"], self.xlsx_path, rules)
        return {"ok": True, "out": out, "issues": issues, "counts": counts,
                "data": saved["data"], "revision": saved["revision"],
                "audit_event_id": saved["audit_event_id"],
                "audit_change_count": saved["audit_change_count"]}

    def get_unpaid_rows(self, data, rules=None):
        """Return the unpaid report using the same calculation as generated Excel."""
        return core.unpaid_report_rows(data, rules)

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

    def export_editable(self, data, rules=None, dirty=True, audit_context=None):
        """Save current edits, then export a revision-bound six-table workbook."""
        if dirty:
            saved = database.save_database(
                self.database_path, data, rules, expected_revision=self.database_revision,
                audit_context=self._audit_context(audit_context, "export_editable"),
            )
            self.database_revision = saved["revision"]
        else:
            saved = {
                "data": database.load_database(self.database_path),
                "revision": database.get_revision(self.database_path),
            }
            self.database_revision = saved["revision"]
        dl = os.path.join(os.path.expanduser("~"), "Downloads")
        os.makedirs(dl, exist_ok=True)
        path = os.path.join(
            dl, f"BKTC_订单编辑_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        result = database.export_editable_workbook(self.database_path, path)
        result.update({"data": saved["data"], "revision": saved["revision"]})
        return result

    def preview_import(self, path):
        prepared = database.prepare_editable_import(self.database_path, path)
        token = uuid.uuid4().hex
        self._pending_imports = {token: prepared}  # 仅保留最后一次预览，避免误用旧令牌
        warnings = [x for x in prepared.issues if x.get("severity") == "warning"]
        return {
            "ok": True,
            "token": token,
            "path": prepared.workbook_path,
            "base_revision": prepared.base_revision,
            "current_revision": prepared.current_revision,
            "stale": prepared.stale,
            "action_counts": prepared.action_counts,
            "table_counts": prepared.table_counts,
            "changes": prepared.changes[:500],
            "changes_total": len(prepared.changes),
            "changes_truncated": len(prepared.changes) > 500,
            "warnings": warnings[:100],
        }

    def apply_import(self, token):
        prepared = self._pending_imports.get(token)
        if not prepared:
            raise RuntimeError("差异预览已失效，请重新选择 Excel 并预览")
        result = database.apply_editable_import(
            self.database_path, prepared,
            audit_context=self._audit_context(
                {"source": "xlsx_import", "actions": ["xlsx_confirmed_import"]},
                "xlsx_import",
            ),
        )
        self._pending_imports.clear()
        self.database_revision = result["revision"]
        return {
            "ok": True,
            "path": self.database_path,
            "data": result["data"],
            "rules": database.get_rules(self.database_path),
            "revision": result["revision"],
            "backup": result["backup"],
            "audit_event_id": result["audit_event_id"],
            "audit_change_count": result["audit_change_count"],
        }

    def set_operator_name(self, name):
        name = str(name or "").strip()
        if not name:
            raise RuntimeError("操作人不能为空")
        self.operator_name = name
        if self.config_path:
            _save_config(
                self.database_path, self.xlsx_path, self.config_path,
                self.operator_name,
            )
        return {"ok": True, "operator_name": self.operator_name}

    def get_audit_events(self, filters=None, limit=500):
        return database.get_audit_events(self.database_path, filters, limit)

    def export_audit(self, filters=None):
        result = database.get_audit_events(self.database_path, filters, 5000)
        dl = os.path.join(os.path.expanduser("~"), "Downloads")
        os.makedirs(dl, exist_ok=True)
        path = os.path.join(dl, f"BKTC_审计记录_{time.strftime('%Y%m%d_%H%M%S')}.xlsx")
        database.export_audit_xlsx(path, result["events"])
        return {"ok": True, "path": path, "events": len(result["events"]),
                "chain": result["chain"]}

    def pick_store(self):
        w = webview.windows[0]
        res = w.create_file_dialog(webview.OPEN_DIALOG,
                                   file_types=("SQLite (*.db;*.sqlite;*.sqlite3)",
                                               "Legacy JSON (*.json)",
                                               "All files (*.*)"))
        return res[0] if res else None

    def pick_import_xlsx(self):
        w = webview.windows[0]
        res = w.create_file_dialog(webview.OPEN_DIALOG,
                                   file_types=("Excel (*.xlsx)", "All files (*.*)"))
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
        asset_name = (f"BKTC_Ledger-{want}.tar.gz" if sys.platform == "darwin"
                      else f"BKTC_Ledger-{want}.zip")
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
    config_path = _default_config_path()
    config = _load_config(config_path)
    ap = argparse.ArgumentParser(description="BKTC 台账维护工具")
    ap.add_argument("--xlsx",
                    help="台账 Excel 路径（可用 --xlsx 或界面选择）")
    ap.add_argument("--database",
                    help="SQLite 数据库路径")
    ap.add_argument("--store",
                    help="兼容旧版：records.json 路径（首次启动迁移为同目录 .db）")
    ap.add_argument("--debug", action="store_true")
    args = ap.parse_args()

    xlsx_path = (
        args.xlsx or os.environ.get("BKTC_XLSX")
        or config.get("xlsx_path") or DEFAULT_XLSX
    )
    data_path = (
        args.database or args.store or os.environ.get("BKTC_DATABASE")
        or os.environ.get("BKTC_STORE") or config.get("database_path")
        or DEFAULT_DATABASE
    )
    data_path = _prefer_canonical_mac_data(data_path, "BKTC_Ledger.db")
    xlsx_path = _prefer_canonical_mac_data(xlsx_path, "BKTC_Ledger.xlsx")
    legacy_json = DEFAULT_LEGACY_JSON if data_path == DEFAULT_DATABASE else None
    api = Api(
        xlsx_path, data_path, legacy_json_path=legacy_json, config_path=config_path,
        operator_name=config.get("operator_name"),
    )
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
