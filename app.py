#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""上海康肯销售订单管理系统（pywebview 桌面壳，macOS / Windows 通用）。"""
import argparse
import errno
import getpass
import json
import os
import platform
import re
import shutil
import subprocess
import sys
import tarfile
import tempfile
import time
import urllib.error
import urllib.request
import zipfile

__version__ = "1.7.2"
APP_DISPLAY_NAME = "上海康肯销售订单管理系统"
REPO = "danielzhfzh-hue/BKTC_Ledger"
CANONICAL_PROJECT_ROOT = "/Users/danielzhu/projects/订单整理/BKTC_Ledger"

IS_FROZEN = bool(getattr(sys, "frozen", False))
APP_DIR = sys._MEIPASS if IS_FROZEN else os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, APP_DIR)

import core  # noqa: E402
import database  # noqa: E402
import quotation as quotation_export  # noqa: E402
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
    config = {}
    if isinstance(raw.get("database_path"), str) and raw["database_path"].strip():
        config["database_path"] = os.path.abspath(
            os.path.expanduser(raw["database_path"])
        )
    if isinstance(raw.get("operator_name"), str) and raw["operator_name"].strip():
        config["operator_name"] = raw["operator_name"].strip()
    return config


def _save_config(database_path, path=None, operator_name=None):
    path = os.path.abspath(os.path.expanduser(path or _default_config_path()))
    os.makedirs(os.path.dirname(path), exist_ok=True)
    temporary = path + ".tmp"
    with open(temporary, "w", encoding="utf-8") as f:
        json.dump(
            {
                "database_path": os.path.abspath(database_path),
                "operator_name": str(operator_name or getpass.getuser()).strip(),
            },
            f,
            ensure_ascii=False,
            indent=2,
        )
    os.replace(temporary, path)


def _xlsx_path_for_database(database_path):
    """The ledger workbook is always a generated sibling of the database."""
    stem, _ = os.path.splitext(os.path.abspath(os.path.expanduser(database_path)))
    return stem + ".xlsx"

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


def _windows_user_data_dir(root=None):
    """Return the per-user writable data directory used by packaged Windows builds."""
    if root is None:
        root = (
            os.environ.get("LOCALAPPDATA")
            or os.environ.get("APPDATA")
            or os.path.join(os.path.expanduser("~"), "AppData", "Local")
        )
    return os.path.join(os.path.abspath(os.path.expanduser(root)), "BKTC_Ledger", "data")


def _path_is_writable(path):
    """Probe both the containing directory and an existing file for write access."""
    path = os.path.abspath(os.path.expanduser(path))
    directory = os.path.dirname(path) or os.getcwd()
    try:
        os.makedirs(directory, exist_ok=True)
        if os.path.exists(path):
            # Opening without writing catches read-only attributes and ACL denial.
            with open(path, "ab"):
                pass
        fd, probe = tempfile.mkstemp(prefix=".bktc-write-", dir=directory)
        os.close(fd)
        os.unlink(probe)
        return True
    except OSError:
        try:
            if "probe" in locals() and os.path.exists(probe):
                os.unlink(probe)
        except OSError:
            pass
        return False


def _is_packaged_data_path(path, executable=None, platform_name=None, frozen=None):
    """Whether a path lives in the bundled executable's data directory on Windows."""
    platform_name = sys.platform if platform_name is None else platform_name
    frozen = IS_FROZEN if frozen is None else bool(frozen)
    if not frozen or not str(platform_name).lower().startswith("win"):
        return False
    executable_dir = os.path.dirname(os.path.abspath(executable or sys.executable))
    package_data = os.path.abspath(os.path.join(executable_dir, "data"))
    normalized = os.path.abspath(os.path.expanduser(path))
    try:
        return os.path.commonpath((package_data, normalized)) == package_data
    except ValueError:
        return False


def _resolve_windows_database_path(database_path, *, executable=None,
                                   platform_name=None, frozen=None, user_data_dir=None):
    """Move an unwritable packaged database to a per-user directory.

    Explicit paths outside the bundled ``data`` directory are never changed. This keeps
    portable builds usable from protected folders while preserving user-selected paths.
    """
    platform_name = sys.platform if platform_name is None else platform_name
    frozen = IS_FROZEN if frozen is None else bool(frozen)
    if not frozen or not str(platform_name).lower().startswith("win"):
        return database_path
    db_packaged = _is_packaged_data_path(
        database_path, executable=executable, platform_name=platform_name, frozen=frozen
    )
    db_needs_move = db_packaged and not _path_is_writable(database_path)
    if not db_needs_move:
        return database_path

    target_dir = os.path.abspath(os.path.expanduser(
        user_data_dir or _windows_user_data_dir()
    ))
    os.makedirs(target_dir, exist_ok=True)
    new_db = os.path.join(target_dir, "BKTC_Ledger.db")
    if not os.path.exists(new_db) and os.path.exists(database_path):
        shutil.copy2(database_path, new_db)
    return new_db


def _is_write_permission_error(exc):
    return isinstance(exc, PermissionError) or getattr(exc, "errno", None) in (
        errno.EACCES, errno.EPERM,
    )


def _fallback_xlsx_path(path):
    stem, ext = os.path.splitext(os.path.abspath(os.path.expanduser(path)))
    ext = ext or ".xlsx"
    return f"{stem}.generated_{time.strftime('%Y%m%d_%H%M%S')}_{os.getpid()}{ext}"


def _generate_xlsx_with_fallback(data, path, rules=None):
    """Generate the requested workbook, falling back if Windows has it locked/read-only."""
    notice = ""
    try:
        core.backup_xlsx(path)
    except OSError as exc:
        if not _is_write_permission_error(exc):
            raise
        notice = "旧台账无法备份（文件可能正被 Excel 占用）"
    try:
        out, issues, counts = core.generate_xlsx(data, path, rules)
    except OSError as exc:
        if not _is_write_permission_error(exc):
            raise
        fallback = _fallback_xlsx_path(path)
        out, issues, counts = core.generate_xlsx(data, fallback, rules)
        notice = (
            f"原台账无法覆盖（Windows 权限/占用），已生成备用文件：{fallback}；"
            "关闭 Excel 后可再生成到原路径"
        )
    return out, issues, counts, notice


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
    DEFAULT_DATABASE = _portable_default("BKTC_Ledger.db")
    DEFAULT_LEGACY_JSON = _portable_default("BKTC_Ledger.records.json")
else:
    DEFAULT_DATABASE = _source_default("BKTC_Ledger.db", None)
    DEFAULT_LEGACY_JSON = _source_default("BKTC_Ledger.records.json", None)


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
    def __init__(self, data_path, legacy_json_path=None, config_path=None,
                 operator_name=None):
        self.xlsx_path = ""
        self.database_path = ""
        self.database_revision = None
        self.legacy_json_path = None
        self.config_path = config_path
        self.operator_name = str(operator_name or getpass.getuser()).strip() or "unknown"
        self._set_data_path(data_path)
        if legacy_json_path and not str(data_path).lower().endswith(".json"):
            self.legacy_json_path = legacy_json_path

    def _relocate_windows_database(self):
        """Retry a packaged Windows operation from the writable per-user data directory."""
        old_db = self.database_path
        new_db = _resolve_windows_database_path(old_db)
        if new_db == old_db:
            return False
        self.database_path = new_db
        self.xlsx_path = _xlsx_path_for_database(new_db)
        self.database_revision = database.get_revision(new_db) if os.path.exists(new_db) else None
        if self.config_path:
            _save_config(self.database_path, self.config_path, self.operator_name)
        return True

    def _save_database(self, data, rules=None, audit_context=None, source="manual_save"):
        try:
            return database.save_database(
                self.database_path, data, rules, expected_revision=self.database_revision,
                audit_context=self._audit_context(audit_context, source),
            )
        except OSError as exc:
            if not _is_write_permission_error(exc) or not self._relocate_windows_database():
                raise RuntimeError(
                    f"无法写入数据库：{self.database_path}。请将程序移到可写目录，"
                    "或在设置中选择可写的 .db 文件。"
                ) from exc
            return database.save_database(
                self.database_path, data, rules, expected_revision=self.database_revision,
                audit_context=self._audit_context(audit_context, source),
            )

    def _set_data_path(self, path):
        path = os.path.abspath(os.path.expanduser(path))
        if path.lower().endswith(".json"):
            self.legacy_json_path = path
        else:
            self.legacy_json_path = None
        self.database_path = database.database_path_for(path)
        self.xlsx_path = _xlsx_path_for_database(self.database_path)
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

    def load_state(self, store=None):
        if store:
            self._set_data_path(store)
        migration = None
        if not os.path.exists(self.database_path):
            migration = self._init_database()
        data = database.load_database(self.database_path)
        info = database.get_database_info(self.database_path)
        self.database_revision = info["revision"]
        config_warning = None
        if self.config_path:
            try:
                _save_config(self.database_path, self.config_path, self.operator_name)
            except OSError as exc:
                config_warning = f"路径已切换，但无法保存下次启动设置：{exc}"
        return {"store_path": self.database_path, "database_path": self.database_path,
                "xlsx_path": self.xlsx_path,
                "xlsx_exists": os.path.isfile(self.xlsx_path),
                "data": data, "schema": schema_for_js(),
                "rules": database.get_rules(self.database_path),
                "database_info": info,
                "migration": bool(migration and migration.get("migrated")),
                "migration_source": migration.get("source") if migration else None,
                "config_warning": config_warning,
                "operator_name": self.operator_name,
                "version": __version__}

    def save_data(self, data, rules=None, audit_context=None):
        result = self._save_database(data, rules, audit_context)
        self.database_revision = result["revision"]
        return {"ok": True, "path": self.database_path,
                "data": result["data"], "revision": result["revision"],
                "backup": result["backup"], "issues": result["issues"],
                "audit_event_id": result["audit_event_id"],
                "audit_change_count": result["audit_change_count"]}

    def generate(self, data, rules=None, audit_context=None):
        saved = self._save_database(data, rules, audit_context, "generate_ledger")
        self.database_revision = saved["revision"]
        out, issues, counts, notice = _generate_xlsx_with_fallback(
            saved["data"], self.xlsx_path, rules
        )
        return {"ok": True, "out": out, "issues": issues, "counts": counts,
                "data": saved["data"], "revision": saved["revision"],
                "xlsx_path": out, "xlsx_exists": True, "notice": notice,
                "audit_event_id": saved["audit_event_id"],
                "audit_change_count": saved["audit_change_count"]}

    def get_unpaid_rows(self, data, rules=None):
        """Return the unpaid report using the same calculation as generated Excel."""
        return core.unpaid_report_rows(data, rules)

    def list_quotations(self, search=""):
        return database.list_quotations(self.database_path, search)

    def get_quotation(self, quote_id):
        return database.get_quotation(self.database_path, quote_id)

    def suggest_quotation_number(self, quote_date=None):
        return database.suggest_quotation_number(self.database_path, quote_date)

    def quotation_defaults(self, job_no):
        return database.quotation_defaults(self.database_path, job_no)

    def quotation_history(self, customer, model, exclude_quote_id=None):
        return database.quotation_history(
            self.database_path, customer, model, exclude_quote_id
        )

    def quotation_options(self, customer="", language="zh"):
        return database.quotation_options(self.database_path, customer, language)

    def save_quotation(self, quotation):
        if not isinstance(quotation, dict):
            raise RuntimeError("报价单数据格式无效")
        action = "update_quotation" if quotation.get("quote_id") else "create_quotation"
        context = self._audit_context({
            "source": "quotation_save", "actions": [action],
        }, "quotation_save")
        try:
            result = database.save_quotation(
                self.database_path, quotation,
                expected_revision=self.database_revision,
                audit_context=context,
            )
        except OSError as exc:
            if not _is_write_permission_error(exc) or not self._relocate_windows_database():
                raise RuntimeError(
                    f"无法写入数据库：{self.database_path}。请在设置中选择可写的 .db 文件。"
                ) from exc
            result = database.save_quotation(
                self.database_path, quotation,
                expected_revision=self.database_revision,
                audit_context=context,
            )
        self.database_revision = result["revision"]
        return result

    def delete_quotation(self, quote_id):
        context = self._audit_context({
            "source": "quotation_delete", "actions": ["delete_quotation"],
        }, "quotation_delete")
        try:
            result = database.delete_quotation(
                self.database_path, quote_id,
                expected_revision=self.database_revision,
                audit_context=context,
            )
        except OSError as exc:
            if not _is_write_permission_error(exc) or not self._relocate_windows_database():
                raise RuntimeError(
                    f"无法写入数据库：{self.database_path}。请在设置中选择可写的 .db 文件。"
                ) from exc
            result = database.delete_quotation(
                self.database_path, quote_id,
                expected_revision=self.database_revision,
                audit_context=context,
            )
        self.database_revision = result["revision"]
        return result

    def export_quotation(self, quote_id):
        quotation = database.get_quotation(self.database_path, quote_id)
        download_dir = os.path.join(os.path.expanduser("~"), "Downloads", "报价单")
        os.makedirs(download_dir, exist_ok=True)
        raw_name = f"报价单_{quotation['quote_no']}_{quotation['customer']}"
        safe_name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", raw_name).strip(" .")
        safe_name = safe_name[:120] or "报价单"
        path = os.path.join(download_dir, safe_name + ".xlsx")
        notice = ""
        try:
            quotation_export.export_quotation_xlsx(quotation, path)
        except OSError as exc:
            if not _is_write_permission_error(exc):
                raise
            path = os.path.join(
                download_dir,
                f"{safe_name}_{time.strftime('%Y%m%d_%H%M%S')}.xlsx",
            )
            quotation_export.export_quotation_xlsx(quotation, path)
            notice = "标准报价文件可能正被占用，已改用带时间戳文件"
        return {
            "ok": True, "path": path, "quote_id": quotation["quote_id"],
            "quote_no": quotation["quote_no"], "notice": notice,
        }

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

    def set_operator_name(self, name):
        name = str(name or "").strip()
        if not name:
            raise RuntimeError("操作人不能为空")
        self.operator_name = name
        if self.config_path:
            _save_config(self.database_path, self.config_path, self.operator_name)
        return {"ok": True, "operator_name": self.operator_name}

    def get_audit_events(self, filters=None, limit=500):
        return database.get_audit_events(self.database_path, filters, limit)

    def export_audit(self, filters=None):
        result = database.get_audit_events(self.database_path, filters, 5000)
        dl = os.path.join(os.path.expanduser("~"), "Downloads")
        os.makedirs(dl, exist_ok=True)
        path = os.path.join(
            dl, f"上海康肯销售订单管理系统_审计记录_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
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
        """下载并解压更新，同时把当前 DB/XLSX 带入更新目录，避免替换程序时丢数据。"""
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
        # GitHub Release 只放 starter 数据；本机业务数据永远以当前打开的
        # database_path/xlsx_path 为准，随更新包复制一份供用户直接替换使用。
        package_data_dir = (
            os.path.join(dest_dir, "BKTC_Ledger", "data")
            if asset_name.endswith(".zip")
            else os.path.join(dest_dir, "data")
        )
        os.makedirs(package_data_dir, exist_ok=True)
        for source, filename in (
            (self.database_path, "BKTC_Ledger.db"),
            (self.xlsx_path, "BKTC_Ledger.xlsx"),
        ):
            if os.path.isfile(source):
                shutil.copy2(source, os.path.join(package_data_dir, filename))
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
    ap = argparse.ArgumentParser(description=APP_DISPLAY_NAME)
    ap.add_argument("--database",
                    help="SQLite 数据库路径")
    ap.add_argument("--store",
                    help="兼容旧版：records.json 路径（首次启动迁移为同目录 .db）")
    ap.add_argument("--debug", action="store_true")
    args = ap.parse_args()

    data_path = (
        args.database or args.store or os.environ.get("BKTC_DATABASE")
        or os.environ.get("BKTC_STORE") or config.get("database_path")
        or DEFAULT_DATABASE
    )
    data_path = _prefer_canonical_mac_data(data_path, "BKTC_Ledger.db")
    data_path = _resolve_windows_database_path(data_path)
    legacy_json = DEFAULT_LEGACY_JSON if data_path == DEFAULT_DATABASE else None
    api = Api(
        data_path, legacy_json_path=legacy_json, config_path=config_path,
        operator_name=config.get("operator_name"),
    )
    webview.create_window(
        APP_DISPLAY_NAME,
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
