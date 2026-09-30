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
import secrets
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.request
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

__version__ = "1.14.4"
APP_DISPLAY_NAME = "上海康肯销售订单管理系统"
REPO = "danielzhfzh-hue/BKTC_Ledger"
CANONICAL_PROJECT_ROOT = "/Users/danielzhu/projects/订单整理/BKTC_Ledger"

IS_FROZEN = bool(getattr(sys, "frozen", False))
APP_DIR = sys._MEIPASS if IS_FROZEN else os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, APP_DIR)

import core  # noqa: E402
import database  # noqa: E402
import erp_store  # noqa: E402
import order_ops  # noqa: E402
import move_in_request  # noqa: E402
import quotation as quotation_export  # noqa: E402
import updater  # noqa: E402
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


def _application_install_dir(executable=None, platform_name=None, frozen=None):
    """Return the writable directory containing the installed application.

    PyInstaller's ``sys._MEIPASS`` is an internal resource directory and may be
    temporary.  Use the real executable location instead, and for macOS place
    output beside the ``.app`` bundle rather than inside its read-only bundle.
    """
    platform_name = sys.platform if platform_name is None else platform_name
    frozen = IS_FROZEN if frozen is None else bool(frozen)
    if not frozen:
        return os.path.dirname(os.path.abspath(__file__))

    executable_path = os.path.abspath(executable or sys.executable)
    if platform_name == "darwin":
        current = executable_path
        while True:
            if current.lower().endswith(".app"):
                return os.path.dirname(current)
            parent = os.path.dirname(current)
            if parent == current:
                break
            current = parent
    return os.path.dirname(executable_path)


def _move_in_request_output_dir():
    """Choose the installation-local move-in request directory.

    A protected install location (for example ``Program Files``) cannot be
    written by a normal user.  In that case retain the same stable directory
    name under the user's application-data directory instead of failing after
    the form has already been filled in.
    """
    preferred = os.path.join(_application_install_dir(), "搬入依頼書")
    try:
        os.makedirs(preferred, exist_ok=True)
        if _path_is_writable(os.path.join(preferred, ".write-test")):
            return preferred, ""
    except OSError:
        pass

    fallback = os.path.join(os.path.dirname(_default_config_path()), "搬入依頼書")
    os.makedirs(fallback, exist_ok=True)
    if not _path_is_writable(os.path.join(fallback, ".write-test")):
        raise RuntimeError(
            f"搬入依頼書保存目录不可写：{preferred}；备用目录也不可写：{fallback}"
        )
    return fallback, f"安装目录不可写，已改用用户数据目录：{fallback}"


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
                 operator_name=None, update_health_marker=None,
                 update_result_file=None):
        self.xlsx_path = ""
        self.database_path = ""
        self.database_revision = None
        self.legacy_json_path = None
        self.config_path = config_path
        self.operator_name = str(operator_name or getpass.getuser()).strip() or "unknown"
        self._editable_import = None
        self.update_health_marker = update_health_marker
        self.update_result_file = update_result_file
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
        previous = (self.database_path, self.xlsx_path, self.database_revision,
                    self.legacy_json_path, self._editable_import)
        try:
            result = self._load_state(store)
        except Exception:
            (self.database_path, self.xlsx_path, self.database_revision,
             self.legacy_json_path, self._editable_import) = previous
            raise
        if store:
            self._editable_import = None
        if self.update_health_marker:
            updater.mark_healthy(
                self.update_health_marker, self.update_result_file,
                __version__, self.database_path,
            )
        return result

    def _load_state(self, store=None):
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

    def preview_editable_import(self, workbook_path):
        self._editable_import = None
        prepared = database.prepare_editable_import(self.database_path, workbook_path)
        token = secrets.token_urlsafe(24)
        self._editable_import = (token, prepared)
        return {"token": token, "path": prepared.workbook_path,
                "base_revision": prepared.base_revision, "revision": prepared.current_revision,
                "stale": prepared.stale, "changes": prepared.changes,
                "action_counts": prepared.action_counts, "table_counts": prepared.table_counts,
                "issues": prepared.issues}

    def get_financial_coverage(self, table, job_no="", customer=""):
        return database.get_financial_coverage(self.database_path, table, job_no, customer)

    def pick_editable_workbook(self, exporting=False):
        if not webview.windows:
            raise RuntimeError("应用窗口尚未就绪")
        result = webview.windows[0].create_file_dialog(
            webview.SAVE_DIALOG if exporting else webview.OPEN_DIALOG,
            file_types=("Excel (*.xlsx)",),
            save_filename="BKTC_Ledger_可编辑.xlsx" if exporting else "",
        )
        return result[0] if result else None

    def confirm_editable_import(self, token, confirmed=False, audit_context=None):
        pending = self._editable_import
        if confirmed is not True:
            raise ValueError("必须审查差异并二次确认后才能导入")
        if not pending or not secrets.compare_digest(str(token), pending[0]):
            raise ValueError("差异预览已失效，请重新预览")
        result = database.apply_editable_import(
            self.database_path, pending[1], self._audit_context(audit_context, "xlsx_import")
        )
        self._editable_import = None
        self.database_revision = result["revision"]
        return {"ok": True, "path": self.database_path, "data": result["data"],
                "revision": result["revision"], "backup": result["backup"],
                "issues": result["issues"], "audit_event_id": result["audit_event_id"]}

    def export_editable_workbook(self, workbook_path):
        if not str(workbook_path).lower().endswith(".xlsx"):
            raise ValueError("可编辑工作簿必须保存为 .xlsx")
        return database.export_editable_workbook(self.database_path, workbook_path)

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

    def get_erp_workspace(self, data=None, rules=None):
        """Build ERP read models from current UI data or the authoritative database.

        Passing the in-memory six-table payload lets the dashboard preview unsaved
        edits while all calculations still use the canonical Python domain layer.
        """
        if not isinstance(data, dict):
            data = database.load_database(self.database_path)
        if not isinstance(rules, dict):
            rules = database.get_rules(self.database_path)
        work_items = erp_store.list_work_items(self.database_path)
        return order_ops.build_workspace(data, rules, work_items=work_items)

    def list_master_data(self):
        return erp_store.list_master_data(self.database_path)

    def save_master_record(self, kind, record):
        context = self._audit_context({
            "source": "master_data_save", "actions": [f"save_master_{kind}"],
        }, "master_data_save")
        result = erp_store.save_master_record(
            self.database_path, kind, record,
            expected_revision=self.database_revision,
            audit_context=context,
        )
        self.database_revision = result["revision"]
        result["master_data"] = erp_store.list_master_data(self.database_path)
        return result

    def list_work_items(self, job_no="", status=""):
        return erp_store.list_work_items(self.database_path, job_no, status)

    def save_work_item(self, item):
        context = self._audit_context({
            "source": "work_item_save", "actions": ["save_work_item"],
        }, "work_item_save")
        result = erp_store.save_work_item(
            self.database_path, item,
            expected_revision=self.database_revision,
            audit_context=context,
        )
        self.database_revision = result["revision"]
        result["work_items"] = erp_store.list_work_items(self.database_path)
        return result

    def complete_work_item(self, item_id):
        context = self._audit_context({
            "source": "work_item_complete", "actions": ["complete_work_item"],
        }, "work_item_complete")
        result = erp_store.complete_work_item(
            self.database_path, item_id,
            expected_revision=self.database_revision,
            audit_context=context,
        )
        self.database_revision = result["revision"]
        result["work_items"] = erp_store.list_work_items(self.database_path)
        return result


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

    def get_move_in_request_defaults(self, shipment_id):
        """Read a shipment and its linked order/devices from the current SQLite DB."""
        if not os.path.exists(self.database_path):
            self.load_state()
        current_revision = database.get_revision(self.database_path)
        if (self.database_revision is not None
                and current_revision != self.database_revision):
            raise database.StaleImportError(
                f"数据库已由其他窗口更新（当前修订 {current_revision}，本窗口为 {self.database_revision}），"
                "请重新加载后再生成搬入依頼書。"
            )
        data = database.load_database(self.database_path)
        defaults = move_in_request.shipment_defaults(data, shipment_id=shipment_id)
        defaults["database_revision"] = current_revision
        return defaults

    def export_move_in_request(self, form):
        """Validate the selected batch against SQLite, then export a move-in request."""
        if not isinstance(form, dict):
            raise RuntimeError("搬入依頼書数据格式无效")
        if not os.path.exists(self.database_path):
            self.load_state()
        current_revision = database.get_revision(self.database_path)
        expected = form.get("database_revision")
        if expected not in (None, "") and int(expected) != current_revision:
            raise database.StaleImportError(
                f"数据库已变化（当前修订 {current_revision}，窗口读取为 {expected}），请重新打开搬入依頼書窗口。"
            )
        data = database.load_database(self.database_path)
        defaults = move_in_request.shipment_defaults(
            data, shipment_id=form.get("shipment_id"),
            job_no=form.get("job_no"), batch=form.get("batch"),
        )
        # Keep the authoritative device/order fields from SQLite.  Only the
        # explicitly supplemental form fields are accepted from the dialog.
        merged = dict(defaults)
        for key in ("issue_date", "move_in_date", "move_in_time", "contact",
                    "vehicle_type", "vehicle_tonnage", "vehicle_note", "department",
                    "other_department", "notes", "manager", "issuer", "customer",
                    "address", "has_customer_id"):
            if key in form:
                merged[key] = form[key]
        errors = move_in_request.validate_form(merged)
        if errors:
            raise RuntimeError("；".join(errors))
        download_dir, notice = _move_in_request_output_dir()
        raw_name = f"搬入依頼書_{merged['job_no']}_{merged['batch']}_{merged.get('move_in_date') or '未定'}"
        safe_name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", raw_name).strip(" .")
        safe_name = safe_name[:150] or "搬入依頼書"
        path = os.path.join(download_dir, safe_name + ".xlsx")
        try:
            move_in_request.export_xlsx(merged, path)
        except OSError as exc:
            if not _is_write_permission_error(exc):
                raise
            path = os.path.join(
                download_dir, f"{safe_name}_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
            )
            move_in_request.export_xlsx(merged, path)
            notice = "目标文件可能正被 Excel 占用，已改用带时间戳文件"
        return {
            "ok": True, "path": path, "notice": notice,
            "job_no": merged["job_no"], "batch": merged["batch"],
            "device_count": len(merged["devices"]),
            "database_revision": current_revision,
        }

    def set_operator_name(self, name):
        name = str(name or "").strip()
        if not name:
            raise RuntimeError("操作人不能为空")
        self.operator_name = name
        if self.config_path:
            _save_config(self.database_path, self.config_path, self.operator_name)
        return {"ok": True, "operator_name": self.operator_name}

    def get_audit_events(self, filters=None, limit=500, offset=0):
        return database.get_audit_events(self.database_path, filters, limit, offset)

    def export_audit(self, filters=None):
        result = database.get_audit_events(self.database_path, filters, limit=None)
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
            "checksum_url": asset_url + ".sha256",
        }

    @staticmethod
    def _download_update_file(url, path, limit):
        parsed = urlparse(url or "")
        if parsed.scheme != "https" or parsed.hostname != "github.com":
            raise ValueError("更新文件只允许从 GitHub HTTPS Release 下载")
        if "/releases/download/" not in parsed.path:
            raise ValueError("更新链接不是 GitHub Release 资产")
        request = urllib.request.Request(url, headers={"User-Agent": "BKTC_Ledger"})
        total = 0
        with urllib.request.urlopen(request, timeout=180) as response, open(path, "wb") as output:
            while True:
                chunk = response.read(1024 * 1024)
                if not chunk:
                    break
                total += len(chunk)
                if total > limit:
                    raise ValueError("更新文件超过大小限制")
                output.write(chunk)

    def _install_locations(self):
        if not IS_FROZEN:
            raise RuntimeError("开发运行模式不支持应用内安装更新")
        if os.name == "nt":
            target = os.path.dirname(os.path.abspath(sys.executable))
            return target, os.path.basename(sys.executable), "Windows"
        executable = Path(sys.executable).resolve()
        for parent in (executable, *executable.parents):
            if parent.name.endswith(".app"):
                return str(parent), str(executable.relative_to(parent)), "macOS"
        raise RuntimeError("无法确定当前 macOS 应用包位置")

    @staticmethod
    def _find_release_app(unpacked, platform_name):
        root = Path(unpacked)
        if platform_name == "Windows":
            matches = list(root.rglob("BKTC_Ledger.exe"))
            if not matches:
                raise RuntimeError("更新包中未找到 BKTC_Ledger.exe")
            return str(matches[0].parent)
        matches = list(root.rglob("BKTC_Ledger.app"))
        if not matches:
            raise RuntimeError("更新包中未找到 BKTC_Ledger.app")
        return str(matches[0])

    @staticmethod
    def _copy_windows_helper(target_dir, helper_dir):
        shutil.copytree(
            target_dir, helper_dir,
            ignore=shutil.ignore_patterns("data", "备份", "搬入依頼书"),
        )

    def install_update(self, asset_url, checksum_url, asset_name, version):
        """Verify, stage and hand off to an isolated helper for atomic swap/rollback."""
        target_dir, executable_relative, platform_name = self._install_locations()
        expected_name = ("BKTC_Ledger-Windows.zip" if platform_name == "Windows"
                         else "BKTC_Ledger-macOS.tar.gz")
        if asset_name != expected_name:
            raise ValueError("更新资产与当前操作系统不匹配")
        expected_url = f"https://github.com/{REPO}/releases/download/v{version}/{asset_name}"
        if asset_url != expected_url or checksum_url != expected_url + ".sha256":
            raise ValueError("更新包与校验文件必须来自同一个 GitHub Release")
        if not os.path.isfile(self.database_path):
            raise RuntimeError(f"找不到当前数据库，已取消更新：{self.database_path}")

        install_parent = os.path.dirname(target_dir)
        result_root = os.path.dirname(os.path.abspath(self.config_path or _default_config_path()))
        os.makedirs(result_root, exist_ok=True)
        result_path = os.path.join(result_root, "last_update_result.json")
        update_root = tempfile.mkdtemp(prefix="bktc-ledger-update-")
        stage_root = None
        helper_root = None
        try:
            archive = os.path.join(update_root, asset_name)
            checksum_path = archive + ".sha256"
            self._download_update_file(asset_url, archive, updater.MAX_ARCHIVE_BYTES)
            self._download_update_file(checksum_url, checksum_path, 4096)
            with open(checksum_path, encoding="utf-8") as checksum_file:
                updater.verify_sha256(archive, checksum_file.read())

            unpacked = os.path.join(update_root, "unpacked")
            updater.extract_archive(archive, unpacked)
            release_app = self._find_release_app(unpacked, platform_name)
            stage_root = tempfile.mkdtemp(prefix=".BKTC_Ledger-stage-", dir=install_parent)
            stage_dir = os.path.join(stage_root, "program")
            shutil.copytree(release_app, stage_dir, symlinks=True)

            # Keep an application-support snapshot permanently as the pre-update
            # recovery point; SQLite backup accounts for committed WAL contents.
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
            safe_version = re.sub(r"[^0-9.]", "", str(version)) or "unknown"
            backup_dir = os.path.join(result_root, "更新备份")
            os.makedirs(backup_dir, exist_ok=True)
            db_backup = os.path.join(
                backup_dir,
                f"{Path(self.database_path).stem}_pre_update_v{safe_version}_{timestamp}.db",
            )
            updater.snapshot_sqlite(self.database_path, db_backup)
            xlsx_existed = os.path.isfile(self.xlsx_path)
            xlsx_backup = os.path.join(update_root, "ledger.xlsx") if xlsx_existed else None
            if xlsx_existed:
                shutil.copy2(self.xlsx_path, xlsx_backup)

            # Windows portable databases live inside the replaceable app folder.
            # Put the consistent SQLite snapshot and its workbook into staged data;
            # macOS and relocated Windows databases remain at their configured path.
            try:
                db_relative = os.path.relpath(self.database_path, target_dir)
                db_inside_app = not db_relative.startswith(".." + os.sep) and db_relative != ".."
                if db_inside_app:
                    staged_db = os.path.join(stage_dir, db_relative)
                    os.makedirs(os.path.dirname(staged_db), exist_ok=True)
                    shutil.copy2(db_backup, staged_db)
                    if xlsx_existed:
                        staged_xlsx = _xlsx_path_for_database(staged_db)
                        os.makedirs(os.path.dirname(staged_xlsx), exist_ok=True)
                        shutil.copy2(xlsx_backup, staged_xlsx)
            except ValueError:
                pass

            stamp = f"v{safe_version}_{timestamp}"
            previous_dir = target_dir + ".previous-" + stamp
            health_marker = os.path.join(update_root, "healthy.json")
            helper_root = tempfile.mkdtemp(prefix="bktc-ledger-updater-")
            if platform_name == "Windows":
                helper_app = os.path.join(helper_root, "BKTC_Ledger")
                self._copy_windows_helper(target_dir, helper_app)
                helper_executable = os.path.join(
                    helper_app, os.path.basename(sys.executable)
                )
            else:
                # macOS permits renaming a running app bundle; the helper has
                # already imported its code and runs with cwd outside the bundle.
                helper_executable = sys.executable
            state_path = os.path.join(helper_root, "update_state.json")
            state = {
                "target_dir": target_dir, "stage_dir": stage_dir,
                "previous_dir": previous_dir, "parent_pid": os.getpid(),
                "executable_relative": executable_relative,
                "health_marker": health_marker, "result_path": result_path,
                "database_path": self.database_path, "database_backup": db_backup,
                "xlsx_path": self.xlsx_path, "xlsx_backup": xlsx_backup,
                "xlsx_existed": xlsx_existed, "version": version,
                "temporary_dir": update_root, "stage_parent": stage_root,
            }
            with open(state_path, "w", encoding="utf-8") as state_file:
                json.dump(state, state_file, ensure_ascii=False)
            command = [helper_executable, "--apply-update", state_path]
            if os.name == "nt":
                flags = getattr(subprocess, "DETACHED_PROCESS", 0) | getattr(
                    subprocess, "CREATE_NEW_PROCESS_GROUP", 0
                )
                subprocess.Popen(command, cwd=helper_root, creationflags=flags)
            else:
                subprocess.Popen(command, cwd=helper_root, start_new_session=True)
            threading.Timer(1.0, self._close_for_update).start()
            return {"ok": True, "version": version}
        except Exception:
            if stage_root and os.path.isdir(stage_root):
                shutil.rmtree(stage_root, ignore_errors=True)
            shutil.rmtree(update_root, ignore_errors=True)
            if helper_root and os.path.isdir(helper_root):
                shutil.rmtree(helper_root, ignore_errors=True)
            raise

    @staticmethod
    def _close_for_update():
        try:
            if webview.windows:
                webview.windows[0].destroy()
        except Exception:
            pass

    def get_update_result(self):
        result_path = os.path.join(
            os.path.dirname(os.path.abspath(self.config_path or _default_config_path())),
            "last_update_result.json",
        )
        return updater.take_result(result_path)

    def open_path(self, path):
        if sys.platform == "darwin":
            subprocess.Popen(["open", path])
        elif os.name == "nt":
            os.startfile(path)  # noqa: S606
        else:
            subprocess.Popen(["xdg-open", path])
        return True


def main():
    ap = argparse.ArgumentParser(description=APP_DISPLAY_NAME)
    ap.add_argument("--database",
                    help="SQLite 数据库路径")
    ap.add_argument("--store",
                    help="兼容旧版：records.json 路径（首次启动迁移为同目录 .db）")
    ap.add_argument("--debug", action="store_true")
    ap.add_argument("--apply-update", help=argparse.SUPPRESS)
    ap.add_argument("--update-health-marker", help=argparse.SUPPRESS)
    ap.add_argument("--update-result-file", help=argparse.SUPPRESS)
    args = ap.parse_args()
    if args.apply_update:
        raise SystemExit(updater.run_helper(args.apply_update))

    config_path = _default_config_path()
    config = _load_config(config_path)

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
        update_health_marker=args.update_health_marker,
        update_result_file=args.update_result_file,
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
