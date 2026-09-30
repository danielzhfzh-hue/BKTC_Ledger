"""Verified, rollback-capable installation updates for the desktop app."""
import hashlib
import hmac
import json
import os
import shutil
import sqlite3
import stat
import subprocess
import tarfile
import tempfile
import time
import zipfile
from pathlib import Path, PurePosixPath


MAX_ARCHIVE_BYTES = 2 * 1024 * 1024 * 1024
MAX_ARCHIVE_FILES = 100_000


def sha256_file(path):
    digest = hashlib.sha256()
    with open(path, "rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def parse_sha256(text):
    fields = str(text).strip().split()
    if not fields or len(fields[0]) != 64:
        raise ValueError("SHA-256 校验文件格式无效")
    try:
        int(fields[0], 16)
    except ValueError as exc:
        raise ValueError("SHA-256 校验值无效") from exc
    return fields[0].lower()


def verify_sha256(path, checksum_text):
    expected = parse_sha256(checksum_text)
    actual = sha256_file(path)
    if not hmac.compare_digest(actual, expected):
        raise ValueError("更新包 SHA-256 校验失败，已取消安装")
    return actual


def _safe_member_path(root, name):
    name = str(name).replace("\\", "/")
    path = PurePosixPath(name)
    if (not name or "\0" in name or name.startswith("/") or path.is_absolute()
            or any(part in ("", ".", "..") for part in path.parts)
            or (path.parts and ":" in path.parts[0])):
        raise ValueError(f"压缩包包含不安全路径：{name}")
    destination = Path(root).joinpath(*path.parts)
    if os.path.commonpath((os.path.abspath(root), os.path.abspath(destination))) != os.path.abspath(root):
        raise ValueError(f"压缩包路径超出目标目录：{name}")
    return destination


def _ensure_inside_root(root, path):
    resolved_root = os.path.realpath(root)
    resolved_path = os.path.realpath(path)
    if os.path.commonpath((resolved_root, resolved_path)) != resolved_root:
        raise ValueError("压缩包链接超出目标目录")
    return resolved_path


def extract_archive(archive, destination):
    """Extract only regular files and directories beneath destination."""
    root = Path(destination).resolve()
    root.mkdir(parents=True, exist_ok=True)
    total_size = 0
    if zipfile.is_zipfile(archive):
        with zipfile.ZipFile(archive) as package:
            entries = package.infolist()
            if len(entries) > MAX_ARCHIVE_FILES:
                raise ValueError("更新包文件数量异常")
            for entry in entries:
                target = _safe_member_path(root, entry.filename)
                _ensure_inside_root(root, target.parent)
                mode = entry.external_attr >> 16
                if stat.S_ISLNK(mode):
                    raise ValueError("更新包不允许包含符号链接")
                total_size += entry.file_size
                if total_size > MAX_ARCHIVE_BYTES:
                    raise ValueError("更新包解压体积超过安全上限")
                if entry.is_dir():
                    target.mkdir(parents=True, exist_ok=True)
                    continue
                file_type = stat.S_IFMT(mode)
                if file_type and file_type != stat.S_IFREG:
                    raise ValueError("更新包包含非普通文件")
                target.parent.mkdir(parents=True, exist_ok=True)
                with package.open(entry) as source, open(target, "wb") as output:
                    shutil.copyfileobj(source, output)
                if mode:
                    os.chmod(target, mode & 0o777)
        return root

    if not tarfile.is_tarfile(archive):
        raise ValueError("无法识别更新包格式")
    with tarfile.open(archive, "r:gz") as package:
        entries = package.getmembers()
        if len(entries) > MAX_ARCHIVE_FILES:
            raise ValueError("更新包文件数量异常")
        for entry in entries:
            target = _safe_member_path(root, entry.name)
            _ensure_inside_root(root, target.parent)
            if entry.isdir():
                if target.is_symlink():
                    raise ValueError("更新包路径与符号链接冲突")
                target.mkdir(parents=True, exist_ok=True)
                continue
            if entry.issym():
                linkname = entry.linkname.replace("\\", "/")
                link_parts = PurePosixPath(linkname).parts
                if not linkname or os.path.isabs(linkname) or (link_parts and ":" in link_parts[0]):
                    raise ValueError("更新包包含不安全符号链接")
                link_target = (Path(_ensure_inside_root(root, target.parent)) / linkname).resolve()
                if os.path.commonpath((str(root), str(link_target))) != str(root):
                    raise ValueError("更新包符号链接超出目标目录")
                if os.path.lexists(target):
                    raise ValueError("更新包路径重复")
                os.symlink(entry.linkname, target)
                continue
            if not entry.isfile():
                raise ValueError("更新包不允许包含链接或特殊文件")
            total_size += entry.size
            if total_size > MAX_ARCHIVE_BYTES:
                raise ValueError("更新包解压体积超过安全上限")
            target.parent.mkdir(parents=True, exist_ok=True)
            source = package.extractfile(entry)
            if source is None:
                raise ValueError(f"无法读取更新包文件：{entry.name}")
            if target.is_symlink():
                raise ValueError("更新包路径与符号链接冲突")
            _ensure_inside_root(root, target)
            with source, open(target, "wb") as output:
                shutil.copyfileobj(source, output)
            os.chmod(target, entry.mode & 0o777)
    return root


def snapshot_sqlite(source_path, destination_path):
    """Create a transactionally consistent SQLite backup, including WAL state."""
    source_path = os.path.abspath(source_path)
    destination_path = os.path.abspath(destination_path)
    os.makedirs(os.path.dirname(destination_path), exist_ok=True)
    source = sqlite3.connect(source_path, timeout=30)
    destination = sqlite3.connect(destination_path)
    try:
        source.backup(destination)
        check = destination.execute("PRAGMA integrity_check").fetchone()[0]
        if check != "ok":
            raise sqlite3.DatabaseError(f"数据库快照完整性校验失败：{check}")
    finally:
        destination.close()
        source.close()
    return destination_path


def restore_sqlite(snapshot_path, destination_path):
    """Restore a database through SQLite's backup API, never by copying a live DB."""
    os.makedirs(os.path.dirname(os.path.abspath(destination_path)), exist_ok=True)
    source = sqlite3.connect(os.path.abspath(snapshot_path), timeout=30)
    destination = sqlite3.connect(os.path.abspath(destination_path), timeout=30)
    try:
        source.backup(destination)
        check = destination.execute("PRAGMA integrity_check").fetchone()[0]
        if check != "ok":
            raise sqlite3.DatabaseError(f"数据库恢复后完整性校验失败：{check}")
    finally:
        destination.close()
        source.close()


def mark_healthy(marker_path, result_path, version, database_path):
    marker = {"version": str(version), "database_path": os.path.abspath(database_path)}
    _write_json(result_path, {"ok": True, **marker})
    _write_json(marker_path, marker)


def _write_json(path, value):
    if not path:
        return
    path = os.path.abspath(path)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    temporary = path + ".tmp"
    with open(temporary, "w", encoding="utf-8") as output:
        json.dump(value, output, ensure_ascii=False)
    os.replace(temporary, path)


def take_result(path):
    if not path or not os.path.isfile(path):
        return None
    try:
        with open(path, encoding="utf-8") as source:
            result = json.load(source)
    finally:
        try:
            os.remove(path)
        except OSError:
            pass
    return result if isinstance(result, dict) else None


def _pid_running(pid):
    if not pid:
        return False
    try:
        os.kill(int(pid), 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    return True


def _wait_for_exit(pid, timeout):
    deadline = time.monotonic() + timeout
    while _pid_running(pid):
        if time.monotonic() >= deadline:
            raise TimeoutError("等待旧程序退出超时，未替换现有版本")
        time.sleep(0.25)


def _stop_process(process):
    if process is None or process.poll() is not None:
        return
    process.terminate()
    try:
        process.wait(timeout=10)
    except subprocess.TimeoutExpired:
        process.kill()
        process.wait(timeout=10)


def _remove_path(path):
    if os.path.isdir(path) and not os.path.islink(path):
        shutil.rmtree(path)
    elif os.path.lexists(path):
        os.remove(path)


def apply_update(state, *, wait_for_exit=_wait_for_exit, launch=subprocess.Popen,
                 wait_timeout=120, health_timeout=90):
    """Swap in the staged app; restore the old app and data if startup is unhealthy."""
    target = os.path.abspath(state["target_dir"])
    stage = os.path.abspath(state["stage_dir"])
    previous = os.path.abspath(state["previous_dir"])
    result_path = state["result_path"]
    wait_for_exit(state.get("parent_pid"), wait_timeout)
    if not os.path.isdir(target) or not os.path.isdir(stage):
        raise FileNotFoundError("升级目录不完整，原程序未更改")

    process = None
    moved_old = False
    new_installed = False
    try:
        os.replace(target, previous)
        moved_old = True
        os.replace(stage, target)
        new_installed = True
        executable = os.path.join(target, state["executable_relative"])
        marker_path = state["health_marker"]
        args = [executable, "--update-health-marker", marker_path,
                "--update-result-file", result_path]
        process = launch(args, cwd=target)
        deadline = time.monotonic() + health_timeout
        while not os.path.isfile(marker_path) and time.monotonic() < deadline:
            if process.poll() is not None:
                break
            time.sleep(0.25)
        if os.path.isfile(marker_path):
            try:
                _remove_path(previous)
            except OSError:
                pass
            if state.get("temporary_dir"):
                shutil.rmtree(state["temporary_dir"], ignore_errors=True)
            if state.get("stage_parent"):
                shutil.rmtree(state["stage_parent"], ignore_errors=True)
            return {"ok": True, "version": state.get("version")}
        raise RuntimeError("新版本未能在规定时间内完成数据库加载，已回滚")
    except Exception as exc:
        _stop_process(process)
        try:
            if new_installed and os.path.lexists(target):
                _remove_path(target)
            if moved_old and os.path.lexists(previous):
                os.replace(previous, target)
            if (moved_old and state.get("database_backup")
                    and os.path.isfile(state["database_backup"])):
                restore_sqlite(state["database_backup"], state["database_path"])
            xlsx_backup = state.get("xlsx_backup")
            xlsx_path = state.get("xlsx_path")
            if moved_old and xlsx_path:
                if xlsx_backup and os.path.isfile(xlsx_backup):
                    os.makedirs(os.path.dirname(os.path.abspath(xlsx_path)), exist_ok=True)
                    shutil.copy2(xlsx_backup, xlsx_path)
                elif state.get("xlsx_existed") is False and os.path.isfile(xlsx_path):
                    os.remove(xlsx_path)
            message = f"应用升级失败，已恢复旧版本及数据库：{exc}"
            _write_json(result_path, {"ok": False, "message": message})
            old_executable = os.path.join(target, state["executable_relative"])
            launch([old_executable, "--update-result-file", result_path], cwd=target)
            if state.get("temporary_dir"):
                shutil.rmtree(state["temporary_dir"], ignore_errors=True)
            if state.get("stage_parent"):
                shutil.rmtree(state["stage_parent"], ignore_errors=True)
        except Exception as rollback_error:
            _write_json(result_path, {
                "ok": False,
                "message": f"升级失败且自动回滚未完成，请保留现场并联系管理员：{rollback_error}",
            })
            raise RuntimeError(f"{exc}; rollback failed: {rollback_error}") from rollback_error
        return {"ok": False, "message": message}


def run_helper(state_path):
    with open(state_path, encoding="utf-8") as source:
        state = json.load(source)
    try:
        result = apply_update(state)
        _schedule_helper_cleanup(os.path.dirname(os.path.abspath(state_path)))
        return 0 if result.get("ok") else 1
    except Exception as exc:
        _write_json(state.get("result_path"), {"ok": False, "message": str(exc)})
        if state.get("temporary_dir"):
            shutil.rmtree(state["temporary_dir"], ignore_errors=True)
        if state.get("stage_parent"):
            shutil.rmtree(state["stage_parent"], ignore_errors=True)
        _schedule_helper_cleanup(os.path.dirname(os.path.abspath(state_path)))
        return 1


def _schedule_helper_cleanup(helper_root):
    helper_root = os.path.abspath(helper_root)
    if (os.path.dirname(helper_root) != os.path.abspath(tempfile.gettempdir())
            or not os.path.basename(helper_root).startswith("bktc-ledger-updater-")):
        return
    if os.name == "nt":
        command = f'timeout /t 3 /nobreak >nul & rmdir /s /q "{helper_root}"'
        flags = getattr(subprocess, "DETACHED_PROCESS", 0) | getattr(
            subprocess, "CREATE_NEW_PROCESS_GROUP", 0
        ) | getattr(subprocess, "CREATE_NO_WINDOW", 0)
        subprocess.Popen(["cmd.exe", "/c", command], creationflags=flags,
                         stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                         stderr=subprocess.DEVNULL)
    else:
        subprocess.Popen(
            ["/bin/sh", "-c", 'sleep 3; rm -rf -- "$1"', "bktc-update-cleanup", helper_root],
            start_new_session=True, stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
