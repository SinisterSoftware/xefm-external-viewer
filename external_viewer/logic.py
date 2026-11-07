"""Route files to configurable external viewers."""

import atexit
import fnmatch
import logging
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

from xefm.user_api import extension

logger = logging.getLogger(__name__)
_tmp_root = None


def _tmp_dir():
    global _tmp_root
    if _tmp_root is None or not _tmp_root.exists():
        _tmp_root = Path(tempfile.mkdtemp(prefix="xefm_viewer_"))
        atexit.register(shutil.rmtree, _tmp_root, ignore_errors=True)
    return _tmp_root


def _materialize(path):
    if "://" not in str(path):
        return path
    target = Path(tempfile.mkdtemp(dir=_tmp_dir())) / path.name
    target.write_bytes(path.read_bytes())
    return target


def _find_viewer(filename, viewers):
    if not isinstance(viewers, (list, tuple)):
        raise ValueError("VIEWERS must be a list of rules")
    for rule in viewers:
        if not isinstance(rule, dict) or "pattern" not in rule or "command" not in rule:
            raise ValueError("each VIEWERS rule needs pattern and command")
        patterns = rule["pattern"]
        if isinstance(patterns, str):
            patterns = [patterns]
        if not isinstance(patterns, (list, tuple)) or not patterns or not all(
            isinstance(pattern, str) and pattern for pattern in patterns
        ):
            raise ValueError("pattern must be a non-empty string or list of strings")
        if any(fnmatch.fnmatchcase(filename.lower(), pattern.lower()) for pattern in patterns):
            return rule
    return None


def _open_office(kind, path):
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
        prog_id, collection, args = {
            "excel": ("Excel.Application", "Workbooks", (0, True)),
            "word": ("Word.Application", "Documents", (False, True)),
            "powerpoint": ("PowerPoint.Application", "Presentations", (True, False, True)),
        }[kind]
        app = win32com.client.Dispatch(prog_id)
        app.Visible = True
        getattr(app, collection).Open(str(path), *args)
    finally:
        pythoncom.CoUninitialize()


def _launch(ctx, command, target):
    if command is None:
        os.startfile(str(target))
    elif callable(command):
        return command(ctx, target) is not False
    elif isinstance(command, str) and command in {
        "office:excel", "office:word", "office:powerpoint"
    }:
        _open_office(command.split(":", 1)[1], target)
    elif isinstance(command, (list, tuple)) and command and all(
        isinstance(arg, str) and arg for arg in command
    ):
        subprocess.Popen([*command, str(target)])
    else:
        raise ValueError("command must be a non-empty argv list, callable, None, or office:excel/word/powerpoint")
    return True


def _viewers():
    handle = extension(__name__)
    if handle is None or handle.settings is None:
        raise RuntimeError("extension settings are unavailable")
    return handle.settings.VIEWERS


def open_with_external_viewer(ctx, event):
    """Claim Enter only when an external viewer was started successfully."""
    return open_path(ctx, event.path)


def open_path(ctx, path):
    try:
        rule = _find_viewer(path.name, _viewers())
        if rule is None:
            return False
        return _launch(ctx, rule["command"], _materialize(path))
    except Exception as exc:
        logger.exception("External viewer failed for %s", path)
        ctx.message(f"External viewer failed for {path.name}: {exc}; falling back to XeFM")
        return False


def view_file_external(ctx):
    """View the focused file externally, or use XeFM's normal viewer."""
    entry = ctx.pane.focused
    if entry is not None and entry.is_file:
        if open_path(ctx, entry.path):
            return
        suffix = entry.path.suffix.lower()
        action = None
        if suffix in (".eml", ".msg"):
            action = "mail_archive.open"
        elif suffix in (".db", ".sqlite", ".sqlite3", ".db3"):
            action = "sqlite_archive.open"
        if action and action in ctx.action_names():
            previous_path = ctx.pane.path
            ctx.invoke(action)
            if ctx.pane.path != previous_path:
                return
    ctx.invoke("view_file")
