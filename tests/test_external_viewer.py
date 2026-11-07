from pathlib import Path
import sys
from types import SimpleNamespace

import pytest

from external_viewer import logic
from external_viewer.logic import _find_viewer, _open_office
from external_viewer.settings import Settings


@pytest.fixture
def context(monkeypatch):
    messages = []
    invoked = []
    monkeypatch.setattr(
        logic, "extension", lambda name: SimpleNamespace(settings=Settings())
    )
    return SimpleNamespace(
        message=messages.append,
        messages=messages,
        invoke=invoked.append,
        invoked=invoked,
        action_names=lambda: [],
        pane=SimpleNamespace(focused=None, path=Path(".")),
    )


def test_rules_are_case_insensitive_and_ordered():
    rules = [
        {"pattern": "*.PDF", "command": ["first"]},
        {"pattern": ["*"], "command": ["second"]},
    ]
    assert _find_viewer("REPORT.pdf", rules) == rules[0]
    assert _find_viewer("other", rules) == rules[1]
    assert _find_viewer("other", []) is None


def test_default_rules_leave_mail_and_sqlite_to_other_extensions():
    for filename in ("test.eml", "test.msg", "test.sqlite", "test.db"):
        assert _find_viewer(filename, Settings.VIEWERS) is None


def test_hook_launches_argv_with_file_as_one_argument(context, monkeypatch):
    calls = []
    monkeypatch.setattr(logic.subprocess, "Popen", calls.append)
    path = Path("a folder") / "report.PDF"
    assert logic.open_with_external_viewer(context, SimpleNamespace(path=path))
    assert calls == [["sumatrapdf", str(path)]]
    assert not context.messages


def test_missing_program_declines_hook_and_reports_error(context, monkeypatch):
    def missing(argv):
        raise FileNotFoundError("sumatrapdf")

    monkeypatch.setattr(logic.subprocess, "Popen", missing)
    assert not logic.open_with_external_viewer(context, SimpleNamespace(path=Path("a.pdf")))
    assert "sumatrapdf" in context.messages[0]


def test_action_falls_back_to_builtin_on_launch_failure(context, monkeypatch):
    monkeypatch.setattr(logic, "_launch", lambda *args: False)
    context.pane.focused = SimpleNamespace(is_file=True, path=Path("a.pdf"))
    logic.view_file_external(context)
    assert context.invoked == ["view_file"]


def test_action_does_not_fall_back_after_success(context, monkeypatch):
    monkeypatch.setattr(logic, "_launch", lambda *args: True)
    context.pane.focused = SimpleNamespace(is_file=True, path=Path("a.pdf"))
    logic.view_file_external(context)
    assert context.invoked == []


def test_unmatched_and_empty_focus_fall_back(context):
    context.pane.focused = SimpleNamespace(is_file=True, path=Path("a.txt"))
    logic.view_file_external(context)
    context.pane.focused = None
    logic.view_file_external(context)
    assert context.invoked == ["view_file", "view_file"]


@pytest.mark.parametrize(("filename", "action"), [
    ("a.eml", "mail_archive.open"),
    ("a.MSG", "mail_archive.open"),
    ("a.db", "sqlite_archive.open"),
    ("a.SQLITE3", "sqlite_archive.open"),
])
def test_archive_action_only_claims_when_pane_changes(context, filename, action):
    context.pane.focused = SimpleNamespace(is_file=True, path=Path(filename))
    context.action_names = lambda: [action]

    def navigate(name):
        context.invoked.append(name)
        context.pane.path = Path("archive")

    context.invoke = navigate
    logic.view_file_external(context)
    assert context.invoked == [action]
    context.invoke = context.invoked.append
    logic.view_file_external(context)
    assert context.invoked[-2:] == [action, "view_file"]
    context.action_names = lambda: []
    logic.view_file_external(context)
    assert context.invoked[-1] == "view_file"


@pytest.mark.parametrize("viewers", [None, [{}], [{"pattern": [], "command": None}]])
def test_invalid_rules_report_and_fall_back(context, monkeypatch, viewers):
    monkeypatch.setattr(logic, "_viewers", lambda: viewers)
    assert not logic.open_path(context, Path("a.pdf"))
    assert "External viewer failed" in context.messages[0]


def test_custom_rules_callables_and_system_association(context, monkeypatch):
    calls = []
    monkeypatch.setattr(logic.os, "startfile", lambda path: calls.append(path), raising=False)
    monkeypatch.setattr(logic, "_viewers", lambda: [
        {"pattern": "*.pdf", "command": lambda ctx, path: calls.append(path)},
        {"pattern": "*.txt", "command": None},
    ])
    assert logic.open_path(context, Path("a.pdf"))
    assert logic.open_path(context, Path("a.txt"))
    assert calls == [Path("a.pdf"), "a.txt"]


def test_callable_can_decline(context, monkeypatch):
    monkeypatch.setattr(logic, "_viewers", lambda: [
        {"pattern": "*", "command": lambda ctx, path: False},
    ])
    assert not logic.open_path(context, Path("a.pdf"))


def test_virtual_file_materialized_with_unique_paths(context, monkeypatch, tmp_path):
    class VirtualFile:
        name = "a.pdf"

        def __str__(self):
            return "zip://archive/a.pdf"

        def read_bytes(self):
            return b"pdf"

    monkeypatch.setattr(logic, "_tmp_root", tmp_path)
    calls = []
    monkeypatch.setattr(logic.subprocess, "Popen", calls.append)
    assert logic.open_path(context, VirtualFile())
    assert logic.open_path(context, VirtualFile())
    targets = [Path(args[-1]) for args in calls]
    assert targets[0] != targets[1]
    assert all(target.read_bytes() == b"pdf" for target in targets)


def test_office_failure_is_reported_and_declines(context, monkeypatch):
    def fail(kind, path):
        raise OSError("Office not installed")

    monkeypatch.setattr(logic, "_open_office", fail)
    assert not logic.open_path(context, Path("old.xls"))
    assert "Office not installed" in context.messages[0]


@pytest.mark.parametrize(("kind", "prog_id", "collection", "arguments"), [
    ("excel", "Excel.Application", "Workbooks", (0, True)),
    ("word", "Word.Application", "Documents", (False, True)),
    ("powerpoint", "PowerPoint.Application", "Presentations", (True, False, True)),
])
def test_office_read_only_and_com_cleanup(monkeypatch, kind, prog_id, collection, arguments):
    calls = []
    pythoncom = SimpleNamespace(
        CoInitialize=lambda: calls.append("init"),
        CoUninitialize=lambda: calls.append("uninit"),
    )
    app = SimpleNamespace(**{
        collection: SimpleNamespace(Open=lambda *args: calls.append(args))
    })
    client = SimpleNamespace(Dispatch=lambda name: calls.append(name) or app)
    monkeypatch.setitem(sys.modules, "pythoncom", pythoncom)
    monkeypatch.setitem(sys.modules, "win32com", SimpleNamespace(client=client))
    monkeypatch.setitem(sys.modules, "win32com.client", client)
    _open_office(kind, Path("old.doc"))
    assert calls == ["init", prog_id, ("old.doc", *arguments), "uninit"]
    assert app.Visible is True


def test_default_mail_route_does_not_claim_event(context):
    assert not logic.open_with_external_viewer(
        context, SimpleNamespace(path=Path("mail.eml"))
    )
