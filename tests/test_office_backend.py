import json
import subprocess
from types import SimpleNamespace
import pytest
from word_text_tool import office_backend
from word_text_tool.safe_output import OperationError


def test_timeout_stops_only_owned_process_and_helper(tmp_path, monkeypatch):
    calls = []
    class Process:
        def wait(self, timeout):
            if not calls:
                raise subprocess.TimeoutExpired("test", timeout)
        def kill(self):
            calls.append("helper")
    monkeypatch.setattr(office_backend.subprocess, "Popen", lambda *args, **kwargs: Process())
    monkeypatch.setattr(office_backend, "stop_owned_office", lambda p: calls.append(p))
    with pytest.raises(OperationError, match="超时"):
        office_backend.run_office({}, tmp_path, timeout=0.01)
    assert calls == [tmp_path / "owner.json", "helper"]


def test_session_does_not_quit_reused_user_instance(tmp_path, monkeypatch):
    import pythoncom
    import win32process
    from win32com.client import dynamic
    calls = []
    app = SimpleNamespace(Hwnd=123, Quit=lambda: calls.append("quit"))
    monkeypatch.setattr(win32process, "EnumProcesses", lambda: [777])
    monkeypatch.setattr(win32process, "GetWindowThreadProcessId", lambda hwnd: (1, 777))
    monkeypatch.setattr(pythoncom, "CoCreateInstance", lambda *args: object())
    monkeypatch.setattr(dynamic, "Dispatch", lambda obj: app)
    with pytest.raises(OperationError, match="复用"):
        with office_backend.OfficeSession("wps", tmp_path / "owner.json"):
            pytest.fail("reused user instance accepted")
    assert calls == []
    assert not (tmp_path / "owner.json").exists()


@pytest.fixture
def private_office(monkeypatch):
    """Protocol fake for identity/security. Does not represent a real Excel run."""
    import pythoncom
    import win32api
    import win32process
    from win32com.client import dynamic
    calls = []
    book = SimpleNamespace(ReadOnly=False, Close=lambda **kw: calls.append(("close", kw)))
    def open_book(path, **kwargs):
        calls.append(("open", path, kwargs))
        return book
    app = SimpleNamespace(Hwnd=123, Version="16.0", Workbooks=SimpleNamespace(Count=0, Open=open_book),
                          Quit=lambda: calls.append(("quit",)))
    state = {"executable": "C:/Microsoft Office/EXCEL.EXE"}
    monkeypatch.setattr(win32process, "EnumProcesses", lambda: [])
    monkeypatch.setattr(win32process, "GetWindowThreadProcessId", lambda hwnd: (1, 777))
    monkeypatch.setattr(pythoncom, "CoCreateInstance", lambda *args: calls.append(("create", args)) or object())
    monkeypatch.setattr(dynamic, "Dispatch", lambda obj: state.get("app", app))
    monkeypatch.setattr(win32api, "OpenProcess", lambda *args: SimpleNamespace(Close=lambda: None))
    monkeypatch.setattr(win32process, "GetModuleFileNameEx", lambda *args: state["executable"])
    monkeypatch.setattr(win32process, "GetProcessTimes", lambda h: {"CreationTime": "test"})
    return app, book, state, calls


def test_excel_identity_secure_open_and_owned_cleanup(tmp_path, private_office):
    import pythoncom
    app, book, state, calls = private_office
    with office_backend.OfficeSession("excel", tmp_path / "owner.json") as session:
        assert session.engine == "excel"
        assert session.identity == "Microsoft Excel 16.0"
        assert app.AutomationSecurity == 3 and app.EnableEvents is False
        assert app.AskToUpdateLinks is False
        assert session.open(tmp_path / "input.xlsx") is book
    create = next(c for c in calls if c[0] == "create")
    assert create[1][0] == "Excel.Application"
    assert create[1][2] == pythoncom.CLSCTX_LOCAL_SERVER
    opened = next(c for c in calls if c[0] == "open")
    assert opened[2]["UpdateLinks"] == 0
    assert opened[2]["ReadOnly"] is False
    assert opened[2]["Password"] == ""
    assert calls[-2:] == [("close", {"SaveChanges": False}), ("quit",)]


@pytest.mark.parametrize("requested,actual", [("excel", "et.exe"), ("wps", "EXCEL.EXE"), ("excel", "unrelated.exe")])
def test_explicit_engine_never_accepts_different_process(tmp_path, private_office, requested, actual):
    app, book, state, calls = private_office
    state["executable"] = "C:/Office/" + actual
    with pytest.raises(OperationError, match="实际办公软件不符"):
        with office_backend.OfficeSession(requested, tmp_path / "owner.json"):
            pytest.fail("wrong engine accepted")
    assert not any(c[0] == "open" for c in calls)


def test_auto_reports_actual_wps_identity_for_excel_alias(tmp_path, private_office):
    app, book, state, calls = private_office
    state["executable"] = "C:/WPS/et.exe"
    with office_backend.OfficeSession("auto", tmp_path / "owner.json") as session:
        assert session.engine == "wps"
        assert session.identity.startswith("WPS 表格")


def test_excel_readonly_temporary_book_refused_and_closed(tmp_path, private_office):
    app, book, state, calls = private_office
    book.ReadOnly = True
    with pytest.raises(OperationError, match="不可写"):
        with office_backend.OfficeSession("excel", tmp_path / "owner.json") as session:
            session.open(tmp_path / "input.xlsx")
    assert calls[-2:] == [("close", {"SaveChanges": False}), ("quit",)]


def test_session_refuses_restored_user_workbook_before_changing_app(tmp_path, monkeypatch):
    import pythoncom
    import win32api
    import win32process
    from win32com.client import dynamic
    calls = []
    app = SimpleNamespace(Hwnd=123, Workbooks=SimpleNamespace(Count=1), Quit=lambda: calls.append("quit"))
    monkeypatch.setattr(win32process, "EnumProcesses", lambda: [])
    monkeypatch.setattr(win32process, "GetWindowThreadProcessId", lambda hwnd: (1, 777))
    monkeypatch.setattr(pythoncom, "CoCreateInstance", lambda *args: object())
    monkeypatch.setattr(dynamic, "Dispatch", lambda obj: app)
    monkeypatch.setattr(win32api, "OpenProcess", lambda *args: SimpleNamespace(Close=lambda: None))
    monkeypatch.setattr(win32process, "GetModuleFileNameEx", lambda *args: "C:/WPS/et.exe")
    monkeypatch.setattr(win32process, "GetProcessTimes", lambda h: {"CreationTime": "test"})
    with pytest.raises(OperationError, match="其他工作簿"):
        with office_backend.OfficeSession("wps", tmp_path / "owner.json"):
            pytest.fail("restored workbook accepted")
    assert not hasattr(app, "DisplayAlerts")
    assert calls == []
    assert not (tmp_path / "owner.json").exists()
