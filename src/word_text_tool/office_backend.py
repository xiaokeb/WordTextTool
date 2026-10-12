"""Isolated, owned desktop Office instances; never attach to a user's workbook."""
import json
import os
from pathlib import Path
import subprocess
import sys

from word_text_tool.safe_output import OperationError


def run_office(request: dict, directory: Path, timeout: float = 120, executable=None) -> dict:
    request_path = directory / "request.json"
    result_path = directory / "result.json"
    request_path.write_text(json.dumps(request, ensure_ascii=False), encoding="utf-8")
    entry = Path(__file__).resolve().parents[1] / "main.py"
    command = [str(executable or sys.executable)]
    if executable is None and not getattr(sys, "frozen", False):
        command.append(str(entry))
    command.extend(["--office-job", str(request_path), str(result_path)])
    process = subprocess.Popen(command, stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                               stderr=subprocess.DEVNULL, creationflags=subprocess.CREATE_NO_WINDOW)
    try:
        process.wait(timeout=timeout)
    except subprocess.TimeoutExpired as exc:
        # Only this helper and the independently identified Office process belong
        # to this operation. No process-name termination or global Office cleanup.
        stop_owned_office(directory / "owner.json")
        process.kill()
        process.wait(timeout=10)
        raise OperationError("办公软件处理超时，本次文件未完成") from exc
    if not result_path.exists():
        stop_owned_office(directory / "owner.json")
        raise OperationError("办公软件任务异常退出，请检查引擎安装及自动化权限")
    result = json.loads(result_path.read_text(encoding="utf-8"))
    if not result.get("ok"):
        raise OperationError(result.get("error", "办公软件操作失败"))
    return result


def stop_owned_office(owner_path: Path):
    """Terminate only the still-identical, newly created private process."""
    import win32api
    import win32process
    import pywintypes
    try:
        owner = json.loads(owner_path.read_text(encoding="utf-8"))
        handle = win32api.OpenProcess(0x1000 | 1, False, owner["pid"])
        try:
            if str(win32process.GetProcessTimes(handle)["CreationTime"]) == owner["created"]:
                win32api.TerminateProcess(handle, 1)
        finally:
            handle.Close()
    except (OSError, ValueError, KeyError, pywintypes.error):
        pass


class OfficeSession:
    def __init__(self, engine, owner_path):
        self.engine = engine
        self.owner_path = Path(owner_path)
        self.app = None
        self.book = None
        self.identity = ""
        self.owned = False

    def __enter__(self):
        import pythoncom
        import win32api
        import win32process
        from win32com.client import dynamic
        pythoncom.CoInitialize()
        choices = [("excel", "Excel.Application"), ("wps", "ket.Application")]
        if self.engine != "auto":
            choices = [choice for choice in choices if choice[0] == self.engine]
        errors = []
        last_error = None
        for label, progid in choices:
            try:
                before = set(win32process.EnumProcesses())
                dispatch = pythoncom.CoCreateInstance(progid, None,
                                                     pythoncom.CLSCTX_LOCAL_SERVER, pythoncom.IID_IDispatch)
                self.app = dynamic.Dispatch(dispatch)
                hwnd = int(self.app.Hwnd)
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                if not pid or pid in before:
                    self.app = None
                    raise OperationError("引擎复用了已有进程，不能安全隔离用户工作簿")
                handle = win32api.OpenProcess(0x0400 | 0x10, False, pid)
                try:
                    executable = win32process.GetModuleFileNameEx(handle, 0).lower()
                    created = str(win32process.GetProcessTimes(handle)["CreationTime"])
                finally:
                    handle.Close()
                # A new process can still restore user documents at startup.
                # Establish ownership only after confirming an empty instance,
                # before changing visibility, alerts, or any other app state.
                if self.app.Workbooks.Count != 0:
                    self.app = None
                    raise OperationError("新实例已载入其他工作簿，不能安全处理")
                self.owned = True
                self.owner_path.write_text(json.dumps({"pid": pid, "created": created}), encoding="utf-8")
                actual = "wps" if Path(executable).name == "et.exe" else "excel" if Path(executable).name == "excel.exe" else ""
                if not actual or (self.engine != "auto" and actual != self.engine):
                    raise OperationError("注册接口与实际办公软件不符，请选择实际安装的引擎")
                self.engine = actual
                self.identity = f"{'WPS 表格' if actual == 'wps' else 'Microsoft Excel'} {self.app.Version}"
                self.app.Visible = False
                self.app.DisplayAlerts = False
                self.app.AutomationSecurity = 3
                self.app.EnableEvents = False
                self.app.AskToUpdateLinks = False
                if int(self.app.AutomationSecurity) != 3 or bool(self.app.EnableEvents):
                    raise OperationError("办公软件无法禁用宏或事件，已拒绝打开文件")
                return self
            except Exception as exc:
                last_error = exc
                errors.append(f"{label}：{str(exc) if isinstance(exc, OperationError) else '未安装或自动化不可用'}")
                self._close()
        pythoncom.CoUninitialize()
        raise OperationError("；".join(errors)) from last_error

    def open(self, source):
        self.book = self.app.Workbooks.Open(str(source), UpdateLinks=0, ReadOnly=False,
                                           Password="", WriteResPassword="", IgnoreReadOnlyRecommended=True,
                                           Notify=False, AddToMru=False)
        if self.book.ReadOnly:
            raise OperationError("临时工作簿不可写")
        return self.book

    def _close(self):
        if self.book is not None:
            try:
                self.book.Close(SaveChanges=False)
            except Exception:
                pass
            self.book = None
        if self.owned and self.app is not None:
            try:
                self.app.Quit()
            except Exception:
                pass
        self.app = None
        self.owned = False

    def __exit__(self, *args):
        import pythoncom
        self._close()
        pythoncom.CoUninitialize()


def execute_job(request_path: Path, result_path: Path) -> int:
    from word_text_tool.office_validation import inspect_package, workbook_sheets
    request = json.loads(request_path.read_text(encoding="utf-8"))
    root = request_path.parent.resolve()
    result = {}
    try:
        source = root / "input.xlsx"
        expected_sheets = workbook_sheets(inspect_package(source))
        with OfficeSession(request["engine"], root / "owner.json") as session:
            book = session.open(source)
            if book.Worksheets.Count != len(expected_sheets) or book.Sheets.Count != len(expected_sheets):
                raise OperationError("工作表范围与预检结果不一致")
            for index, (name, _) in enumerate(expected_sheets, 1):
                sheet = book.Worksheets.Item(index)
                if sheet.Name != name or int(sheet.Visible) != -1 or sheet.ProtectContents:
                    raise OperationError("工作表必须全部可见、有效且未保护")
            if request["action"] == "number":
                from word_text_tool.excel_processor import number_workbook
                number_workbook(book, request["number"], request["fixed"], request["fixed_text"])
                book.SaveAs(str(root / "numbered.xlsx"), FileFormat=51)
            elif request["action"] == "pdf":
                for index in range(1, book.Worksheets.Count + 1):
                    page = book.Worksheets.Item(index).PageSetup
                    page.Zoom = False
                    page.FitToPagesWide = 1
                    page.FitToPagesTall = False
                    page.TopMargin = session.app.CentimetersToPoints(1.3)
                    page.BottomMargin = session.app.CentimetersToPoints(1.3)
                    page.LeftMargin = session.app.CentimetersToPoints(0.9)
                    page.RightMargin = session.app.CentimetersToPoints(0.9)
                book.ExportAsFixedFormat(0, str(root / "export.pdf"), 0, True, False)
            else:
                raise OperationError("未知的办公操作")
            result = {"ok": True, "engine": session.engine, "identity": session.identity}
    except Exception as exc:
        result = {"ok": False, "error": str(exc) if isinstance(exc, OperationError) else "办公软件无法完成操作，请检查文件模板、引擎和自动化权限"}
        if request.get("diagnostics"):
            import traceback
            result["diagnostic"] = traceback.format_exc()
    result_path.write_text(json.dumps(result, ensure_ascii=False), encoding="utf-8")
    return 0 if result["ok"] else 1
