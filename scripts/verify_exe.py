"""Start the candidate with a minimal PATH and close only its new window."""
import argparse
import json
import os
from pathlib import Path
import subprocess
import time
import hashlib

import win32api
import win32con
import win32gui
import win32process


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("exe", type=Path)
    args = parser.parse_args()
    exe = args.exe.resolve()
    before = set(win32process.EnumProcesses())
    environment = dict(os.environ)
    windows = Path(os.environ.get("SystemRoot", "C:/Windows"))
    environment["PATH"] = str(windows / "System32") + os.pathsep + str(windows)
    environment.pop("PYTHONPATH", None)
    environment.pop("QT_QPA_PLATFORM", None)
    startup = subprocess.STARTUPINFO()
    startup.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startup.wShowWindow = win32con.SW_HIDE
    process = subprocess.Popen([str(exe)], env=environment, startupinfo=startup)
    found = []
    def inspect(hwnd, _):
        if win32gui.GetWindowText(hwnd) != "Word 文本处理工具":
            return
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        if pid in before:
            return
        handle = win32api.OpenProcess(0x0400 | 0x10, False, pid)
        try:
            if Path(win32process.GetModuleFileNameEx(handle, 0)).resolve() == exe:
                found.append(hwnd)
        finally:
            handle.Close()
    deadline = time.monotonic() + 25
    while time.monotonic() < deadline and not found and process.poll() is None:
        win32gui.EnumWindows(inspect, None)
        time.sleep(0.2)
    if not found:
        process.terminate()
        process.wait(timeout=10)
        raise RuntimeError("Candidate did not create its main window")
    win32gui.PostMessage(found[0], win32con.WM_CLOSE, 0, 0)
    code = process.wait(timeout=15)
    report = {"main_window_created": True, "minimal_path": True, "normal_close_exit_code": code,
              "candidate_sha256": hashlib.sha256(exe.read_bytes()).hexdigest().upper()}
    destination = Path(__file__).resolve().parents[1] / "build" / "exe-verification.json"
    destination.write_text(json.dumps(report, indent=2), encoding="utf-8")
    print(json.dumps(report))
    return code


if __name__ == "__main__":
    raise SystemExit(main())
