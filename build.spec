# -*- mode: python ; coding: utf-8 -*-
import os
import sys
from pathlib import Path
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# Restrict DLL discovery in this build process and its children to the active
# environment and Windows. Inherited tool PATHs can supply incompatible ICU DLLs.
environment_root = Path(sys.prefix)
windows_root = Path(os.environ.get("SystemRoot", r"C:\Windows"))
build_dll_directories = [
    environment_root,
    environment_root / "DLLs",
    environment_root / "Library" / "bin",
    environment_root / "Scripts",
    environment_root / "Lib" / "site-packages" / "PyQt6" / "Qt6" / "bin",
    windows_root / "System32",
    windows_root,
]
os.environ["PATH"] = os.pathsep.join(
    str(directory) for directory in build_dll_directories if directory.is_dir()
)

project_root = Path(SPECPATH)
a = Analysis(
    [str(project_root / "src" / "main.py")],
    pathex=[str(project_root / "src")],
    binaries=[],
    datas=collect_data_files("docx"),
    hiddenimports=collect_submodules("docx") + ["pythoncom", "pywintypes", "win32timezone", "win32com.client.dynamic"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["PyQt5", "PySide2", "PySide6", "pytest", "pytestqt", "IPython", "tkinter"],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz, a.scripts, a.binaries, a.datas, [],
    name="Word 文本处理工具",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
)
