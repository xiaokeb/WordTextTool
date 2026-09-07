@echo off
setlocal
cd /d "%~dp0"
where conda >nul 2>nul
if errorlevel 1 (
  echo [ERROR] Conda is not on PATH. Run this script from a Conda terminal.
  exit /b 1
)
call conda run --no-capture-output -n WordMod python -m pip check
if errorlevel 1 exit /b 1
call conda run --no-capture-output -n WordMod python -m PyInstaller --noconfirm --clean build.spec
if errorlevel 1 exit /b 1
copy /Y "README.md" "dist\README.md" >nul
if errorlevel 1 exit /b 1
echo [OK] Build files are in dist.
exit /b 0
