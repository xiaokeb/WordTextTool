@echo off
setlocal
cd /d "%~dp0"
where conda >nul 2>nul
if errorlevel 1 (
  echo [ERROR] Conda is not on PATH. Run this script from a Conda terminal.
  exit /b 1
)
call conda run --no-capture-output -n WordMod python scripts\build_candidate.py
if errorlevel 1 exit /b 1
echo [OK] Candidate files are in dist\V2.0-candidate.
exit /b 0
