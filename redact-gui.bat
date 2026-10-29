@echo off
cd /d "%~dp0"
where python >nul 2>nul
if errorlevel 1 (
  echo Python not found. Please install Python 3.11 and ensure it is on PATH.
  pause
  exit /b 1
)
python -m local_redact.launch_gui %*
