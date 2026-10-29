@echo off
cd /d "%~dp0"
where python >nul 2>nul
if errorlevel 1 (
  echo Python not found. Please install Python 3.11 and ensure it is on PATH.
  pause
  exit /b 1
)
python -c "import lxml, win32com.client" 1>nul 2>nul
if errorlevel 1 (
  echo [local-redact] Installing requirements.txt ...
  python -m pip install -r "%~dp0requirements.txt"
  if errorlevel 1 (
    echo Install failed. Run: python -m pip install -r requirements.txt
    pause
    exit /b 1
  )
)
python -m local_redact %*
