@echo off
cd /d "%~dp0"
python -m local_redact.gui %*
