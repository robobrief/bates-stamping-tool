@echo off
setlocal
cd /d "%~dp0"
call .venv\Scripts\activate.bat
python bundle_gui.py
pause
