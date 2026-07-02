@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
"D:\Python Worksheet\.venv\Scripts\python.exe" "%~dp0order_watcher.py" test
pause
