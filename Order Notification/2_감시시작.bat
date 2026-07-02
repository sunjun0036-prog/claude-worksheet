@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
title Order Board Watcher - close to stop
"D:\Python Worksheet\.venv\Scripts\python.exe" "%~dp0order_watcher.py" watch
pause
