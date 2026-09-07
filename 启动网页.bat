@echo off
chcp 65001 >nul
title QQ群成员导出工具 - 网页版
cd /d "%~dp0"
python -m pip install flask psutil --quiet >nul 2>&1
start /b cmd /c "timeout /t 2 /nobreak >nul & start "" http://127.0.0.1:8123"
python app.py
pause
