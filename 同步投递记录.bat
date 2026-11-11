@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo 还没有虚拟环境。请先按 README 的「下载后怎么用」执行 python -m venv .venv 和 pip install。
    pause
    exit /b 1
)
".venv\Scripts\python.exe" job_sync.py
if errorlevel 1 (
    echo.
    echo 程序异常退出，上面是错误信息。
)
pause
