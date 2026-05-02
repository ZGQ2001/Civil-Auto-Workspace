@echo off
REM ─────────────────────────────────────────────────────────────────
REM Civil Auto Workspace — 一键启动
REM ─────────────────────────────────────────────────────────────────

setlocal
chcp 65001 >nul
cd /d "%~dp0"

if not exist .venv (
    echo [!] 未发现 .venv 虚拟环境，请先运行 scripts\setup_env.bat
    pause
    exit /b 1
)

call .venv\Scripts\activate.bat
set PYTHONIOENCODING=utf-8
set PYTHONPATH=%CD%\src;%PYTHONPATH%

python -m civil_auto.main

if errorlevel 1 (
    echo.
    echo [!] 程序异常退出 (errorlevel=%errorlevel%)
    pause
)
endlocal
