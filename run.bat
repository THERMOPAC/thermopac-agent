@echo off
REM ============================================================
REM  Thermopac Drawing Structuring Agent v1.0.21
REM  THERMOPAC ERP | SolidWorks WRITE Agent | Phase 1
REM  --
REM  Uses bundled venv Python if available, falls back to system Python.
REM  The Inno Setup installer places the bundled python\ folder here.
REM ============================================================
title ThermopacStructurer v1.0.21
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8

set "AGENT_DIR=%~dp0"
if "%AGENT_DIR:~-1%"=="\" set "AGENT_DIR=%AGENT_DIR:~0,-1%"

REM -- Change to agent root so Python module imports resolve correctly
cd /d "%AGENT_DIR%"

REM ── Auto-correct APPDATA api_url before every launch ─────────────────────
REM    fix_appdata_url.ps1 replaces the production URL with the dev server URL
REM    Preserves node_token and all other settings. No admin rights needed.
PowerShell -NoProfile -ExecutionPolicy Bypass -File "%AGENT_DIR%\fix_appdata_url.ps1"

REM -- Prefer bundled Python (installed by Inno Setup or build-installer.bat)
set "PYEXE=%AGENT_DIR%\python\python.exe"

REM -- Fall back to venv Python (bootstrap.bat setup)
if not exist "%PYEXE%" set "PYEXE=%AGENT_DIR%\venv\Scripts\python.exe"

REM -- Fall back to system Python
if not exist "%PYEXE%" set "PYEXE=python"

echo.
echo  ThermopacStructurer -- SolidWorks Drawing Structuring Agent
echo  THERMOPAC ERP Integration  ^|  Phase 1  ^|  v1.0.21
echo.

"%PYEXE%" "%AGENT_DIR%\agent\main_structurer.py" "%AGENT_DIR%\config.ini" %*
pause
