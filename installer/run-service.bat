@echo off
:: ThermopacStructurer — silent service launcher (scheduled task / no console)
setlocal
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
set "AGENT_DIR=%~dp0"
if "%AGENT_DIR:~-1%"=="\" set "AGENT_DIR=%AGENT_DIR:~0,-1%"

set "PYEXE=%AGENT_DIR%\python\python.exe"
if not exist "%PYEXE%" set "PYEXE=%AGENT_DIR%\venv\Scripts\python.exe"
if not exist "%PYEXE%" set "PYEXE=python"

start "" /B "%PYEXE%" "%AGENT_DIR%\agent\main_structurer.py" "%AGENT_DIR%\config.ini"
endlocal
