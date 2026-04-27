@echo off
:: Runs silently as a scheduled task (no console window)
setlocal
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
set "AGENT_DIR=%~dp0"
if "%AGENT_DIR:~-1%"=="\" set "AGENT_DIR=%AGENT_DIR:~0,-1%"

set "PYEXE=%AGENT_DIR%\python\python.exe"
set "AGENTEXE=%AGENT_DIR%\ThermopacAgent.exe"

if exist "%PYEXE%" (
    start "" /B "%PYEXE%" "%AGENT_DIR%\agent\main.py"
) else if exist "%AGENTEXE%" (
    start "" /B "%AGENTEXE%"
)
endlocal
