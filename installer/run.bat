@echo off
setlocal

set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"

:: Determine the install directory (same folder as this bat)
set "AGENT_DIR=%~dp0"
if "%AGENT_DIR:~-1%"=="\" set "AGENT_DIR=%AGENT_DIR:~0,-1%"

:: Prefer the bundled Python if present
set "PYEXE=%AGENT_DIR%\python\python.exe"

:: If bundled Python exists, use it to run the agent source
if exist "%PYEXE%" (
    "%PYEXE%" "%AGENT_DIR%\agent\main.py" %*
    goto :done
)

:: Otherwise use the PyInstaller EXE (packed agent)
set "AGENTEXE=%AGENT_DIR%\ThermopacAgent.exe"
if exist "%AGENTEXE%" (
    "%AGENTEXE%" %*
    goto :done
)

echo ERROR: Neither python\python.exe nor ThermopacAgent.exe found in %AGENT_DIR%
echo Please re-run the installer or contact support.
pause
exit /b 1

:done
endlocal
