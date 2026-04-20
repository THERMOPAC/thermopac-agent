@echo off
setlocal

title ThermopacAgent v1.0.29

set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"

:: Determine the install directory (same folder as this bat)
set "AGENT_DIR=%~dp0"
if "%AGENT_DIR:~-1%"=="\" set "AGENT_DIR=%AGENT_DIR:~0,-1%"

:: Prefer the bundled Python if present
set "PYEXE=%AGENT_DIR%\python\python.exe"

if exist "%PYEXE%" (
    echo [ThermopacAgent] Starting agent via bundled Python...
    "%PYEXE%" "%AGENT_DIR%\agent\main.py" %*
    if errorlevel 1 (
        echo.
        echo ============================================================
        echo  Agent exited with an error  (errorlevel %errorlevel%)
        echo  Scroll up to see the traceback.
        echo  Check config.ini and re-run, or contact support.
        echo ============================================================
        pause
    )
    goto :done
)

:: Fallback to PyInstaller EXE
set "AGENTEXE=%AGENT_DIR%\ThermopacAgent.exe"
if exist "%AGENTEXE%" (
    echo [ThermopacAgent] Starting agent via packed EXE...
    "%AGENTEXE%" %*
    if errorlevel 1 (
        echo.
        echo ============================================================
        echo  Agent exited with an error  (errorlevel %errorlevel%)
        echo ============================================================
        pause
    )
    goto :done
)

echo.
echo ============================================================
echo  ERROR: Neither python\python.exe nor ThermopacAgent.exe
echo  found in: %AGENT_DIR%
echo  Please re-run the installer.
echo ============================================================
pause
exit /b 1

:done
endlocal
