@echo off
setlocal

title ThermopacAgent v1.0.30

set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"

:: Determine the install directory (same folder as this bat)
set "AGENT_DIR=%~dp0"
if "%AGENT_DIR:~-1%"=="\" set "AGENT_DIR=%AGENT_DIR:~0,-1%"

:: ── Priority 1: PyInstaller EXE (produced by CI installer) ─────────────────
set "AGENTEXE=%AGENT_DIR%\ThermopacAgent.exe"
if exist "%AGENTEXE%" (
    echo [ThermopacAgent] Starting agent via ThermopacAgent.exe...
    "%AGENTEXE%" %*
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

:: ── Priority 2: Bundled Python + raw source (dev / ZIP installs) ───────────
set "PYEXE=%AGENT_DIR%\python\python.exe"
set "MAINPY=%AGENT_DIR%\agent\main.py"
if exist "%PYEXE%" (
    if exist "%MAINPY%" (
        echo [ThermopacAgent] Starting agent via bundled Python + source...
        "%PYEXE%" "%MAINPY%" %*
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
    ) else (
        echo [ThermopacAgent] ERROR: python\python.exe found but agent\main.py is missing.
        echo  This usually means ThermopacAgent.exe is also missing.
        echo  Please re-run the installer.
        pause
        exit /b 1
    )
)

echo.
echo ============================================================
echo  ERROR: Neither ThermopacAgent.exe nor python\python.exe
echo  found in: %AGENT_DIR%
echo  Please re-run the installer.
echo ============================================================
pause
exit /b 1

:done
endlocal
