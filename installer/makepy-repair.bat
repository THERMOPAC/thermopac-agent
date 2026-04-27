@echo off
setlocal
echo ThermopacStructurer -- Repair SolidWorks COM Cache
echo ===================================================
echo.

set "AGENT_DIR=%~dp0"
if "%AGENT_DIR:~-1%"=="\" set "AGENT_DIR=%AGENT_DIR:~0,-1%"

set "PYEXE=%AGENT_DIR%\python\python.exe"
if not exist "%PYEXE%" set "PYEXE=%AGENT_DIR%\venv\Scripts\python.exe"
if not exist "%PYEXE%" (
    echo ERROR: Python not found.
    echo Run bootstrap.bat or reinstall the agent.
    pause
    exit /b 1
)

echo Detecting SolidWorks version from registry...
set "SW_PROGID="
for /L %%Y in (32,-1,27) do (
    reg query "HKCR\SldWorks.Application.%%Y" >nul 2>&1
    if not errorlevel 1 (
        if "!SW_PROGID!"=="" set "SW_PROGID=SldWorks.Application.%%Y"
    )
)

if "!SW_PROGID!"=="" (
    echo ERROR: SolidWorks not found in registry.
    echo Install SolidWorks and try again.
    pause
    exit /b 1
)

echo Found: !SW_PROGID!
echo.
echo Running makepy (may take 10-30 seconds)...
"%PYEXE%" -m win32com.client.makepy "!SW_PROGID!"
if errorlevel 1 (
    echo.
    echo Method A failed -- trying gencache approach...
    "%PYEXE%" -c "import win32com.client; win32com.client.gencache.EnsureDispatch('!SW_PROGID!')" 2>nul
)
echo.
echo Done. SolidWorks COM cache rebuilt.
echo Restart ThermopacStructurer.
pause
endlocal
