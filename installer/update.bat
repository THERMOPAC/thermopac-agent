@echo off
setlocal

set INSTALL_DIR=C:\Program Files\ThermopacAgent
set SCRIPT_DIR=%~dp0

echo ============================================================
echo  ThermopacAgent update utility
echo  Copies agent\ and extractor\ into %INSTALL_DIR%
echo ============================================================
echo.

if not exist "%INSTALL_DIR%" (
    echo ERROR: %INSTALL_DIR% does not exist.
    echo        Run the installer first.
    pause
    exit /b 1
)

echo Stopping any running agent process...
taskkill /F /IM python.exe /FI "WINDOWTITLE eq ThermopacAgent*" >nul 2>&1
timeout /t 2 /nobreak >nul

echo Copying agent\ ...
xcopy /E /Y /I "%SCRIPT_DIR%..\agent" "%INSTALL_DIR%\agent" >nul
if errorlevel 1 (
    echo ERROR: Failed to copy agent\
    pause
    exit /b 1
)

echo Copying extractor\ ...
xcopy /E /Y /I "%SCRIPT_DIR%..\extractor" "%INSTALL_DIR%\extractor" >nul
if errorlevel 1 (
    echo ERROR: Failed to copy extractor\
    pause
    exit /b 1
)

echo.
echo Update complete.  Restart the agent with:
echo   "%INSTALL_DIR%\run.bat"
echo.
pause
endlocal
