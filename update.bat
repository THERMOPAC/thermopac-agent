@echo off
setlocal
title ThermopacAgent v1.0.31 - Updater

echo ============================================================
echo  ThermopacAgent v1.0.31 - Auto Updater
echo  Updates an existing installation in-place.
echo ============================================================
echo.

:: ── Detect install directory ─────────────────────────────────────────────────
set "INSTALL_DIR=C:\Program Files\ThermopacAgent"

if not exist "%INSTALL_DIR%" (
    echo ERROR: No existing installation found at:
    echo   %INSTALL_DIR%
    echo.
    echo Please run the full installer first.
    echo Download from: https://thermopac-communication-thermopacllp.replit.app
    echo.
    pause
    exit /b 1
)

echo Install directory: %INSTALL_DIR%
echo.

:: ── Stop any running agent ────────────────────────────────────────────────────
echo Stopping any running agent processes...
schtasks /End /TN "ThermopacAgent" >nul 2>&1
taskkill /F /FI "WINDOWTITLE eq ThermopacAgent*" /T >nul 2>&1
taskkill /F /FI "IMAGENAME eq ThermopacAgent.exe" /T >nul 2>&1
taskkill /F /FI "IMAGENAME eq python.exe" /T >nul 2>&1
taskkill /F /FI "IMAGENAME eq pythonw.exe" /T >nul 2>&1
timeout /t 3 /nobreak >nul
echo Done.
echo.

:: ── Copy updated files ────────────────────────────────────────────────────────
echo Copying agent\  ...
xcopy /E /I /Y "%~dp0agent" "%INSTALL_DIR%\agent\" >nul
if errorlevel 1 ( echo   WARNING: agent copy had errors )

echo Copying extractor\  ...
xcopy /E /I /Y "%~dp0extractor" "%INSTALL_DIR%\extractor\" >nul
if errorlevel 1 ( echo   WARNING: extractor copy had errors )

echo Copying run.bat  ...
copy /Y "%~dp0run.bat" "%INSTALL_DIR%\run.bat" >nul
if errorlevel 1 ( echo   WARNING: run.bat copy had errors )

echo.
echo ============================================================
echo  Update complete!  v1.0.31 is now installed.
echo ============================================================
echo.

set /p LAUNCH="Launch agent now? [Y/N]: "
if /i "%LAUNCH%"=="Y" (
    echo Starting agent...
    start "" "%INSTALL_DIR%\run.bat"
)

echo.
echo You can start the agent any time from:
echo   %INSTALL_DIR%\run.bat
echo   or Start Menu ^> ThermopacAgent
echo.
pause
endlocal
