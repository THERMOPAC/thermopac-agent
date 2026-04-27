@echo off
REM ============================================================
REM  Thermopac Drawing Structuring Agent v1.0.21
REM  INSTALL UPDATE SCRIPT
REM  --
REM  Run as Administrator from the extracted ZIP folder.
REM  Copies all Python source files to the installed location,
REM  overwriting old versions. Preserves config.ini (your settings).
REM ============================================================

title ThermopacStructurer Updater v1.0.21

REM ── Require Administrator ────────────────────────────────────────────────────
net session >nul 2>&1
if %errorLevel% NEQ 0 (
    echo.
    echo  [ERROR] This script must be run as Administrator.
    echo.
    echo  Right-click install_update.bat and choose "Run as administrator".
    echo.
    pause
    exit /b 1
)

REM ── Detect install path ───────────────────────────────────────────────────────
set "INSTALL_DIR=C:\Program Files\ThermopacStructuringAgent"

if not exist "%INSTALL_DIR%\" (
    echo.
    echo  [ERROR] Install directory not found:
    echo         %INSTALL_DIR%
    echo.
    echo  Run the full installer first, then use this script to update.
    echo.
    pause
    exit /b 1
)

REM ── Locate this script's directory (the extracted ZIP root) ───────────────────
set "SRC=%~dp0"
if "%SRC:~-1%"=="\" set "SRC=%SRC:~0,-1%"

echo.
echo  Thermopac Drawing Structuring Agent — Update to v1.0.21
echo  --------------------------------------------------------
echo  Source : %SRC%
echo  Target : %INSTALL_DIR%
echo.

REM ── Copy agent Python source (overwrite) ────────────────────────────────────
echo  Updating agent\...
xcopy /E /I /Y "%SRC%\agent"      "%INSTALL_DIR%\agent"      >nul 2>&1

echo  Updating extractor\...
xcopy /E /I /Y "%SRC%\extractor"  "%INSTALL_DIR%\extractor"  >nul 2>&1

echo  Updating structurer\...
xcopy /E /I /Y "%SRC%\structurer" "%INSTALL_DIR%\structurer" >nul 2>&1

REM ── Copy bat/ps1 helper scripts (overwrite) ──────────────────────────────────
for %%F in (
    run.bat
    fix_appdata_url.ps1
    set_testing_mode.bat
    set_dev_url.bat
    set_prod_url.bat
    set_node_token.bat
) do (
    if exist "%SRC%\%%F" (
        echo  Updating %%F...
        copy /Y "%SRC%\%%F" "%INSTALL_DIR%\%%F" >nul 2>&1
    )
)

REM ── Preserve config.ini (do NOT overwrite user settings) ─────────────────────
echo.
echo  config.ini preserved (your settings were not changed).
echo.

echo  ============================================================
echo   Update complete!  ThermopacStructuringAgent is now v1.0.21
echo  ============================================================
echo.
echo  You can now close this window and restart the agent.
echo.
pause
