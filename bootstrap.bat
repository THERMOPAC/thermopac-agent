@echo off
REM ============================================================
REM  Thermopac Drawing Structuring Agent — Bootstrap installer
REM  Version: v1.0.5  |  Phase 1
REM  Requires: Windows 10/11 x64, Python 3.9+ already installed
REM  Run as Administrator for best results
REM ============================================================
setlocal enabledelayedexpansion

echo.
echo  =========================================================
echo   Thermopac Drawing Structuring Agent  v1.0.5
echo   Bootstrap Installer
echo   THERMOPAC ERP ^| SolidWorks WRITE Agent ^| Phase 1
echo  =========================================================
echo.
echo  IMPORTANT: This is a WRITE agent.
echo  It creates and updates .slddrw files from DDS job data.
echo  It must run on a PC with a licensed SolidWorks installation.
echo.

REM ── Locate Python ──────────────────────────────────────────
set PYTHON=
for %%P in (python3.11 python3.10 python3.9 python) do (
    where %%P >nul 2>&1
    if !errorlevel!==0 (
        set PYTHON=%%P
        goto :found_python
    )
)
echo ERROR: Python not found in PATH.
echo.
echo Install Python 3.11 from https://python.org/downloads/
echo Make sure to check "Add Python to PATH" during installation.
echo.
pause
exit /b 1

:found_python
for /f "tokens=*" %%V in ('!PYTHON! --version 2^>^&1') do set PYVER=%%V
echo Found: !PYVER! (!PYTHON!)
echo.

REM ── Set install paths ────────────────────────────────────────
set AGENT_DIR=%~dp0
set VENV_DIR=%AGENT_DIR%venv
set DATA_DIR=C:\ThermopacStructurer
set TEMP_DIR=%DATA_DIR%\temp
set LOGS_DIR=%DATA_DIR%\logs
set STAGING_DIR=C:\ThermopacStaging\drawings

echo Install locations:
echo   Agent source:    %AGENT_DIR%
echo   Python venv:     %VENV_DIR%
echo   Data / logs:     %DATA_DIR%
echo   Staging root:    %STAGING_DIR%
echo.

REM ── Create data directories ──────────────────────────────────
echo Creating data directories...
if not exist "%TEMP_DIR%"    mkdir "%TEMP_DIR%"
if not exist "%LOGS_DIR%"    mkdir "%LOGS_DIR%"
if not exist "%STAGING_DIR%" mkdir "%STAGING_DIR%"
echo Done.
echo.

REM ── Create virtual environment ────────────────────────────────
echo Creating Python virtual environment...
if exist "%VENV_DIR%" (
    echo   (venv already exists — skipping creation)
) else (
    !PYTHON! -m venv "%VENV_DIR%"
    if !errorlevel! neq 0 (
        echo ERROR: Failed to create virtual environment
        pause & exit /b 1
    )
    echo   Created: %VENV_DIR%
)
echo.

REM ── Install dependencies ──────────────────────────────────────
echo Installing Python dependencies (pywin32, requests)...
"%VENV_DIR%\Scripts\python.exe" -m pip install --quiet --upgrade pip
"%VENV_DIR%\Scripts\pip.exe" install --quiet pywin32 requests

if !errorlevel! neq 0 (
    echo ERROR: pip install failed. Check your internet connection.
    pause & exit /b 1
)

REM ── Run pywin32 post-install script (registers COM DLLs) ──────
echo.
echo Running pywin32 post-install (registers COM DLLs for SolidWorks)...
"%VENV_DIR%\Scripts\python.exe" "%VENV_DIR%\Scripts\pywin32_postinstall.py" -install 2>nul
echo Done.
echo.

REM ── Write start_structurer.bat ────────────────────────────────
echo Writing start_structurer.bat launcher...
(
    echo @echo off
    echo title ThermopacStructurer v1.0.5 - THERMOPAC ERP
    echo set AGENT_DIR=%%~dp0
    echo "%VENV_DIR%\Scripts\python.exe" "%%AGENT_DIR%%agent\main_structurer.py"
) > "%AGENT_DIR%start_structurer.bat"
echo   Created: %AGENT_DIR%start_structurer.bat
echo.

REM ── Write test.bat ────────────────────────────────────────────
(
    echo @echo off
    echo REM Thermopac Drawing Structuring Agent — Self-Test
    echo setlocal
    echo set AGENT_DIR=%%~dp0
    echo "%VENV_DIR%\Scripts\python.exe" "%%AGENT_DIR%%agent\main_structurer.py" --test
    echo echo.
    echo echo Self-test complete. Check %LOGS_DIR%\agent.log for details.
    echo pause
    echo endlocal
) > "%AGENT_DIR%test.bat"
echo   Created: %AGENT_DIR%test.bat
echo.

REM ── Create Desktop shortcut ───────────────────────────────────
echo Creating Desktop shortcut...
set SHORTCUT_TARGET=%AGENT_DIR%start_structurer.bat
set SHORTCUT_PATH=%USERPROFILE%\Desktop\ThermopacStructurer.lnk
powershell -Command ^
  "$s = (New-Object -COM WScript.Shell).CreateShortcut('%SHORTCUT_PATH%'); $s.TargetPath = '%SHORTCUT_TARGET%'; $s.WorkingDirectory = '%AGENT_DIR%'; $s.Description = 'Thermopac Drawing Structuring Agent v1.0.5'; $s.Save()"
if exist "%SHORTCUT_PATH%" (
    echo   Created: Desktop\ThermopacStructurer.lnk
) else (
    echo   (shortcut creation skipped - requires desktop access)
)
echo.

REM ── Configuration reminder ────────────────────────────────────
echo ============================================================
echo   CONFIGURATION REQUIRED  (read INSTALL.md for full guide)
echo ============================================================
echo.
echo Edit config.ini in this folder:
echo   %AGENT_DIR%config.ini
echo.
echo Required settings:
echo.
echo   [cloud]
echo   node_id    = PC-DESIGN-01          ^<-- your node ID
echo   node_token = REPLACE_ME            ^<-- your node token
echo.
echo   [solidworks]
echo   solidworks_version = 2019          ^<-- your installed version
echo.
echo   [structurer]
echo   template_path = C:\SolidWorks Templates\Standard_A1.drwdot
echo   staging_root  = C:\ThermopacStaging\drawings
echo.
echo IMPORTANT: Set template_path to your standard .drwdot template.
echo            The agent will FAIL pre-flight if template_path is empty.
echo.

REM ── Run self-test ─────────────────────────────────────────────
echo ============================================================
echo   SELF-TEST (running now)
echo ============================================================
echo.
echo NOTE: Authentication will fail until you edit config.ini.
echo       template_path must also be set before jobs can run.
echo.
"%VENV_DIR%\Scripts\python.exe" "%AGENT_DIR%agent\main_structurer.py" --test
echo.
echo ============================================================
echo   Bootstrap complete!
echo ============================================================
echo.
echo To start the structuring agent:
echo   Double-click Desktop shortcut "ThermopacStructurer"
echo   OR run: start_structurer.bat
echo.
echo To run self-test again:
echo   run: test.bat
echo.
echo Logs:  %LOGS_DIR%\agent.log
echo.
pause
endlocal
