@echo off
REM ============================================================
REM  Thermopac Drawing Structuring Agent v1.0.5
REM  Installer Build Pipeline
REM ============================================================
REM  Run from the structurer_pkg\ directory:
REM      installer\build-installer.bat
REM
REM  What this does:
REM    1. Download Python 3.11.9 embeddable into dist\python\
REM    2. Enable site-packages in embeddable Python
REM    3. Install pip into embeddable Python
REM    4. pip install pywin32 + requests into embeddable Python
REM    5. Run pywin32 post-install (registers COM DLLs)
REM    6. Generate SolidWorks COM type library cache (makepy) — 3 methods
REM    7. Copy agent sources into dist\ThermopacStructuringAgent\
REM         Copies: agent\  extractor\  structurer\
REM    8. Write run.bat / run-service.bat / makepy-repair.bat into dist\
REM    9. Write config.ini template (includes [structurer] section)
REM   10. Check for Inno Setup 6 — reports path but does NOT compile .exe
REM         (to compile run:  iscc installer\setup.iss)
REM
REM  Prerequisites (build machine):
REM    - Windows 10+ x64
REM    - SolidWorks 2019-2024 installed (for makepy cache)
REM    - Internet access (to download Python embeddable)
REM    - Inno Setup 6 optional — for .exe compilation after this script
REM
REM  CI mode (GitHub Actions / no SolidWorks):
REM    Set environment variable  CI=true  before running.
REM    In CI mode the SolidWorks detection and makepy steps are skipped.
REM    The installer still packages Python + agent sources; makepy runs at
REM    first agent startup on the target machine.
REM ============================================================

setlocal enabledelayedexpansion

REM Allow CI to inject STRUCTURER_VERSION via env var; fall back to local default
if "%STRUCTURER_VERSION%"=="" set STRUCTURER_VERSION=1.0.24
set PY_VERSION=3.11.9
set PY_ZIP=python-%PY_VERSION%-embed-amd64.zip
set PY_URL=https://www.python.org/ftp/python/%PY_VERSION%/%PY_ZIP%
set GET_PIP_URL=https://bootstrap.pypa.io/get-pip.py

REM Directories (relative to structurer_pkg\)
REM %~dp0 = structurer_pkg\installer\  so %~dp0.. = structurer_pkg\
set PKG_ROOT=%~dp0..
set DIST_DIR=%PKG_ROOT%\dist
set PY_DIR=%DIST_DIR%\python
set AGENT_DIST=%DIST_DIR%\ThermopacStructuringAgent
set OUT_DIR=%PKG_ROOT%\installer_output

echo ============================================================
echo  ThermopacStructuringAgent v%STRUCTURER_VERSION% -- Installer Build
echo  THERMOPAC ERP  ^|  SolidWorks Drawing Structuring Agent
echo  Phase 1 -- WRITE ONLY
echo ============================================================
echo.

REM ── CI mode detection ─────────────────────────────────────────────────────────
if /i "!CI!"=="true" (
    echo [CI] CI=true detected -- skipping SolidWorks detection and makepy.
    echo [CI] The installer will still package Python + agent sources.
    echo [CI] SolidWorks COM cache will be generated at first agent startup.
    echo.
    set SW_PROGID=SldWorks.Application.27
    set SW_VERSION=2019
    goto :sw_done
)

REM ── Detect SolidWorks (local build only) ──────────────────────────────────────
set SW_PROGID=
set SW_VERSION=0
for %%v in (32 31 30 29 28 27) do (
    if "!SW_PROGID!"=="" (
        reg query "HKCR\SldWorks.Application.%%v" >nul 2>&1
        if !ERRORLEVEL!==0 (
            set SW_PROGID=SldWorks.Application.%%v
            if %%v==32 set SW_VERSION=2024
            if %%v==31 set SW_VERSION=2023
            if %%v==30 set SW_VERSION=2022
            if %%v==29 set SW_VERSION=2021
            if %%v==28 set SW_VERSION=2020
            if %%v==27 set SW_VERSION=2019
        )
    )
)
if "!SW_PROGID!"=="" (
    echo [ERROR] SolidWorks not detected on this build machine.
    echo         The makepy COM cache cannot be generated.
    echo         Install SolidWorks or supply a pre-generated gen_py cache:
    echo         %PY_DIR%\Lib\site-packages\win32com\gen_py\
    echo         Or set CI=true to skip this check (CI/CD builds).
    echo.
    pause
    exit /b 1
)
echo [OK] SolidWorks !SW_VERSION! detected: !SW_PROGID!
echo.

:sw_done

REM ── Create dist directories ──────────────────────────────────────────────────
if not exist "%PY_DIR%"      mkdir "%PY_DIR%"
if not exist "%AGENT_DIST%"  mkdir "%AGENT_DIST%"
if not exist "%OUT_DIR%"     mkdir "%OUT_DIR%"

REM ── Download Python embeddable ───────────────────────────────────────────────
set PY_ZIP_PATH=%TEMP%\%PY_ZIP%
if not exist "%PY_DIR%\python.exe" (
    if not exist "%PY_ZIP_PATH%" (
        echo [STEP] Downloading Python %PY_VERSION% embeddable...
        powershell -Command "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; (New-Object Net.WebClient).DownloadFile('%PY_URL%', '%PY_ZIP_PATH%')"
        if errorlevel 1 (
            echo [ERROR] Download failed. Check internet connection.
            pause & exit /b 1
        )
    ) else (
        echo [OK] Using cached download: %PY_ZIP_PATH%
    )
    echo [STEP] Extracting Python embeddable...
    powershell -Command "Add-Type -A System.IO.Compression.FileSystem; [IO.Compression.ZipFile]::ExtractToDirectory('%PY_ZIP_PATH%', '%PY_DIR%')"
    if errorlevel 1 (echo [ERROR] Extraction failed. & pause & exit /b 1)
    echo [OK] Python extracted to %PY_DIR%
) else (
    echo [OK] Python already present at %PY_DIR%
)

set PY=%PY_DIR%\python.exe
set PIP=%PY_DIR%\Scripts\pip.exe

REM ── Enable site-packages in embeddable Python ─────────────────────────────────
echo [STEP] Enabling site-packages in embeddable Python...
for %%f in ("%PY_DIR%\python*._pth") do (
    powershell -Command "(gc '%%f') -replace '#import site','import site' | sc '%%f'"
)
echo [OK] site-packages enabled

REM ── Install pip ──────────────────────────────────────────────────────────────
if not exist "%PIP%" (
    echo [STEP] Installing pip...
    set GET_PIP=%TEMP%\get-pip.py
    if not exist "!GET_PIP!" (
        powershell -Command "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; (New-Object Net.WebClient).DownloadFile('%GET_PIP_URL%', '!GET_PIP!')"
    )
    "%PY%" "!GET_PIP!" --no-warn-script-location
    if errorlevel 1 (echo [ERROR] pip install failed & pause & exit /b 1)
    echo [OK] pip installed
) else (
    echo [OK] pip already present
)

REM ── Install packages ──────────────────────────────────────────────────────────
echo [STEP] Installing pywin32...
"%PY%" -m pip install "pywin32>=306" --no-warn-script-location
if errorlevel 1 (echo [ERROR] pywin32 install failed & pause & exit /b 1)
echo [OK] pywin32 installed

echo [STEP] Installing requests...
"%PY%" -m pip install "requests>=2.28.0" --no-warn-script-location
if errorlevel 1 (echo [ERROR] requests install failed & pause & exit /b 1)
echo [OK] requests installed

REM ── pywin32 post-install ──────────────────────────────────────────────────────
echo [STEP] Running pywin32 post-install (registers COM DLLs)...
if exist "%PY_DIR%\Scripts\pywin32_postinstall.py" (
    "%PY%" "%PY_DIR%\Scripts\pywin32_postinstall.py" -install
) else (
    for /r "%PY_DIR%" %%f in (pywin32_postinstall.py) do (
        "%PY%" "%%f" -install
        goto :post_done
    )
)
:post_done
echo [OK] pywin32 post-install done

REM ── Generate SolidWorks makepy cache (skip in CI) ────────────────────────────
if /i "!CI!"=="true" (
    echo [CI] Skipping makepy -- no SolidWorks in CI environment.
    echo [CI] The agent generates the COM cache at first startup on the target machine.
    goto :makepy_done
)

echo [STEP] Generating SolidWorks COM type library cache...
echo        (Required on machine with SolidWorks installed)
echo.

REM Method A: standard makepy via ProgID
echo [Makepy-A] python -m win32com.client.makepy "!SW_PROGID!"
"%PY%" -m win32com.client.makepy "!SW_PROGID!"
if not errorlevel 1 (
    echo [OK] Makepy cache generated via Method A
    goto :makepy_done
)
echo [WARN] Method A failed -- trying Method B (TLB file scan)...

REM Method B: find TLB from SolidWorks install directory
set SW_DIR=
for %%y in (2024 2023 2022 2021 2020 2019) do (
    if "!SW_DIR!"=="" (
        for %%b in ("SOFTWARE\SolidWorks\SolidWorks %%y\Setup" "SOFTWARE\WOW6432Node\SolidWorks\SolidWorks %%y\Setup") do (
            if "!SW_DIR!"=="" (
                for /f "tokens=2*" %%a in ('reg query "HKLM\%%~b" /v "SldWorks dir" 2^>nul') do (
                    set SW_DIR=%%b
                )
            )
        )
    )
)

if not "!SW_DIR!"=="" (
    echo [Makepy-B] SolidWorks dir: !SW_DIR!
    for %%t in ("!SW_DIR!\*.tlb" "!SW_DIR!\sldworks.exe") do (
        if exist "%%t" (
            echo [Makepy-B] Trying: %%t
            "%PY%" -m win32com.client.makepy "%%t"
            if not errorlevel 1 (
                echo [OK] Makepy cache generated from %%t
                goto :makepy_done
            )
        )
    )
)

echo [WARN] Method B failed -- trying Method C (inline pythoncom)...
"%PY%" -c "import winreg,pythoncom,win32com.client.gencache as g,sys; p='!SW_PROGID!'; k=winreg.OpenKey(winreg.HKEY_CLASSES_ROOT,p+r'\CLSID'); c=winreg.QueryValue(k,''); winreg.CloseKey(k); k=winreg.OpenKey(winreg.HKEY_CLASSES_ROOT,r'CLSID\{}\TypeLib'.format(c)); t=winreg.QueryValue(k,''); winreg.CloseKey(k); k=winreg.OpenKey(winreg.HKEY_CLASSES_ROOT,r'TypeLib\{}'.format(t)); vs=[winreg.EnumKey(k,i) for i in range(100) if True]; winreg.CloseKey(k); maj,mn=[int(x) for x in sorted(vs)[-1].split('.')[:2]]; g.EnsureModule(t,0,maj,mn); print('OK')" 2>&1
if not errorlevel 1 (
    echo [OK] Makepy cache generated via Method C
    goto :makepy_done
)

echo [WARN] All makepy methods failed.
echo        The agent will attempt makepy at runtime (3 fallback methods).
echo        If runtime makepy also fails:
echo          %PY% -m win32com.client.makepy "!SW_PROGID!"
echo.

:makepy_done
echo.

REM ── Copy agent sources ─────────────────────────────────────────────────────────
echo [STEP] Copying agent sources to %AGENT_DIST%...
set SRC=%PKG_ROOT%

REM Structuring agent copies three Python folders:
REM   agent\       main_structurer.py, structure_job_client.py, structure_job_runner.py, ...
REM   extractor\   sw_instance.py  (shared COM helpers)
REM   structurer\  solidworks_structurer.py  (Phase 1 WRITE logic)

if exist "%AGENT_DIST%\agent"      rmdir /s /q "%AGENT_DIST%\agent"
if exist "%AGENT_DIST%\extractor"  rmdir /s /q "%AGENT_DIST%\extractor"
if exist "%AGENT_DIST%\structurer" rmdir /s /q "%AGENT_DIST%\structurer"

xcopy "%SRC%\agent"      "%AGENT_DIST%\agent\"      /e /i /q /y
xcopy "%SRC%\extractor"  "%AGENT_DIST%\extractor\"  /e /i /q /y
xcopy "%SRC%\structurer" "%AGENT_DIST%\structurer\"  /e /i /q /y

REM Remove __pycache__
for /d /r "%AGENT_DIST%" %%d in (__pycache__) do (
    if exist "%%d" rmdir /s /q "%%d"
)
echo [OK] Agent sources copied (agent\ extractor\ structurer\)

REM ── Write run.bat ──────────────────────────────────────────────────────────────
echo [STEP] Writing run.bat...
(
    echo @echo off
    echo title ThermopacStructuringAgent v%STRUCTURER_VERSION%
    echo set PYTHONUTF8=1
    echo set PYTHONIOENCODING=utf-8
    echo echo.
    echo echo  ThermopacStructuringAgent ^| SolidWorks Drawing Structuring Agent
    echo echo  THERMOPAC ERP Integration  ^|  Phase 1  ^|  v%STRUCTURER_VERSION%
    echo echo.
    echo "%PY_DIR%\python.exe" "%AGENT_DIST%\agent\main_structurer.py" --config "%AGENT_DIST%\config.ini" %%*
    echo pause
) > "%AGENT_DIST%\run.bat"
echo [OK] run.bat written

REM ── Write run-service.bat ──────────────────────────────────────────────────────
(
    echo @echo off
    echo set PYTHONUTF8=1
    echo set PYTHONIOENCODING=utf-8
    echo start "" /B "%PY_DIR%\python.exe" "%AGENT_DIST%\agent\main_structurer.py" --config "%AGENT_DIST%\config.ini"
) > "%AGENT_DIST%\run-service.bat"
echo [OK] run-service.bat written

REM ── Write makepy-repair.bat ────────────────────────────────────────────────────
(
    echo @echo off
    echo echo ThermopacStructuringAgent -- SolidWorks COM Cache Repair
    echo echo.
    echo "%PY_DIR%\python.exe" -m win32com.client.makepy "!SW_PROGID!"
    echo if errorlevel 1 ^(
    echo     echo Method A failed. Retrying via gencache...
    echo     "%PY_DIR%\python.exe" -c "import win32com.client.gencache as g; g.EnsureDispatch('!SW_PROGID!')"
    echo ^)
    echo echo.
    echo echo Done. Restart ThermopacStructuringAgent.
    echo pause
) > "%AGENT_DIST%\makepy-repair.bat"
echo [OK] makepy-repair.bat written

REM ── Write config.ini template ──────────────────────────────────────────────────
if not exist "%AGENT_DIST%\config.ini" (
    echo [STEP] Writing config.ini template...
    (
        echo ; ThermopacStructuringAgent configuration
        echo ; Generated by build-installer.bat v%STRUCTURER_VERSION%
        echo ; Edit before first run -- see INSTALL.md
        echo.
        echo [cloud]
        echo api_url    = https://thermopac-communication-thermopacllp.replit.app
        echo node_id    = %COMPUTERNAME%
        echo node_token = REPLACE_WITH_YOUR_TOKEN
        echo.
        echo [agent]
        echo ; testing ^| production
        echo mode              = production
        echo poll_interval_sec = 10
        echo job_timeout_sec   = 600
        echo max_retries       = 3
        echo.
        echo [paths]
        echo temp_dir = C:\ThermopacAgent\temp
        echo log_dir  = C:\ThermopacAgent\logs
        echo.
        echo [solidworks]
        echo solidworks_version = !SW_VERSION!
        echo visible            = false
        echo model_search_path  =
        echo.
        echo [structurer]
        echo ; REQUIRED: absolute path to the approved SolidWorks drawing template
        echo template_path =
        echo ; Root folder where structured drawings are staged
        echo staging_root  = C:\ThermopacStaging\drawings
    ) > "%AGENT_DIST%\config.ini"
    echo [OK] config.ini template written
) else (
    echo [OK] config.ini already exists -- preserving existing file
)

REM ── Copy gen_py cache into dist\python (if it was generated above) ─────────────
set GEN_PY_SRC=%APPDATA%\Python\Python311\site-packages\win32com\gen_py
set GEN_PY_ALT=%PY_DIR%\Lib\site-packages\win32com\gen_py
if exist "%GEN_PY_SRC%" (
    if not exist "%GEN_PY_ALT%" mkdir "%GEN_PY_ALT%"
    xcopy "%GEN_PY_SRC%" "%GEN_PY_ALT%\" /e /i /q /y >nul
    echo [OK] gen_py cache bundled from %GEN_PY_SRC%
) else if exist "%GEN_PY_ALT%" (
    echo [OK] gen_py cache already in dist\python
) else (
    echo [WARN] gen_py cache not found -- agent will generate at runtime
)

REM ── Check for Inno Setup (report only -- do NOT compile .exe here) ─────────────
echo.
echo [STEP] Checking for Inno Setup 6...
set ISCC=
for %%p in (
    "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
    "C:\Program Files\Inno Setup 6\ISCC.exe"
) do (
    if exist %%p set ISCC=%%p
)

if not "!ISCC!"=="" (
    echo [OK] Inno Setup found: !ISCC!

    if /i "!CI!"=="true" (
        REM ── CI: compile the .exe immediately ────────────────────────────────────
        REM Note: !ISCC! already contains quotes around the path from the FOR loop
        echo [CI] Compiling installer EXE...
        !ISCC! /Q "%~dp0setup.iss"
        if errorlevel 1 (
            echo [ERROR] ISCC compile failed.
            exit /b 1
        )
        echo [CI] ISCC compile complete.
        echo.
        echo [CI] Contents of installer_output\ ^(post-compile debug^):
        powershell -Command "Get-ChildItem installer_output | Format-Table Name,Length,LastWriteTime -AutoSize"
        echo.
        echo [CI] Expected: ThermopacStructuringAgent-Setup-v%STRUCTURER_VERSION%.exe
    ) else (
        echo.
        echo      To compile the installer .exe, run:
        echo        !ISCC! "%~dp0setup.iss"
        echo      Output: installer_output\ThermopacStructuringAgent-Setup-v%STRUCTURER_VERSION%.exe
    )
) else (
    echo [WARN] Inno Setup 6 not found.
    echo        Download from: https://jrsoftware.org/isinfo.php
    echo        Then compile with: iscc installer\setup.iss
    if /i "!CI!"=="true" (
        echo [CI] Cannot compile without Inno Setup -- ensure CI runner has Inno Setup installed.
        exit /b 1
    )
)

echo.
echo ============================================================
echo  BUILD COMPLETE -- ThermopacStructuringAgent v%STRUCTURER_VERSION%
echo ============================================================
echo.
echo  Agent files  : %AGENT_DIST%\
echo  Python       : %PY_DIR%\
echo  Installer src: %~dp0setup.iss
echo  Output dir   : installer_output\  (at repo root)
echo.
echo  To compile the .exe installer (Inno Setup required):
echo    iscc installer\setup.iss
echo.
echo  To distribute WITHOUT an .exe (source ZIP):
echo    ZIP contents of: %DIST_DIR%\  plus  installer\setup.ps1
echo    Users run: powershell -ExecutionPolicy Bypass -File setup.ps1
echo.
if /i "!CI!"=="true" (
    echo [CI] Exiting for CI runner.
    exit /b 0
)
pause
