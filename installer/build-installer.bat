@echo off
REM ThermopacAgent — Installer Build Pipeline
REM ─────────────────────────────────────────────────────────────────────────
REM Run from the local-agent\ directory:  installer\build-installer.bat
REM
REM What this does:
REM   1. Downloads Python 3.11.9 embeddable into dist\python\
REM   2. Enables site-packages in the embeddable Python
REM   3. Installs pip into the embeddable Python
REM   4. pip install pywin32 + requests into the embeddable Python
REM   5. Runs pywin32 post-install to register COM DLLs
REM   6. Generates SolidWorks COM type library cache (makepy) — three methods
REM   7. Copies agent sources into dist\ThermopacAgent\
REM   8. Copies run.bat / run-service.bat / makepy-repair.bat
REM   9. (Optional) Compiles setup.iss into ThermopacAgent-Setup-v%AGENT_VERSION%.exe
REM         if Inno Setup 6 is detected on the build machine
REM
REM Prerequisites (build machine):
REM   - Windows 10+ x64
REM   - SolidWorks installed (for makepy cache generation)
REM   - Internet access (to download Python embeddable)
REM   - Inno Setup 6 (optional — for .exe installer compilation)
REM ─────────────────────────────────────────────────────────────────────────

setlocal enabledelayedexpansion

set AGENT_VERSION=1.0.51
set PY_VERSION=3.11.9
set PY_ZIP=python-%PY_VERSION%-embed-amd64.zip
set PY_URL=https://www.python.org/ftp/python/%PY_VERSION%/%PY_ZIP%
set GET_PIP_URL=https://bootstrap.pypa.io/get-pip.py

REM Directories (relative to local-agent\)
set DIST_DIR=%~dp0..\dist
set PY_DIR=%DIST_DIR%\python
set AGENT_DIST=%DIST_DIR%\ThermopacAgent
set OUT_DIR=%~dp0..\installer_output

echo ============================================================
echo  ThermopacAgent v%AGENT_VERSION% — Installer Build
echo ============================================================
echo.

REM ── Detect SolidWorks ────────────────────────────────────────────────────
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
if "%SW_PROGID%"=="" (
    echo [ERROR] SolidWorks not detected on this build machine.
    echo         The makepy cache cannot be generated.
    echo         Install SolidWorks or copy a pre-generated gen_py cache into:
    echo         %PY_DIR%\Lib\site-packages\win32com\gen_py\
    echo.
    pause
    exit /b 1
)
echo [OK] SolidWorks %SW_VERSION% detected: %SW_PROGID%
echo.

REM ── Create dist directories ────────────────────────────────────────────────
if not exist "%PY_DIR%"         mkdir "%PY_DIR%"
if not exist "%AGENT_DIST%"     mkdir "%AGENT_DIST%"
if not exist "%OUT_DIR%"        mkdir "%OUT_DIR%"

REM ── Download Python embeddable ─────────────────────────────────────────────
set PY_ZIP_PATH=%TEMP%\%PY_ZIP%
if not exist "%PY_DIR%\python.exe" (
    if not exist "%PY_ZIP_PATH%" (
        echo [STEP] Downloading Python %PY_VERSION% embeddable...
        powershell -Command "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; (New-Object Net.WebClient).DownloadFile('%PY_URL%', '%PY_ZIP_PATH%')"
        if errorlevel 1 (
            echo [ERROR] Download failed. Check internet connection.
            pause & exit /b 1
        )
    )
    echo [STEP] Extracting Python embeddable...
    powershell -Command "Add-Type -A System.IO.Compression.FileSystem; [IO.Compression.ZipFile]::ExtractToDirectory('%PY_ZIP_PATH%', '%PY_DIR%')"
    if errorlevel 1 (echo [ERROR] Extraction failed. & pause & exit /b 1)
    echo [OK] Python extracted to %PY_DIR%
) else (
    echo [OK] Python already present
)

set PY=%PY_DIR%\python.exe
set PIP=%PY_DIR%\Scripts\pip.exe

REM ── Enable site-packages ────────────────────────────────────────────────────
echo [STEP] Enabling site-packages in embeddable Python...
for %%f in ("%PY_DIR%\python*._pth") do (
    powershell -Command "(gc '%%f') -replace '#import site','import site' | sc '%%f'"
)
echo [OK] site-packages enabled

REM ── Install pip ─────────────────────────────────────────────────────────────
if not exist "%PIP%" (
    echo [STEP] Installing pip...
    set GET_PIP=%TEMP%\get-pip.py
    if not exist "%GET_PIP%" (
        powershell -Command "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; (New-Object Net.WebClient).DownloadFile('%GET_PIP_URL%', '%GET_PIP%')"
    )
    "%PY%" "%GET_PIP%" --no-warn-script-location
    if errorlevel 1 (echo [ERROR] pip install failed & pause & exit /b 1)
    echo [OK] pip installed
) else (
    echo [OK] pip already present
)

REM ── Install packages ──────────────────────────────────────────────────────
echo [STEP] Installing pywin32...
"%PY%" -m pip install "pywin32>=306" --no-warn-script-location
if errorlevel 1 (echo [ERROR] pywin32 install failed & pause & exit /b 1)
echo [OK] pywin32 installed

echo [STEP] Installing requests...
"%PY%" -m pip install "requests>=2.31.0" --no-warn-script-location
if errorlevel 1 (echo [ERROR] requests install failed & pause & exit /b 1)
echo [OK] requests installed

REM ── pywin32 post-install ──────────────────────────────────────────────────
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

REM ── Generate SolidWorks makepy cache ─────────────────────────────────────
echo [STEP] Generating SolidWorks COM type library cache...
echo        (Must be done on a machine with SolidWorks installed)
echo.

REM Method A: standard makepy
echo [Makepy-A] python -m win32com.client.makepy "%SW_PROGID%"
"%PY%" -m win32com.client.makepy "%SW_PROGID%"
if not errorlevel 1 (
    echo [OK] Makepy cache generated via Method A
    goto :makepy_done
)
echo [WARN] Method A failed — trying Method B (TLB file scan)...

REM Method B: find and load TLB from SolidWorks install dir
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

echo [WARN] Method B failed — trying Method C (inline pythoncom)...

REM Method C: inline Python via pythoncom.LoadTypeLib
"%PY%" -c "import winreg,pythoncom,win32com.client.gencache as g,sys; p='%SW_PROGID%'; k=winreg.OpenKey(winreg.HKEY_CLASSES_ROOT,p+r'\CLSID'); c=winreg.QueryValue(k,''); winreg.CloseKey(k); k=winreg.OpenKey(winreg.HKEY_CLASSES_ROOT,r'CLSID\{}\TypeLib'.format(c)); t=winreg.QueryValue(k,''); winreg.CloseKey(k); k=winreg.OpenKey(winreg.HKEY_CLASSES_ROOT,r'TypeLib\{}'.format(t)); vs=[winreg.EnumKey(k,i) for i in range(100) if True]; winreg.CloseKey(k); maj,min_=[int(x) for x in sorted(vs)[-1].split('.')[:2]]; g.EnsureModule(t,0,maj,min_); print('OK')" 2>&1
if not errorlevel 1 (
    echo [OK] Makepy cache generated via Method C
    goto :makepy_done
)

echo [WARN] All makepy methods failed.
echo        The agent will attempt makepy at runtime using 3 fallback methods.
echo        If runtime makepy also fails, run: %PY% -m win32com.client.makepy "%SW_PROGID%"
echo.
goto :makepy_done

:makepy_done
echo.

REM ── Copy agent sources ────────────────────────────────────────────────────
echo [STEP] Copying agent sources to %AGENT_DIST%...
set SRC=%~dp0..

REM Agent Python source folders
if exist "%AGENT_DIST%\agent"     rmdir /s /q "%AGENT_DIST%\agent"
if exist "%AGENT_DIST%\extractor" rmdir /s /q "%AGENT_DIST%\extractor"

xcopy "%SRC%\agent"     "%AGENT_DIST%\agent\"     /e /i /q /y
xcopy "%SRC%\extractor" "%AGENT_DIST%\extractor\" /e /i /q /y

if exist "%SRC%\_com_helper.py" copy /y "%SRC%\_com_helper.py" "%AGENT_DIST%\"

REM Remove __pycache__
for /d /r "%AGENT_DIST%" %%d in (__pycache__) do (
    if exist "%%d" rmdir /s /q "%%d"
)

echo [OK] Agent sources copied

REM ── Write config.ini template ─────────────────────────────────────────────
if not exist "%AGENT_DIST%\config.ini" (
    echo [STEP] Writing config.ini template...
    (
        echo ; ThermopacAgent configuration
        echo ; Edit before first run
        echo.
        echo [cloud]
        echo api_url    = https://5d05ae61-8225-4651-bb76-b4e20a4ddabb-00-3mex6zlihlmft.janeway.replit.dev
        echo node_id    = %COMPUTERNAME%
        echo node_token = REPLACE_WITH_YOUR_TOKEN
        echo.
        echo [agent]
        echo mode = testing
        echo poll_interval_sec = 10
        echo job_timeout_sec   = 600
        echo max_retries       = 3
        echo.
        echo [paths]
        echo temp_dir = C:\ThermopacAgent\temp
        echo log_dir  = C:\ThermopacAgent\logs
        echo.
        echo [solidworks]
        echo solidworks_version = %SW_VERSION%
        echo visible = false
        echo model_search_path =
    ) > "%AGENT_DIST%\config.ini"
    echo [OK] config.ini template written
)

REM ── Write run.bat ─────────────────────────────────────────────────────────
echo [STEP] Writing run scripts...
(
    echo @echo off
    echo title ThermopacAgent v%AGENT_VERSION%
    echo echo.
    echo echo  ThermopacAgent ^| SolidWorks Extraction Agent
    echo echo  THERMOPAC ERP Integration
    echo echo.
    echo "%PY_DIR%\python.exe" "%AGENT_DIST%\agent\main.py" --config "%AGENT_DIST%\config.ini" %%*
    echo pause
) > "%AGENT_DIST%\run.bat"

(
    echo @echo off
    echo "%PY_DIR%\python.exe" "%AGENT_DIST%\agent\main.py" --config "%AGENT_DIST%\config.ini"
) > "%AGENT_DIST%\run-service.bat"

(
    echo @echo off
    echo echo ThermopacAgent — SolidWorks COM Cache Repair
    echo echo.
    echo "%PY_DIR%\python.exe" -m win32com.client.makepy "%SW_PROGID%"
    echo if errorlevel 1 ^(
    echo     echo Method A failed. Retrying...
    echo     "%PY_DIR%\python.exe" -c "import win32com.client.gencache as g; g.EnsureDispatch('%SW_PROGID%')"
    echo ^)
    echo echo.
    echo echo Done. Restart ThermopacAgent.
    echo pause
) > "%AGENT_DIST%\makepy-repair.bat"

echo [OK] run.bat / run-service.bat / makepy-repair.bat written

REM ── (Optional) Compile InnoSetup installer ────────────────────────────────
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
    echo [STEP] Compiling installer...
    !ISCC! "%~dp0setup.iss"
    if not errorlevel 1 (
        echo [OK] Installer compiled: %OUT_DIR%\ThermopacAgent-Setup-v%AGENT_VERSION%.exe
    ) else (
        echo [WARN] Inno Setup compilation failed (non-fatal)
    )
) else (
    echo [WARN] Inno Setup 6 not found — skipping .exe compilation
    echo        To compile: install https://jrsoftware.org/isinfo.php
    echo        Then run: iscc installer\setup.iss
)

echo.
echo ============================================================
echo  BUILD COMPLETE
echo ============================================================
echo  Agent files : %AGENT_DIST%\
echo  Python      : %PY_DIR%\
echo  Installer   : %OUT_DIR%\  (if Inno Setup was found)
echo.
echo  To distribute WITHOUT compiling an .exe:
echo    ZIP the contents of: %DIST_DIR%\
echo    plus: installer\setup.ps1
echo    Users run: powershell -ExecutionPolicy Bypass -File setup.ps1
echo.
pause
