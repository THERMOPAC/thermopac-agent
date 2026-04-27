@echo off
REM ============================================================
REM  set_prod_url.bat — Revert api_url to the PRODUCTION server.
REM  No administrator rights required.
REM ============================================================

setlocal

set "APPDATA_CFG=%APPDATA%\ThermopacStructuringAgent\config.ini"

echo.
echo  Removing api_url override from APPDATA (restoring production URL)...
echo.

if not exist "%APPDATA_CFG%" (
    echo  No APPDATA config found — nothing to revert.
    echo.
    pause
    exit /b 0
)

PowerShell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$cfg = '%APPDATA_CFG%';" ^
  "$lines = [System.IO.File]::ReadAllLines($cfg);" ^
  "$out = $lines | Where-Object { $_ -notmatch '^api_url\s*=' };" ^
  "[System.IO.File]::WriteAllLines($cfg, $out, [System.Text.UTF8Encoding]::new($false))"

echo  Done.
echo.
echo  api_url override removed. Agent will use the URL from:
echo  C:\Program Files\ThermopacStructuringAgent\config.ini
echo.
echo  Restart the Structuring Agent to apply.
echo.
pause
