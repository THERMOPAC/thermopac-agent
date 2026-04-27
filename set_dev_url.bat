@echo off
REM ============================================================
REM  set_dev_url.bat
REM  Points the Structuring Agent at the Replit DEV server.
REM  No administrator rights required.
REM  Writes api_url to: %APPDATA%\ThermopacStructuringAgent\config.ini
REM ============================================================
setlocal

set "APPDATA_DIR=%APPDATA%\ThermopacStructuringAgent"
set "APPDATA_CFG=%APPDATA_DIR%\config.ini"
set "DEV_URL=https://5d05ae61-8225-4651-bb76-b4e20a4ddabb-00-3mex6zlihlmft.janeway.replit.dev"
set "PS1=%TEMP%\thermopac_set_dev.ps1"

if not exist "%APPDATA_DIR%" mkdir "%APPDATA_DIR%"

REM Write the PowerShell script to a temp file (avoids batch ^ quoting hell)
> "%PS1%" echo $cfg = $env:APPDATA + '\ThermopacStructuringAgent\config.ini'
>> "%PS1%" echo $url = 'https://5d05ae61-8225-4651-bb76-b4e20a4ddabb-00-3mex6zlihlmft.janeway.replit.dev'
>> "%PS1%" echo if (-not (Test-Path $cfg)) { New-Item -ItemType File -Path $cfg -Force | Out-Null }
>> "%PS1%" echo $txt = [IO.File]::ReadAllText($cfg)
>> "%PS1%" echo $newLine = 'api_url = ' + $url
>> "%PS1%" echo if ($txt -match '(?m)^api_url\s*=') {
>> "%PS1%" echo     $txt = [regex]::Replace($txt, '(?m)^api_url\s*=.*', $newLine)
>> "%PS1%" echo } elseif ($txt -match '(?m)^\[cloud\]') {
>> "%PS1%" echo     $txt = [regex]::Replace($txt, '(?m)^\[cloud\]', "[cloud]`r`n" + $newLine)
>> "%PS1%" echo } else {
>> "%PS1%" echo     $txt = $txt.TrimEnd() + "`r`n`r`n[cloud]`r`n" + $newLine + "`r`n"
>> "%PS1%" echo }
>> "%PS1%" echo [IO.File]::WriteAllText($cfg, $txt, [Text.UTF8Encoding]::new($false))
>> "%PS1%" echo Write-Host " Done. api_url set to: $url"

echo.
echo  Writing dev URL to APPDATA config...
PowerShell -NoProfile -ExecutionPolicy Bypass -File "%PS1%"
del /q "%PS1%" 2>nul

echo.
echo  Config: %APPDATA_CFG%
echo.
echo  *** Close this window and RESTART the Structuring Agent ***
echo.
pause
