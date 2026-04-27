@echo off
REM ============================================================
REM  Thermopac Structuring Agent — Set Node Token
REM  Writes node_token to user APPDATA (no admin needed).
REM  Run this ONCE after registering your node in the ERP.
REM ============================================================
title Set Node Token

set "CFG_DIR=%APPDATA%\ThermopacStructuringAgent"
set "CFG=%CFG_DIR%\config.ini"

echo.
echo  ============================================================
echo   Thermopac Structuring Agent -- Set Node Token
echo  ============================================================
echo.
echo  Paste the token issued by EPC ^> Drawing Controls ^> Agent Nodes
echo  (the token is shown ONCE only)
echo.
set /p TOKEN=  Token: 

if "%TOKEN%"=="" (
    echo.
    echo  ERROR: No token entered. Exiting.
    pause
    exit /b 1
)

set /p NODE_ID=  Node ID (press Enter to use default 'runnervmxu3fp'): 
if "%NODE_ID%"=="" set "NODE_ID=runnervmxu3fp"

mkdir "%CFG_DIR%" 2>nul

powershell -NoProfile -Command ^
  "$f='%CFG%'; $t='%TOKEN%'; $n='%NODE_ID%'; $c = if (Test-Path $f) { Get-Content $f } else { @() }; $c = $c | Where-Object { $_ -notmatch '^\[cloud\]' -and $_ -notmatch '^\s*node_token\s*=' -and $_ -notmatch '^\s*node_id\s*=' }; $c = @('[cloud]', \"node_id    = $n\", \"node_token = $t\") + $c; Set-Content $f $c -Encoding UTF8"

echo.
echo  Done!  Token saved to:
echo  %CFG%
echo.
echo  Now start the agent using the desktop shortcut or Start Menu.
echo.
pause
