@echo off
REM ============================================================
REM  Thermopac Structuring Agent — Set Testing Mode
REM  Writes mode = testing to user APPDATA (no admin needed).
REM  Run this ONCE, then start the agent normally.
REM ============================================================
title Set Testing Mode

set "CFG_DIR=%APPDATA%\ThermopacStructuringAgent"
set "CFG=%CFG_DIR%\config.ini"

mkdir "%CFG_DIR%" 2>nul

powershell -NoProfile -Command ^
  "$f='%CFG%'; $enc=[System.Text.UTF8Encoding]::new($false); $c = if (Test-Path $f) { [System.IO.File]::ReadAllText($f, $enc) } else { '' }; $lines = $c -split \"`n\" | Where-Object { $_ -notmatch '^\s*mode\s*=' }; if (-not ($lines -match '^\[agent\]')) { $lines += '[agent]' }; $idx = [array]::IndexOf($lines, ($lines | Where-Object { $_ -match '^\[agent\]' } | Select-Object -First 1)); $out = ($lines[0..$idx] + 'mode = testing' + $lines[($idx+1)..($lines.Length-1)]) -join \"`n\"; [System.IO.File]::WriteAllText($f, $out, $enc)"

echo.
echo  Done!  mode = testing written to:
echo  %CFG%
echo.
echo  Now start the agent using run.bat in the install folder.
echo.
pause
