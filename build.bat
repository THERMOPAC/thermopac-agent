@echo off
REM ThermopacAgent — PyInstaller build script
REM Run from the local-agent\ directory: build.bat

echo ============================================================
echo  ThermopacAgent — Build
echo ============================================================

REM Verify Python and PyInstaller are available
python --version >nul 2>&1 || (echo ERROR: Python not found && exit /b 1)
pyinstaller --version >nul 2>&1 || (echo ERROR: PyInstaller not found. Run: pip install pyinstaller && exit /b 1)

REM Clean previous build
if exist dist\ThermopacAgent rmdir /s /q dist\ThermopacAgent
if exist build rmdir /s /q build

echo Building EXE...

pyinstaller ^
  --name ThermopacAgent ^
  --onedir ^
  --console ^
  --clean ^
  --noconfirm ^
  --add-data "config.ini;." ^
  --add-data "extractor;extractor" ^
  --add-data "agent;agent" ^
  --hidden-import win32com ^
  --hidden-import win32com.client ^
  --hidden-import pythoncom ^
  --hidden-import pywintypes ^
  --hidden-import requests ^
  --hidden-import configparser ^
  --paths agent ^
  --paths extractor ^
  agent\main.py

if errorlevel 1 (
  echo ERROR: Build failed
  exit /b 1
)

REM Copy config.ini to dist if not already there
if not exist dist\ThermopacAgent\config.ini (
  copy config.ini dist\ThermopacAgent\config.ini
)

echo.
echo ============================================================
echo  Build complete: dist\ThermopacAgent\ThermopacAgent.exe
echo  Edit dist\ThermopacAgent\config.ini before distributing
echo ============================================================
