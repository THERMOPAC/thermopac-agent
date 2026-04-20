ThermopacAgent v1.0.19 — Windows Installer Package
=====================================================

WHAT IS INCLUDED
----------------
setup.ps1            — Self-contained PowerShell installer (recommended)
setup.iss            — Inno Setup 6 script (for building a .exe installer)
build-installer.bat  — Full build pipeline (creates dist\ + optional .exe)
README-installer.txt — This file


QUICKEST PATH: PowerShell Installer
------------------------------------
Requirements on target PC:
  - Windows 10 or later (64-bit)
  - SolidWorks 2019–2024 installed
  - Internet access (downloads Python 3.11 and pywin32 automatically)
  - Administrator rights

Steps:
  1. Right-click setup.ps1 > "Run with PowerShell" as Administrator
     OR from an elevated PowerShell prompt:
     powershell -ExecutionPolicy Bypass -File setup.ps1

  2. The installer will:
     - Detect SolidWorks version automatically
     - Download Python 3.11 embeddable (~10 MB)
     - Install pywin32 + requests
     - Generate SolidWorks COM type library cache (makepy)
     - Create all runtime folders
     - Create Start Menu shortcuts + run scripts
     - Optionally create a scheduled task for auto-start

  3. Edit config.ini (shortcut in Start Menu: "Edit Config")
     Set api_url, node_id, and node_token per your environment.

  4. Run from Start Menu > ThermopacAgent
     Or: C:\Program Files\ThermopacAgent\run.bat


BUILDING A .EXE INSTALLER
--------------------------
To create a standalone ThermopacAgent-Setup-v1.0.19.exe:

  On the BUILD machine (must have SolidWorks installed):

  1. Run: installer\build-installer.bat
     This downloads Python, installs all packages, generates the
     makepy cache, and creates the dist\ tree.

  2. Install Inno Setup 6: https://jrsoftware.org/isinfo.php

  3. Compile: iscc installer\setup.iss
     Output: installer_output\ThermopacAgent-Setup-v1.0.19.exe

  The resulting .exe is a fully self-contained single-file installer
  that bundles Python + all dependencies + pre-generated makepy cache.
  No internet required on the target machine.


WHAT MUST EXIST EXTERNALLY ON THE TARGET PC
--------------------------------------------
  - SolidWorks 2019–2024 (any edition — Standard, Professional, Premium)
    The agent opens and reads drawing files via SolidWorks COM automation.
    SolidWorks must be installed; it does NOT need to be running at agent start.

  - Windows 10 or later (64-bit)

  - Administrator rights for the installer only.
    The agent itself runs as a normal user after installation.

  Nothing else: Python, pywin32, makepy — all handled by the installer.


TROUBLESHOOTING
---------------
If the agent starts but shows "makepy" COM errors:

  1. Run: Start Menu > ThermopacAgent > "Repair COM Cache"
     This runs: python -m win32com.client.makepy SldWorks.Application.XX

  2. If that fails, open the SolidWorks API SDK browser:
     SolidWorks > Tools > API SDK
     This confirms the COM server is registered correctly.

  3. Check SolidWorks version detection:
     Open config.ini and verify solidworks_version matches your SW version.

If the agent cannot connect to the cloud:
  - Verify api_url in config.ini
  - Verify node_token (in production mode, must be admin-issued)
  - The cloud URL: https://thermopac-communication-thermopacllp.replit.app


AFTER INSTALL SUMMARY
----------------------
Install dir : C:\Program Files\ThermopacAgent\
Data dir    : C:\ThermopacAgent\  (logs, temp)
Config      : C:\Program Files\ThermopacAgent\config.ini
Python      : C:\Program Files\ThermopacAgent\python\python.exe
Run         : Start Menu > ThermopacAgent
              OR C:\Program Files\ThermopacAgent\run.bat
Repair COM  : Start Menu > ThermopacAgent > Repair COM Cache
