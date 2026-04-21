#Requires -RunAsAdministrator
<#
.SYNOPSIS
    ThermopacAgent Windows Installer
.DESCRIPTION
    Self-contained installer for the ThermopacAgent SolidWorks extraction service.
    Downloads Python 3.11 Embeddable, installs all dependencies offline-capable,
    generates win32com SolidWorks COM type library cache, and creates run scripts.
.PARAMETER InstallDir
    Installation directory (default: C:\Program Files\ThermopacAgent)
.PARAMETER DataDir
    Runtime data directory for logs and temp files (default: C:\ThermopacAgent)
.PARAMETER Silent
    Skip interactive prompts; use defaults for everything
.EXAMPLE
    powershell -ExecutionPolicy Bypass -File setup.ps1
    powershell -ExecutionPolicy Bypass -File setup.ps1 -InstallDir "D:\ThermopacAgent"
    powershell -ExecutionPolicy Bypass -File setup.ps1 -Silent
#>
param(
    [string]$InstallDir  = "C:\Program Files\ThermopacAgent",
    [string]$DataDir     = "C:\ThermopacAgent",
    [switch]$Silent
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ── Constants ────────────────────────────────────────────────────────────────
$AGENT_VERSION   = "1.0.34"
$PY_VERSION      = "3.11.9"
$PY_EMBED_ZIP    = "python-$PY_VERSION-embed-amd64.zip"
$PY_EMBED_URL    = "https://www.python.org/ftp/python/$PY_VERSION/$PY_EMBED_ZIP"
$GET_PIP_URL     = "https://bootstrap.pypa.io/get-pip.py"
$PY_PTH_FILE     = "python311._pth"
$REQUIRED_PKGS   = @("pywin32>=306", "requests>=2.31.0")
$APP_URL         = "https://5d05ae61-8225-4651-bb76-b4e20a4ddabb-00-3mex6zlihlmft.janeway.replit.dev"

$ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$SourceDir  = Split-Path -Parent $ScriptDir  # local-agent root

# ── Helpers ──────────────────────────────────────────────────────────────────

function Write-Header {
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "  ThermopacAgent $AGENT_VERSION — Windows Installer" -ForegroundColor Cyan
    Write-Host "  THERMOPAC ERP  |  SolidWorks Extraction Agent" -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host ""
}

function Write-Step([string]$msg) {
    Write-Host "[STEP] $msg" -ForegroundColor Yellow
}

function Write-OK([string]$msg) {
    Write-Host "  [OK] $msg" -ForegroundColor Green
}

function Write-Warn([string]$msg) {
    Write-Host "  [WARN] $msg" -ForegroundColor DarkYellow
}

function Write-Fail([string]$msg) {
    Write-Host "  [FAIL] $msg" -ForegroundColor Red
}

function Confirm-Proceed([string]$prompt) {
    if ($Silent) { return }
    $ans = Read-Host "$prompt [Y/n]"
    if ($ans -match '^[Nn]') {
        Write-Host "Aborted." -ForegroundColor Red
        exit 1
    }
}

function Download-File([string]$Url, [string]$Dest) {
    Write-Host "    Downloading: $Url" -ForegroundColor DarkGray
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $wc = New-Object System.Net.WebClient
    $wc.Headers.Add("User-Agent", "ThermopacAgent-Installer/$AGENT_VERSION")
    $wc.DownloadFile($Url, $Dest)
    Write-OK "Downloaded to: $Dest"
}

# ── Step 1: Validate Prerequisites ───────────────────────────────────────────
Write-Header

Write-Step "Checking prerequisites..."

# Windows version
$os = [System.Environment]::OSVersion.Version
if ($os.Major -lt 10) {
    Write-Fail "Windows 10 or later is required (found: $($os.Major).$($os.Minor))"
    exit 1
}
Write-OK "Windows $($os.Major).$($os.Minor) — OK"

# Check SolidWorks is installed
Write-Step "Detecting SolidWorks installation..."
$sw_progid  = $null
$sw_version = 0
$sw_dir     = $null

$sw_versions = @{
    2024 = "SldWorks.Application.32"
    2023 = "SldWorks.Application.31"
    2022 = "SldWorks.Application.30"
    2021 = "SldWorks.Application.29"
    2020 = "SldWorks.Application.28"
    2019 = "SldWorks.Application.27"
}

foreach ($year in ($sw_versions.Keys | Sort-Object -Descending)) {
    $progid = $sw_versions[$year]
    try {
        $key = [Microsoft.Win32.Registry]::ClassesRoot.OpenSubKey($progid)
        if ($key -ne $null) {
            $sw_version = $year
            $sw_progid  = $progid
            $key.Close()

            # Find install directory
            foreach ($base in @(
                "SOFTWARE\SolidWorks\SolidWorks $year\Setup",
                "SOFTWARE\WOW6432Node\SolidWorks\SolidWorks $year\Setup"
            )) {
                try {
                    $rk = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($base)
                    if ($rk -ne $null) {
                        $d = $rk.GetValue("SldWorks dir")
                        if ($d -and (Test-Path $d)) { $sw_dir = $d }
                        $rk.Close()
                    }
                } catch {}
            }
            break
        }
    } catch {}
}

if (-not $sw_progid) {
    Write-Fail "SolidWorks NOT detected.  Install SolidWorks before running this installer."
    Write-Host ""
    Write-Host "  Supported versions: SolidWorks 2019 – 2024" -ForegroundColor Yellow
    if (-not $Silent) { Read-Host "Press Enter to exit" }
    exit 1
}

Write-OK "SolidWorks $sw_version detected  ($sw_progid)"
if ($sw_dir) { Write-OK "SolidWorks directory: $sw_dir" }

# ── Step 2: Create directories ───────────────────────────────────────────────
Write-Step "Creating installation directories..."

foreach ($d in @($InstallDir, "$DataDir\logs", "$DataDir\temp", "$DataDir\config")) {
    if (-not (Test-Path $d)) {
        New-Item -ItemType Directory -Path $d -Force | Out-Null
    }
    Write-OK $d
}

# Set ACLs on data dir so any user can write
try {
    $acl  = Get-Acl $DataDir
    $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
        "Everyone", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
    $acl.SetAccessRule($rule)
    Set-Acl $DataDir $acl
    Write-OK "Permissions set on $DataDir"
} catch {
    Write-Warn "Could not set ACL on $DataDir (non-fatal): $_"
}

# ── Step 3: Python Embeddable ─────────────────────────────────────────────────
Write-Step "Setting up Python $PY_VERSION (embedded)..."

$PyDir   = "$InstallDir\python"
$PyExe   = "$PyDir\python.exe"
$ZipDest = "$env:TEMP\$PY_EMBED_ZIP"

if (-not (Test-Path $PyExe)) {
    if (-not (Test-Path $ZipDest)) {
        Write-Host "    Downloading Python embeddable package..."
        try {
            Download-File $PY_EMBED_URL $ZipDest
        } catch {
            Write-Fail "Download failed: $_"
            Write-Host "  Manual fix: Download $PY_EMBED_URL" -ForegroundColor Yellow
            Write-Host "  and place it at: $ZipDest" -ForegroundColor Yellow
            exit 1
        }
    } else {
        Write-OK "Using cached download: $ZipDest"
    }

    Write-Host "    Extracting Python embeddable..."
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipDest, $PyDir)
    Write-OK "Extracted to $PyDir"
} else {
    Write-OK "Python already present: $PyExe"
}

# ── Enable site-packages in embeddable Python ──────────────────────────────
$pthFile = Get-ChildItem $PyDir -Filter "python*._pth" | Select-Object -First 1
if ($pthFile) {
    $pthContent = Get-Content $pthFile.FullName -Raw
    if ($pthContent -match "#import site") {
        $pthContent = $pthContent -replace "#import site", "import site"
        Set-Content $pthFile.FullName $pthContent -NoNewline
        Write-OK "Enabled site-packages in $($pthFile.Name)"
    }
    # Also add Scripts to path
    if ($pthContent -notmatch "Scripts") {
        Add-Content $pthFile.FullName "`nScripts"
    }
} else {
    Write-Warn "._pth file not found — site-packages may not work"
}

# ── Step 4: pip ───────────────────────────────────────────────────────────────
Write-Step "Installing pip..."

$PipScript = "$env:TEMP\get-pip.py"
$PipExe    = "$PyDir\Scripts\pip.exe"

if (-not (Test-Path $PipExe)) {
    if (-not (Test-Path $PipScript)) {
        try {
            Download-File $GET_PIP_URL $PipScript
        } catch {
            Write-Fail "Could not download get-pip.py: $_"
            exit 1
        }
    }
    Write-Host "    Running get-pip.py..."
    & $PyExe $PipScript --no-warn-script-location 2>&1 | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
    if ($LASTEXITCODE -ne 0) {
        Write-Fail "pip installation failed (exit $LASTEXITCODE)"
        exit 1
    }
    Write-OK "pip installed"
} else {
    Write-OK "pip already present"
}

$PipExe = "$PyDir\Scripts\pip.exe"

# ── Step 5: Install Python packages ─────────────────────────────────────────
Write-Step "Installing Python packages..."

foreach ($pkg in $REQUIRED_PKGS) {
    Write-Host "    pip install $pkg ..."
    & $PyExe -m pip install $pkg --no-warn-script-location 2>&1 |
        ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
    if ($LASTEXITCODE -ne 0) {
        Write-Fail "Failed to install $pkg"
        exit 1
    }
    Write-OK $pkg
}

# ── Step 6: pywin32 post-install (registers COM DLLs) ────────────────────────
Write-Step "Running pywin32 post-install (registers COM support DLLs)..."

$pywin32postinstall = "$PyDir\Scripts\pywin32_postinstall.py"
if (-not (Test-Path $pywin32postinstall)) {
    # Try site-packages location
    $pywin32postinstall = "$PyDir\Lib\site-packages\pywin32_postinstall.py"
}
if (Test-Path $pywin32postinstall) {
    & $PyExe $pywin32postinstall -install 2>&1 |
        ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
    Write-OK "pywin32 post-install complete"
} else {
    Write-Warn "pywin32_postinstall.py not found — skipping (non-fatal)"
}

# ── Step 7: Generate SolidWorks COM type library cache (makepy) ───────────────
Write-Step "Generating SolidWorks COM type library cache (makepy)..."
Write-Host "    This enables IDrawingDoc methods (GetCurrentSheet, GetFirstView, etc.)"

$makepy_ok = $false

# Method A: via ProgID — standard makepy
Write-Host "    [Makepy-A] Trying: python -m win32com.client.makepy `"$sw_progid`"..."
$makepy_out = & $PyExe -m win32com.client.makepy $sw_progid 2>&1
if ($LASTEXITCODE -eq 0) {
    Write-OK "Makepy cache generated via ProgID"
    $makepy_ok = $true
} else {
    Write-Warn "Method A failed: $makepy_out"
}

# Method B: via TLB file found in SolidWorks install directory
if (-not $makepy_ok -and $sw_dir) {
    Write-Host "    [Makepy-B] Scanning SolidWorks directory for type library files..."
    $tlb_candidates = @()

    foreach ($ext in @("*.tlb", "sldworks.exe")) {
        $found = Get-ChildItem -Path $sw_dir -Filter $ext -ErrorAction SilentlyContinue |
            Select-Object -First 3
        $tlb_candidates += $found
    }
    # Also check common sub-folders
    foreach ($sub in @("api\redist", "api")) {
        $sub_path = Join-Path $sw_dir $sub
        if (Test-Path $sub_path) {
            Get-ChildItem -Path $sub_path -Filter "*.tlb" -ErrorAction SilentlyContinue |
                ForEach-Object { $tlb_candidates += $_ }
        }
    }

    foreach ($tlb in $tlb_candidates) {
        Write-Host "    [Makepy-B] Trying: $($tlb.FullName)..."
        $out = & $PyExe -m win32com.client.makepy "`"$($tlb.FullName)`"" 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-OK "Makepy cache generated from $($tlb.Name)"
            $makepy_ok = $true
            break
        }
    }
    if (-not $makepy_ok) {
        Write-Warn "Method B: no TLB produced cache"
    }
}

# Method C: inline Python script using pythoncom.LoadTypeLib directly
if (-not $makepy_ok) {
    Write-Host "    [Makepy-C] Attempting pythoncom.LoadTypeLib via registry walk..."
    $inline = @"
import sys, winreg, pythoncom, win32com.client.gencache as gc
progid = sys.argv[1]
try:
    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid + r'\CLSID') as k:
        clsid = winreg.QueryValue(k, '')
    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r'CLSID\{}\TypeLib'.format(clsid)) as k:
        tl_clsid = winreg.QueryValue(k, '')
    versions = []
    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r'TypeLib\{}'.format(tl_clsid)) as k:
        i = 0
        while True:
            try:
                vs = winreg.EnumKey(k, i); i += 1
                p = vs.split('.')
                versions.append((int(p[0]), int(p[1]) if len(p)>1 else 0, vs))
            except OSError:
                break
    if not versions:
        raise ValueError('no versions')
    major, minor, ver_str = sorted(versions)[-1]
    for arch in ('win64','win32',''):
        sub = r'TypeLib\{}\{}\0'.format(tl_clsid, ver_str)
        kp  = sub + '\\' + arch if arch else sub
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, kp) as k:
                fp = winreg.QueryValue(k,'').strip('"').split(',')[0].strip()
            import os
            if os.path.isfile(fp):
                tl  = pythoncom.LoadTypeLib(fp)
                tla = tl.GetLibAttr()
                gc.EnsureModule(str(tla[0]), 0, tla[3], tla[4],
                                bForDemand=False, bBuildHidden=True)
                print('OK:' + fp)
                sys.exit(0)
        except Exception as ex:
            pass
    raise RuntimeError('no TLB file loaded')
except Exception as ex:
    print('FAIL:' + str(ex), file=sys.stderr)
    sys.exit(1)
"@
    $inline | Set-Content "$env:TEMP\makepy_sw.py" -Encoding UTF8
    $out = & $PyExe "$env:TEMP\makepy_sw.py" $sw_progid 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-OK "Makepy cache generated via inline pythoncom.LoadTypeLib"
        $makepy_ok = $true
    } else {
        Write-Warn "Method C failed: $out"
    }
}

if (-not $makepy_ok) {
    Write-Warn "SolidWorks COM cache could not be pre-generated automatically."
    Write-Host ""
    Write-Host "  This is not a fatal installer error — the agent will attempt" -ForegroundColor Yellow
    Write-Host "  to generate the cache at runtime using three fallback methods." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  If the agent still fails with 'makepy' errors, run manually:" -ForegroundColor Yellow
    Write-Host "    $PyExe -m win32com.client.makepy `"$sw_progid`"" -ForegroundColor White
    Write-Host ""
}

# ── Step 8: Copy agent source files ──────────────────────────────────────────
Write-Step "Installing agent files..."

# Determine source — works whether run from installer\ subfolder or root
$agent_sources = @("agent", "extractor", "_com_helper.py")
foreach ($item in $agent_sources) {
    $src = Join-Path $SourceDir $item
    if (-not (Test-Path $src)) {
        Write-Warn "Source not found (non-fatal): $src"
        continue
    }
    $dest = Join-Path $InstallDir $item
    if (Test-Path $src -PathType Container) {
        if (Test-Path $dest) { Remove-Item $dest -Recurse -Force }
        Copy-Item $src $dest -Recurse -Force
    } else {
        Copy-Item $src $dest -Force
    }
    Write-OK "Installed: $item"
}

# ── Step 9: Write config.ini ──────────────────────────────────────────────────
Write-Step "Creating config.ini..."

$config_path = "$InstallDir\config.ini"
if (-not (Test-Path $config_path)) {
    $machine_name = $env:COMPUTERNAME
    $config_content = @"
; ThermopacAgent configuration
; Generated by installer v$AGENT_VERSION

[cloud]
api_url    = $APP_URL
node_id    = $machine_name
node_token = REPLACE_WITH_YOUR_TOKEN

[agent]
; testing | production
; testing  -- auto-generates token and self-registers with cloud
; production -- requires cloud/admin-issued token, no auto-registration
mode = testing

poll_interval_sec = 10
job_timeout_sec   = 600
max_retries       = 3

[paths]
temp_dir = $DataDir\temp
log_dir  = $DataDir\logs

[solidworks]
; Auto-detected during installation
solidworks_version = $sw_version
; solidworks_progid =
visible = false

; Semicolon-separated folders to search for referenced parts/assemblies
model_search_path =
"@
    Set-Content $config_path $config_content -Encoding UTF8
    Write-OK "Created: $config_path"
} else {
    Write-OK "Existing config.ini preserved: $config_path"
}

# ── Step 10: Create run scripts ───────────────────────────────────────────────
Write-Step "Creating run scripts..."

# run.bat
$run_bat = @"
@echo off
title ThermopacAgent v$AGENT_VERSION
:: UTF-8 output — prevents UnicodeEncodeError on Windows cp1252 consoles
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8
echo.
echo  ThermopacAgent — SolidWorks Extraction Agent
echo  THERMOPAC ERP Integration
echo.
"$PyExe" "$InstallDir\agent\main.py" --config "$config_path" %*
pause
"@
Set-Content "$InstallDir\run.bat" $run_bat -Encoding ASCII
Write-OK "Created: $InstallDir\run.bat"

# run-service.bat (no pause, for scheduled task use)
$run_service_bat = @"
@echo off
"$PyExe" "$InstallDir\agent\main.py" --config "$config_path"
"@
Set-Content "$InstallDir\run-service.bat" $run_service_bat -Encoding ASCII
Write-OK "Created: $InstallDir\run-service.bat"

# makepy-repair.bat — standalone repair tool
$makepy_bat = @"
@echo off
echo ThermopacAgent — SolidWorks COM Cache Repair
echo.
echo Regenerating win32com type library cache for: $sw_progid
"$PyExe" -m win32com.client.makepy "$sw_progid"
if errorlevel 1 (
    echo.
    echo Method A failed. Trying pythoncom approach...
    "$PyExe" "$env:TEMP\makepy_sw.py" "$sw_progid"
)
echo.
echo Done. Restart ThermopacAgent.
pause
"@
Set-Content "$InstallDir\makepy-repair.bat" $makepy_bat -Encoding ASCII
Write-OK "Created: $InstallDir\makepy-repair.bat"

# ── Step 11: Start Menu shortcut ─────────────────────────────────────────────
Write-Step "Creating Start Menu shortcut..."

try {
    $StartMenu = [System.Environment]::GetFolderPath("CommonPrograms")
    $ShortcutDir = "$StartMenu\ThermopacAgent"
    if (-not (Test-Path $ShortcutDir)) { New-Item -ItemType Directory -Path $ShortcutDir -Force | Out-Null }

    $WshShell = New-Object -ComObject WScript.Shell

    # Main shortcut
    $sc = $WshShell.CreateShortcut("$ShortcutDir\ThermopacAgent.lnk")
    $sc.TargetPath = "$InstallDir\run.bat"
    $sc.WorkingDirectory = $InstallDir
    $sc.Description = "ThermopacAgent — SolidWorks Extraction Agent"
    $sc.Save()

    # Config shortcut
    $sc2 = $WshShell.CreateShortcut("$ShortcutDir\Edit Config.lnk")
    $sc2.TargetPath = "notepad.exe"
    $sc2.Arguments  = "`"$config_path`""
    $sc2.Description = "Edit ThermopacAgent configuration"
    $sc2.Save()

    # Repair shortcut
    $sc3 = $WshShell.CreateShortcut("$ShortcutDir\Repair COM Cache.lnk")
    $sc3.TargetPath = "$InstallDir\makepy-repair.bat"
    $sc3.WorkingDirectory = $InstallDir
    $sc3.Description = "Regenerate SolidWorks COM type library cache"
    $sc3.Save()

    Write-OK "Start Menu: $ShortcutDir"
} catch {
    Write-Warn "Could not create Start Menu shortcuts: $_"
}

# ── Step 12: Optional scheduled task ─────────────────────────────────────────
if (-not $Silent) {
    Write-Host ""
    $ans = Read-Host "Create scheduled task to auto-start agent at Windows login? [y/N]"
    if ($ans -match '^[Yy]') {
        Write-Step "Creating scheduled task..."
        $taskXml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo><Description>ThermopacAgent SolidWorks Extraction Service</Description></RegistrationInfo>
  <Triggers><LogonTrigger><Delay>PT30S</Delay><Enabled>true</Enabled></LogonTrigger></Triggers>
  <Principals><Principal runLevel="highestAvailable"><GroupId>S-1-5-32-545</GroupId></Principal></Principals>
  <Settings><MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy><ExecutionTimeLimit>PT0S</ExecutionTimeLimit></Settings>
  <Actions><Exec>
    <Command>$InstallDir\run-service.bat</Command>
    <WorkingDirectory>$InstallDir</WorkingDirectory>
  </Exec></Actions>
</Task>
"@
        $taskXml | Set-Content "$env:TEMP\ThermopacAgent-task.xml" -Encoding Unicode
        schtasks /Create /F /XML "$env:TEMP\ThermopacAgent-task.xml" /TN "ThermopacAgent" 2>&1 | Out-Null
        if ($LASTEXITCODE -eq 0) { Write-OK "Scheduled task created (starts 30s after login)" }
        else { Write-Warn "Scheduled task creation failed (non-fatal)" }
    }
}

# ── Summary ───────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  INSTALLATION COMPLETE — ThermopacAgent v$AGENT_VERSION" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Install directory : $InstallDir" -ForegroundColor White
Write-Host "  Data directory    : $DataDir" -ForegroundColor White
Write-Host "  Config file       : $config_path" -ForegroundColor White
Write-Host "  SolidWorks        : $sw_progid" -ForegroundColor White
Write-Host "  Python            : $PyExe" -ForegroundColor White
Write-Host ""
Write-Host "  NEXT STEPS:" -ForegroundColor Cyan
Write-Host ""
Write-Host "  1. Edit config.ini:" -ForegroundColor Yellow
Write-Host "       $config_path" -ForegroundColor White
Write-Host "     Set [cloud] api_url to your Thermopac ERP URL" -ForegroundColor White
Write-Host "     In testing mode: node_token is auto-generated (leave as-is)" -ForegroundColor White
Write-Host "     In production mode: paste admin-issued node_token" -ForegroundColor White
Write-Host ""
Write-Host "  2. Run the agent:" -ForegroundColor Yellow
Write-Host "       $InstallDir\run.bat" -ForegroundColor White
Write-Host "     Or: Start Menu > ThermopacAgent" -ForegroundColor White
Write-Host ""
Write-Host "  3. If SolidWorks COM errors occur, run:" -ForegroundColor Yellow
Write-Host "       $InstallDir\makepy-repair.bat" -ForegroundColor White
Write-Host ""
if (-not $Silent) { Read-Host "Press Enter to close" }
