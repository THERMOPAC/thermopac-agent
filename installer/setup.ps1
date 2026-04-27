#Requires -RunAsAdministrator
<#
.SYNOPSIS
    ThermopacStructuringAgent Windows Installer (PowerShell)
.DESCRIPTION
    Downloads Python 3.11 Embeddable, installs pywin32 + requests,
    generates SolidWorks COM type library cache, and creates run scripts.
    Used when the Inno Setup .exe installer did not bundle Python.
.PARAMETER InstallDir
    Installation directory (default: C:\Program Files\ThermopacStructuringAgent)
.PARAMETER DataDir
    Runtime data directory for logs and temp files (default: C:\ThermopacStructurer)
.PARAMETER StagingRoot
    Root folder where structured drawings are staged
    (default: C:\ThermopacStaging\drawings)
.PARAMETER Silent
    Skip interactive prompts; use defaults for everything
.EXAMPLE
    powershell -ExecutionPolicy Bypass -File setup.ps1
    powershell -ExecutionPolicy Bypass -File setup.ps1 -Silent
#>
param(
    [string]$InstallDir   = "C:\Program Files\ThermopacStructuringAgent",
    [string]$DataDir      = "C:\ThermopacStructurer",
    [string]$StagingRoot  = "C:\ThermopacStaging\drawings",
    [switch]$Silent
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ── Constants ──────────────────────────────────────────────────────────────────
$AGENT_VERSION  = "1.0.0"
$PY_VERSION     = "3.11.9"
$PY_EMBED_ZIP   = "python-$PY_VERSION-embed-amd64.zip"
$PY_EMBED_URL   = "https://www.python.org/ftp/python/$PY_VERSION/$PY_EMBED_ZIP"
$GET_PIP_URL    = "https://bootstrap.pypa.io/get-pip.py"
$REQUIRED_PKGS  = @("pywin32>=306", "requests>=2.28.0")
$APP_URL        = "https://thermopac-communication-thermopacllp.replit.app"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
# When run via Inno Setup post-install, ScriptDir = {app}
# When run standalone, ScriptDir = structurer_pkg\installer\
$SourceDir = $ScriptDir

# ── Helpers ────────────────────────────────────────────────────────────────────
function Write-Header {
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "  ThermopacStructuringAgent $AGENT_VERSION — Windows Setup" -ForegroundColor Cyan
    Write-Host "  THERMOPAC ERP  |  SolidWorks Drawing Structuring Agent" -ForegroundColor Cyan
    Write-Host "  Phase 1 — WRITE ONLY" -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host ""
}

function Write-Step([string]$msg) { Write-Host "[STEP] $msg" -ForegroundColor Yellow }
function Write-OK([string]$msg)   { Write-Host "  [OK] $msg"   -ForegroundColor Green }
function Write-Warn([string]$msg) { Write-Host "  [WARN] $msg" -ForegroundColor DarkYellow }
function Write-Fail([string]$msg) { Write-Host "  [FAIL] $msg" -ForegroundColor Red }

function Download-File([string]$Url, [string]$Dest) {
    Write-Host "    Downloading: $Url" -ForegroundColor DarkGray
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $wc = New-Object System.Net.WebClient
    $wc.Headers.Add("User-Agent", "ThermopacStructurer-Installer/$AGENT_VERSION")
    $wc.DownloadFile($Url, $Dest)
    Write-OK "Downloaded to: $Dest"
}

# ── Step 1: Prerequisites ──────────────────────────────────────────────────────
Write-Header
Write-Step "Checking prerequisites..."

$os = [System.Environment]::OSVersion.Version
if ($os.Major -lt 10) {
    Write-Fail "Windows 10 or later is required (found: $($os.Major).$($os.Minor))"
    exit 1
}
Write-OK "Windows $($os.Major).$($os.Minor) — OK"

# ── Step 2: Detect SolidWorks ─────────────────────────────────────────────────
Write-Step "Detecting SolidWorks installation..."
$sw_progid  = $null
$sw_version = 0
$sw_dir     = $null

$sw_map = @{
    2024 = "SldWorks.Application.32"
    2023 = "SldWorks.Application.31"
    2022 = "SldWorks.Application.30"
    2021 = "SldWorks.Application.29"
    2020 = "SldWorks.Application.28"
    2019 = "SldWorks.Application.27"
}

foreach ($year in ($sw_map.Keys | Sort-Object -Descending)) {
    $progid = $sw_map[$year]
    try {
        $key = [Microsoft.Win32.Registry]::ClassesRoot.OpenSubKey($progid)
        if ($null -ne $key) {
            $sw_version = $year
            $sw_progid  = $progid
            $key.Close()
            foreach ($base in @(
                "SOFTWARE\SolidWorks\SolidWorks $year\Setup",
                "SOFTWARE\WOW6432Node\SolidWorks\SolidWorks $year\Setup"
            )) {
                try {
                    $rk = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($base)
                    if ($null -ne $rk) {
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
    Write-Fail "SolidWorks NOT detected. Install SolidWorks 2019–2024 before running."
    if (-not $Silent) { Read-Host "Press Enter to exit" }
    exit 1
}
Write-OK "SolidWorks $sw_version ($sw_progid)"
if ($sw_dir) { Write-OK "SolidWorks directory: $sw_dir" }

# ── Step 3: Create directories ────────────────────────────────────────────────
Write-Step "Creating directories..."
foreach ($d in @($InstallDir, "$DataDir\logs", "$DataDir\temp", $StagingRoot)) {
    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
    Write-OK $d
}

# Set ACLs on data/staging dirs so any user can write
foreach ($d in @($DataDir, $StagingRoot)) {
    try {
        $acl  = Get-Acl $d
        $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            "Everyone", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
        $acl.SetAccessRule($rule)
        Set-Acl $d $acl
        Write-OK "Permissions set on $d"
    } catch {
        Write-Warn "Could not set ACL on $d (non-fatal): $_"
    }
}

# ── Step 4: Python Embeddable ─────────────────────────────────────────────────
Write-Step "Setting up Python $PY_VERSION (embedded)..."
$PyDir = "$InstallDir\python"
$PyExe = "$PyDir\python.exe"
$ZipDest = "$env:TEMP\$PY_EMBED_ZIP"

if (-not (Test-Path $PyExe)) {
    if (-not (Test-Path $ZipDest)) {
        try { Download-File $PY_EMBED_URL $ZipDest }
        catch {
            Write-Fail "Download failed: $_"
            exit 1
        }
    } else {
        Write-OK "Using cached download: $ZipDest"
    }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipDest, $PyDir)
    Write-OK "Extracted to $PyDir"
} else {
    Write-OK "Python already present: $PyExe"
}

# Enable site-packages
$pthFile = Get-ChildItem $PyDir -Filter "python*._pth" | Select-Object -First 1
if ($pthFile) {
    $pthContent = Get-Content $pthFile.FullName -Raw
    if ($pthContent -match "#import site") {
        $pthContent = $pthContent -replace "#import site", "import site"
        Set-Content $pthFile.FullName $pthContent -NoNewline
        Write-OK "site-packages enabled in $($pthFile.Name)"
    }
    if ($pthContent -notmatch "Scripts") {
        Add-Content $pthFile.FullName "`nScripts"
    }
}

# ── Step 5: pip ───────────────────────────────────────────────────────────────
Write-Step "Installing pip..."
$PipScript = "$env:TEMP\get-pip.py"
$PipExe    = "$PyDir\Scripts\pip.exe"

if (-not (Test-Path $PipExe)) {
    if (-not (Test-Path $PipScript)) {
        try { Download-File $GET_PIP_URL $PipScript }
        catch { Write-Fail "Could not download get-pip.py: $_"; exit 1 }
    }
    & $PyExe $PipScript --no-warn-script-location 2>&1 |
        ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
    if ($LASTEXITCODE -ne 0) { Write-Fail "pip installation failed"; exit 1 }
    Write-OK "pip installed"
} else {
    Write-OK "pip already present"
}

# ── Step 6: Python packages ───────────────────────────────────────────────────
Write-Step "Installing Python packages..."
foreach ($pkg in $REQUIRED_PKGS) {
    Write-Host "    pip install $pkg ..."
    & $PyExe -m pip install $pkg --no-warn-script-location 2>&1 |
        ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
    if ($LASTEXITCODE -ne 0) { Write-Fail "Failed to install $pkg"; exit 1 }
    Write-OK $pkg
}

# ── Step 7: pywin32 post-install ─────────────────────────────────────────────
Write-Step "Running pywin32 post-install..."
$pywin32post = "$PyDir\Scripts\pywin32_postinstall.py"
if (-not (Test-Path $pywin32post)) {
    $pywin32post = "$PyDir\Lib\site-packages\pywin32_postinstall.py"
}
if (Test-Path $pywin32post) {
    & $PyExe $pywin32post -install 2>&1 |
        ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
    Write-OK "pywin32 post-install complete"
} else {
    Write-Warn "pywin32_postinstall.py not found (non-fatal)"
}

# ── Step 8: SolidWorks makepy cache ──────────────────────────────────────────
Write-Step "Generating SolidWorks COM type library cache (makepy)..."
$makepy_ok = $false

# Method A
Write-Host "    [Makepy-A] python -m win32com.client.makepy `"$sw_progid`"..."
$out = & $PyExe -m win32com.client.makepy $sw_progid 2>&1
if ($LASTEXITCODE -eq 0) {
    Write-OK "Makepy cache generated via ProgID"
    $makepy_ok = $true
} else {
    Write-Warn "Method A failed: $out"
}

# Method B: TLB file scan
if (-not $makepy_ok -and $sw_dir) {
    Write-Host "    [Makepy-B] Scanning SolidWorks directory for TLB files..."
    $tlb_cands = @()
    foreach ($ext in @("*.tlb", "sldworks.exe")) {
        $tlb_cands += Get-ChildItem -Path $sw_dir -Filter $ext -ErrorAction SilentlyContinue |
            Select-Object -First 3
    }
    foreach ($sub in @("api\redist", "api")) {
        $sp = Join-Path $sw_dir $sub
        if (Test-Path $sp) {
            Get-ChildItem -Path $sp -Filter "*.tlb" -ErrorAction SilentlyContinue |
                ForEach-Object { $tlb_cands += $_ }
        }
    }
    foreach ($tlb in $tlb_cands) {
        $out = & $PyExe -m win32com.client.makepy "`"$($tlb.FullName)`"" 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-OK "Makepy cache from $($tlb.Name)"
            $makepy_ok = $true
            break
        }
    }
    if (-not $makepy_ok) { Write-Warn "Method B: no TLB produced cache" }
}

if (-not $makepy_ok) {
    Write-Warn "SolidWorks COM cache could not be pre-generated."
    Write-Host "  The agent will generate the cache at runtime (3 fallback methods)." -ForegroundColor Yellow
}

# ── Step 9: Copy agent files ──────────────────────────────────────────────────
Write-Step "Installing agent files to $InstallDir..."

# Source detection: works from installer\ subfolder OR from {app} (Inno Setup)
$agent_dirs = @("agent", "extractor", "structurer")
foreach ($dir in $agent_dirs) {
    # Try ScriptDir (when setup.ps1 is in installer\ next to the source)
    $src = Join-Path (Split-Path -Parent $ScriptDir) $dir
    if (-not (Test-Path $src)) {
        # When Inno Setup placed setup.ps1 in {app}, source is {app} itself
        $src = Join-Path $ScriptDir $dir
    }
    if (-not (Test-Path $src)) {
        Write-Warn "Source not found (non-fatal): $dir"
        continue
    }
    $dest = Join-Path $InstallDir $dir
    if (Test-Path $dest) { Remove-Item $dest -Recurse -Force }
    Copy-Item $src $dest -Recurse -Force
    Write-OK "Installed: $dir\"
}

# ── Step 10: Write config.ini ─────────────────────────────────────────────────
Write-Step "Creating config.ini..."
$config_path = "$InstallDir\config.ini"
if (-not (Test-Path $config_path)) {
    $cfg = @"
; ThermopacStructuringAgent configuration
; Generated by setup.ps1 v$AGENT_VERSION

[cloud]
api_url    = $APP_URL
node_id    = $env:COMPUTERNAME
node_token = REPLACE_WITH_YOUR_TOKEN

[agent]
; testing | production
mode              = production
poll_interval_sec = 10
job_timeout_sec   = 600
max_retries       = 3

[paths]
temp_dir = $DataDir\temp
log_dir  = $DataDir\logs

[solidworks]
; Explicit ProgID — takes priority over solidworks_version.
; SolidWorks 2019=.27  2020=.28  2021=.29  2022=.30  2023=.31  2024=.32
solidworks_progid = $sw_progid
visible           = false
model_search_path =

[structurer]
; REQUIRED: absolute path to the approved SolidWorks drawing template (.drwdot)
template_path =
; Root folder where structured drawings are staged
staging_root  = $StagingRoot
"@
    Set-Content $config_path $cfg -Encoding UTF8
    Write-OK "Created: $config_path"
} else {
    Write-OK "Existing config.ini preserved: $config_path"
}

# ── Step 11: Write run scripts ────────────────────────────────────────────────
Write-Step "Creating run scripts..."

$run_bat = @"
@echo off
title ThermopacStructuringAgent v$AGENT_VERSION
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8
echo.
echo  ThermopacStructuringAgent -- SolidWorks Drawing Structuring Agent
echo  THERMOPAC ERP Integration  ^|  Phase 1  ^|  v$AGENT_VERSION
echo.
"$PyExe" "$InstallDir\agent\main_structurer.py" --config "$config_path" %*
pause
"@
Set-Content "$InstallDir\run.bat" $run_bat -Encoding ASCII
Write-OK "Created: $InstallDir\run.bat"

$run_svc_bat = @"
@echo off
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8
start "" /B "$PyExe" "$InstallDir\agent\main_structurer.py" --config "$config_path"
"@
Set-Content "$InstallDir\run-service.bat" $run_svc_bat -Encoding ASCII
Write-OK "Created: $InstallDir\run-service.bat"

$makepy_bat = @"
@echo off
echo ThermopacStructuringAgent -- SolidWorks COM Cache Repair
echo.
"$PyExe" -m win32com.client.makepy "$sw_progid"
if errorlevel 1 (
    echo Method A failed. Trying gencache...
    "$PyExe" -c "import win32com.client.gencache as g; g.EnsureDispatch('$sw_progid')"
)
echo.
echo Done. Restart ThermopacStructuringAgent.
pause
"@
Set-Content "$InstallDir\makepy-repair.bat" $makepy_bat -Encoding ASCII
Write-OK "Created: $InstallDir\makepy-repair.bat"

# ── Step 12: Start Menu shortcut ─────────────────────────────────────────────
Write-Step "Creating Start Menu shortcuts..."
try {
    $StartMenu   = [System.Environment]::GetFolderPath("CommonPrograms")
    $ShortcutDir = "$StartMenu\ThermopacStructuringAgent"
    if (-not (Test-Path $ShortcutDir)) {
        New-Item -ItemType Directory -Path $ShortcutDir -Force | Out-Null
    }
    $WshShell = New-Object -ComObject WScript.Shell

    $sc = $WshShell.CreateShortcut("$ShortcutDir\ThermopacStructuringAgent.lnk")
    $sc.TargetPath = "$InstallDir\run.bat"
    $sc.WorkingDirectory = $InstallDir
    $sc.Description = "ThermopacStructuringAgent — SolidWorks Drawing Structuring Agent"
    $sc.Save()

    $sc2 = $WshShell.CreateShortcut("$ShortcutDir\Edit Config.lnk")
    $sc2.TargetPath = "notepad.exe"
    $sc2.Arguments  = "`"$config_path`""
    $sc2.Description = "Edit ThermopacStructuringAgent configuration"
    $sc2.Save()

    $sc3 = $WshShell.CreateShortcut("$ShortcutDir\Repair COM Cache.lnk")
    $sc3.TargetPath = "$InstallDir\makepy-repair.bat"
    $sc3.WorkingDirectory = $InstallDir
    $sc3.Description = "Regenerate SolidWorks COM type library cache"
    $sc3.Save()

    Write-OK "Start Menu: $ShortcutDir"
} catch {
    Write-Warn "Could not create Start Menu shortcuts: $_"
}

# ── Step 13: Desktop shortcut ─────────────────────────────────────────────────
Write-Step "Creating Desktop shortcut..."
try {
    $Desktop = [System.Environment]::GetFolderPath("CommonDesktopDirectory")
    $WshShell = New-Object -ComObject WScript.Shell
    $sc = $WshShell.CreateShortcut("$Desktop\SolidWorks Structuring Agent.lnk")
    $sc.TargetPath = "$InstallDir\run.bat"
    $sc.WorkingDirectory = $InstallDir
    $sc.Description = "ThermopacStructuringAgent — SolidWorks Drawing Structuring Agent"
    $sc.Save()
    Write-OK "Desktop shortcut created"
} catch {
    Write-Warn "Could not create Desktop shortcut: $_"
}

# ── Step 14: Optional scheduled task ─────────────────────────────────────────
if (-not $Silent) {
    Write-Host ""
    $ans = Read-Host "Create scheduled task to auto-start agent at Windows login? [y/N]"
    if ($ans -match '^[Yy]') {
        Write-Step "Creating scheduled task..."
        $taskXml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo><Description>ThermopacStructuringAgent SolidWorks Drawing Structuring Service</Description></RegistrationInfo>
  <Triggers><LogonTrigger><Delay>PT30S</Delay><Enabled>true</Enabled></LogonTrigger></Triggers>
  <Principals><Principal runLevel="highestAvailable"><GroupId>S-1-5-32-545</GroupId></Principal></Principals>
  <Settings><MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy><ExecutionTimeLimit>PT0S</ExecutionTimeLimit></Settings>
  <Actions><Exec>
    <Command>$InstallDir\run-service.bat</Command>
    <WorkingDirectory>$InstallDir</WorkingDirectory>
  </Exec></Actions>
</Task>
"@
        $taskXml | Set-Content "$env:TEMP\ThermopacStructurer-task.xml" -Encoding Unicode
        schtasks /Create /F /XML "$env:TEMP\ThermopacStructurer-task.xml" /TN "ThermopacStructuringAgent" 2>&1 | Out-Null
        if ($LASTEXITCODE -eq 0) { Write-OK "Scheduled task created (starts 30s after login)" }
        else { Write-Warn "Scheduled task creation failed (non-fatal)" }
    }
}

# ── Summary ───────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  INSTALLATION COMPLETE — ThermopacStructuringAgent v$AGENT_VERSION" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Install directory : $InstallDir" -ForegroundColor White
Write-Host "  Data directory    : $DataDir"    -ForegroundColor White
Write-Host "  Staging root      : $StagingRoot" -ForegroundColor White
Write-Host "  Config file       : $config_path" -ForegroundColor White
Write-Host "  SolidWorks        : $sw_progid"  -ForegroundColor White
Write-Host "  Python            : $PyExe"      -ForegroundColor White
Write-Host ""
Write-Host "  NEXT STEPS:" -ForegroundColor Cyan
Write-Host ""
Write-Host "  1. Edit config.ini and set:" -ForegroundColor Yellow
Write-Host "       [cloud]  node_id, node_token" -ForegroundColor White
Write-Host "       [structurer]  template_path  <- REQUIRED" -ForegroundColor White
Write-Host ""
Write-Host "  2. Launch the agent:" -ForegroundColor Yellow
Write-Host "       Start Menu > ThermopacStructuringAgent" -ForegroundColor White
Write-Host "       OR Desktop shortcut: SolidWorks Structuring Agent" -ForegroundColor White
Write-Host ""
Write-Host "  3. If COM errors appear:" -ForegroundColor Yellow
Write-Host "       Start Menu > ThermopacStructuringAgent > Repair COM Cache" -ForegroundColor White
Write-Host ""

if (-not $Silent) { Read-Host "Press Enter to exit" }
