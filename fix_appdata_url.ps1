# fix_appdata_url.ps1
# Automatically corrects the api_url in the APPDATA overlay config.
# Called by run.bat on every startup — no user action required.
# Preserves node_token, mode, and all other settings.

$cfg  = "$env:APPDATA\ThermopacStructuringAgent\config.ini"
$dev  = "https://5d05ae61-8225-4651-bb76-b4e20a4ddabb-00-3mex6zlihlmft.janeway.replit.dev"
$prod = "https://thermopac-communication-thermopacllp.replit.app"

if (-not (Test-Path $cfg)) { exit 0 }

$txt = [IO.File]::ReadAllText($cfg, [Text.UTF8Encoding]::new($false))

# Case 1: production URL present anywhere in the file → replace it
if ($txt.Contains($prod)) {
    $txt = $txt.Replace($prod, $dev)
    [IO.File]::WriteAllText($cfg, $txt, [Text.UTF8Encoding]::new($false))
    Write-Host "[CONFIG FIX] APPDATA api_url updated to dev URL"
    exit 0
}

# Case 2: api_url exists but points to some other non-dev URL → replace it
if ($txt -match "(?m)^api_url\s*=" -and -not $txt.Contains($dev)) {
    $txt = [regex]::Replace($txt, "(?m)^api_url\s*=.*", "api_url = $dev")
    [IO.File]::WriteAllText($cfg, $txt, [Text.UTF8Encoding]::new($false))
    Write-Host "[CONFIG FIX] APPDATA api_url updated to dev URL"
    exit 0
}

# Case 3: api_url not set at all → add it
if (-not ($txt -match "(?m)^api_url\s*=")) {
    if ($txt -match "(?m)^\[cloud\]") {
        $txt = [regex]::Replace($txt, "(?m)^\[cloud\]", "[cloud]`r`napi_url = $dev")
    } else {
        $txt = $txt.TrimEnd() + "`r`n`r`n[cloud]`r`napi_url = $dev`r`n"
    }
    [IO.File]::WriteAllText($cfg, $txt, [Text.UTF8Encoding]::new($false))
    Write-Host "[CONFIG FIX] APPDATA api_url set to dev URL"
    exit 0
}

# Already correct — nothing to do
exit 0
