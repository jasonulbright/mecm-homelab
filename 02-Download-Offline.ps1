#Requires -RunAsAdministrator
#Requires -Version 5.1
<#
.SYNOPSIS
    Downloads offline content needed for the MECM home lab.

.DESCRIPTION
    Creates ADK offline layouts, downloads CM prerequisites, ODBC Driver 18,
    and VC++ 14.50 runtimes. All downloads go into LabSources\SoftwarePackages
    so they can be copied into VMs later.

    Run after 01-Install-Prerequisites.ps1 has completed successfully.

.EXAMPLE
    .\02-Download-Offline.ps1
#>

$ErrorActionPreference = 'Stop'

# ── Load config ──────────────────────────────────────────────────────────────

$Config = Import-PowerShellDataFile -Path (Join-Path $PSScriptRoot 'config.psd1')

# ── Paths ────────────────────────────────────────────────────────────────────

Import-Module AutomatedLab -ErrorAction Stop
$labSources = Get-LabSourcesLocation
$swPkg      = Join-Path $labSources 'SoftwarePackages'

$adkSource    = Join-Path $swPkg 'ADK'
$adkPeSource  = Join-Path $swPkg 'ADKPE'
$cmSource     = Join-Path $swPkg 'CM'
$prereqDest   = Join-Path $swPkg 'CMPrereqs'
$odbcDest     = Join-Path $swPkg 'ODBC'
$vcDest       = Join-Path $swPkg 'VCRedist'

# Ensure directories exist
@($adkSource, $adkPeSource, $cmSource, $prereqDest, $odbcDest, $vcDest) | ForEach-Object {
    if (-not (Test-Path $_)) { New-Item -Path $_ -ItemType Directory -Force | Out-Null }
}

# ── Helper ───────────────────────────────────────────────────────────────────

function Write-Step {
    param([string]$Title)
    Write-Host "`n=== $Title ===" -ForegroundColor Cyan
}

function Write-Status {
    param([string]$Message, [ValidateSet('OK','SKIP','WARN','FAIL','RUN')]$Level = 'OK')
    switch ($Level) {
        'OK'   { Write-Host "  [OK]   $Message" -ForegroundColor Green }
        'SKIP' { Write-Host "  [SKIP] $Message" -ForegroundColor DarkGray }
        'WARN' { Write-Host "  [WARN] $Message" -ForegroundColor Yellow }
        'FAIL' { Write-Host "  [FAIL] $Message" -ForegroundColor Red }
        'RUN'  { Write-Host "  [RUN]  $Message" -ForegroundColor Yellow }
    }
}

# ── 1. ADK Offline Layout ───────────────────────────────────────────────────

Write-Step 'ADK Offline Layout'

$adkSetup = Get-ChildItem $adkSource -Filter 'adksetup*' -ErrorAction SilentlyContinue | Select-Object -First 1
$adkLayout = Join-Path $adkSource 'Offline'

if (-not $adkSetup) {
    Write-Status 'adksetup.exe not found in ADK folder. Download it first (see 01 script).' -Level FAIL
} elseif (Test-Path (Join-Path $adkLayout 'Installers')) {
    Write-Status "ADK offline layout already exists at: $adkLayout" -Level SKIP
} else {
    Write-Status "Creating ADK offline layout (this may take several minutes)..." -Level RUN
    $proc = Start-Process -FilePath $adkSetup.FullName `
        -ArgumentList @('/quiet', '/layout', $adkLayout) `
        -Wait -PassThru -NoNewWindow
    if ($proc.ExitCode -ne 0) {
        Write-Status "ADK layout failed with exit code $($proc.ExitCode)" -Level FAIL
    } else {
        # Copy the bootstrapper into the layout so the VM has everything
        Copy-Item -Path $adkSetup.FullName -Destination $adkLayout -Force
        Write-Status "ADK offline layout created at: $adkLayout"
    }
}

# ── 2. ADK WinPE Offline Layout ─────────────────────────────────────────────

Write-Step 'ADK WinPE Offline Layout'

$adkPeSetup = Get-ChildItem $adkPeSource -Filter 'adkwinpe*' -ErrorAction SilentlyContinue | Select-Object -First 1
$adkPeLayout = Join-Path $adkPeSource 'Offline'

if (-not $adkPeSetup) {
    Write-Status 'adkwinpesetup.exe not found in ADKPE folder. Download it first.' -Level FAIL
} elseif (Test-Path (Join-Path $adkPeLayout 'Installers')) {
    Write-Status "ADK PE offline layout already exists at: $adkPeLayout" -Level SKIP
} else {
    Write-Status "Creating ADK PE offline layout (this may take several minutes)..." -Level RUN
    $proc = Start-Process -FilePath $adkPeSetup.FullName `
        -ArgumentList @('/quiet', '/layout', $adkPeLayout) `
        -Wait -PassThru -NoNewWindow
    if ($proc.ExitCode -ne 0) {
        Write-Status "ADK PE layout failed with exit code $($proc.ExitCode)" -Level FAIL
    } else {
        Copy-Item -Path $adkPeSetup.FullName -Destination $adkPeLayout -Force
        Write-Status "ADK PE offline layout created at: $adkPeLayout"
    }
}

# ── 3. CM Prerequisites (setupdl.exe) ───────────────────────────────────────

Write-Step 'ConfigMgr Prerequisites Download'

# Find setupdl.exe in the CM source tree
$setupDl = Get-ChildItem $cmSource -Recurse -Filter 'setupdl.exe' -ErrorAction SilentlyContinue | Select-Object -First 1

if (-not $setupDl) {
    Write-Status 'setupdl.exe not found in CM source folder. Extract CM baseline first.' -Level FAIL
} elseif ((Get-ChildItem $prereqDest -ErrorAction SilentlyContinue | Measure-Object).Count -gt 10) {
    Write-Status "CM prerequisites already downloaded ($prereqDest)" -Level SKIP
} else {
    Write-Status "Downloading CM prerequisites (requires internet)..." -Level RUN
    $proc = Start-Process -FilePath $setupDl.FullName `
        -ArgumentList @($prereqDest) `
        -Wait -PassThru -NoNewWindow
    if ($proc.ExitCode -ne 0) {
        Write-Status "setupdl.exe failed with exit code $($proc.ExitCode)" -Level FAIL
    } else {
        $fileCount = (Get-ChildItem $prereqDest | Measure-Object).Count
        Write-Status "CM prerequisites downloaded ($fileCount files)"
    }
}

# ── 4. ODBC Driver 18.5.2.1 ─────────────────────────────────────────────────

Write-Step "ODBC Driver $($Config.ODBCVersion)"

$odbcMsi = Get-ChildItem $odbcDest -Filter '*.msi' -ErrorAction SilentlyContinue | Select-Object -First 1

if ($odbcMsi) {
    Write-Status "ODBC MSI already exists: $($odbcMsi.Name)" -Level SKIP
} else {
    Write-Status "Downloading ODBC Driver $($Config.ODBCVersion)..." -Level RUN
    $odbcFile = Join-Path $odbcDest 'msodbcsql.msi'
    try {
        Invoke-WebRequest -Uri $Config.ODBCURL -OutFile $odbcFile -UseBasicParsing
        Write-Status "ODBC Driver downloaded: $odbcFile"
    } catch {
        Write-Status "ODBC download failed: $_" -Level FAIL
    }
}

# ── 5. VC++ 14.50 Runtimes ──────────────────────────────────────────────────

Write-Step 'VC++ 14.50 Runtimes (VS 2026)'

$vcX64 = Join-Path $vcDest 'vc_redist.x64.exe'
$vcX86 = Join-Path $vcDest 'vc_redist.x86.exe'

if (Test-Path $vcX64) {
    Write-Status "vc_redist.x64.exe already exists" -Level SKIP
} else {
    Write-Status "Downloading vc_redist.x64.exe..." -Level RUN
    try {
        Invoke-WebRequest -Uri $Config.VCRedistX64URL -OutFile $vcX64 -UseBasicParsing
        Write-Status "vc_redist.x64.exe downloaded"
    } catch {
        Write-Status "VC++ x64 download failed: $_" -Level FAIL
    }
}

if (Test-Path $vcX86) {
    Write-Status "vc_redist.x86.exe already exists" -Level SKIP
} else {
    Write-Status "Downloading vc_redist.x86.exe..." -Level RUN
    try {
        Invoke-WebRequest -Uri $Config.VCRedistX86URL -OutFile $vcX86 -UseBasicParsing
        Write-Status "vc_redist.x86.exe downloaded"
    } catch {
        Write-Status "VC++ x86 download failed: $_" -Level FAIL
    }
}

# ── Summary ──────────────────────────────────────────────────────────────────

Write-Host "`n=== Download Summary ===" -ForegroundColor Cyan

$items = @(
    @{ Name = 'ADK Offline Layout';   OK = (Test-Path (Join-Path $adkLayout 'Installers' -ErrorAction SilentlyContinue)) }
    @{ Name = 'ADK PE Offline Layout'; OK = (Test-Path (Join-Path $adkPeLayout 'Installers' -ErrorAction SilentlyContinue)) }
    @{ Name = 'CM Prerequisites';     OK = ((Get-ChildItem $prereqDest -ErrorAction SilentlyContinue | Measure-Object).Count -gt 10) }
    @{ Name = 'ODBC Driver MSI';      OK = [bool](Get-ChildItem $odbcDest -Filter '*.msi' -ErrorAction SilentlyContinue) }
    @{ Name = 'VC++ x64 Runtime';     OK = (Test-Path $vcX64) }
    @{ Name = 'VC++ x86 Runtime';     OK = (Test-Path $vcX86) }
)

$allOK = $true
foreach ($item in $items) {
    if ($item.OK) {
        Write-Host "  [OK]   $($item.Name)" -ForegroundColor Green
    } else {
        $allOK = $false
        Write-Host "  [FAIL] $($item.Name)" -ForegroundColor Red
    }
}

Write-Host ""
if ($allOK) {
    Write-Host "All downloads complete! Next step:" -ForegroundColor Green
    Write-Host "  .\03-Deploy-Infrastructure.ps1" -ForegroundColor White
} else {
    Write-Host "Some downloads failed. Fix the issues above and re-run this script." -ForegroundColor Yellow
}
Write-Host ""
