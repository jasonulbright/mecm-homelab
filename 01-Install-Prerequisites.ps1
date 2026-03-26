#Requires -RunAsAdministrator
#Requires -Version 5.1
<#
.SYNOPSIS
    Installs host-side prerequisites for the MECM home lab.

.DESCRIPTION
    Enables Hyper-V, installs AutomatedLab, creates the LabSources folder
    structure, and checks for required ISOs. Run this once before the other
    scripts.

.EXAMPLE
    .\01-Install-Prerequisites.ps1
#>

$ErrorActionPreference = 'Stop'

# ── Load config ──────────────────────────────────────────────────────────────

$Config = Import-PowerShellDataFile -Path (Join-Path $PSScriptRoot 'config.psd1')

# ── Helper ───────────────────────────────────────────────────────────────────

function Write-Status {
    param([string]$Message, [ValidateSet('OK','WARN','FAIL','INFO')]$Level = 'OK')
    switch ($Level) {
        'OK'   { Write-Host "  [OK]   $Message" -ForegroundColor Green }
        'WARN' { Write-Host "  [WARN] $Message" -ForegroundColor Yellow }
        'FAIL' { Write-Host "  [FAIL] $Message" -ForegroundColor Red }
        'INFO' { Write-Host "  [INFO] $Message" -ForegroundColor Cyan }
    }
}

# ── 1. Hyper-V ───────────────────────────────────────────────────────────────

Write-Host "`n=== Step 1: Hyper-V ===" -ForegroundColor Cyan

$hv = Get-WindowsOptionalFeature -Online -FeatureName Microsoft-Hyper-V -ErrorAction SilentlyContinue
if ($hv.State -ne 'Enabled') {
    Write-Host "  Enabling Hyper-V..." -ForegroundColor Yellow
    Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Hyper-V -All -NoRestart | Out-Null

    Write-Host ""
    Write-Host "  *** REBOOT REQUIRED ***" -ForegroundColor Red
    Write-Host "  Hyper-V has been enabled but requires a reboot." -ForegroundColor Yellow
    Write-Host "  After rebooting, run this script again to continue." -ForegroundColor Yellow
    Write-Host ""
    $reboot = Read-Host "  Reboot now? (y/N)"
    if ($reboot -eq 'y') {
        Restart-Computer -Force
    }
    return
}
Write-Status 'Hyper-V enabled'

# ── 2. AutomatedLab module (vendored) ────────────────────────────────────────

Write-Host "`n=== Step 2: AutomatedLab ===" -ForegroundColor Cyan

$vendoredAL = Join-Path $PSScriptRoot 'lib\AutomatedLab'
$moduleDirs = Get-ChildItem $vendoredAL -Directory | Where-Object { Test-Path (Join-Path $_.FullName '*.psd1') }

# Install vendored modules to PSModulePath if not already there
$targetPath = Join-Path $env:ProgramFiles 'WindowsPowerShell\Modules'
foreach ($mod in $moduleDirs) {
    $dest = Join-Path $targetPath $mod.Name
    if (-not (Test-Path $dest)) {
        Copy-Item $mod.FullName $dest -Recurse -Force
        Write-Status "Installed: $($mod.Name)" -Level INFO
    }
}

$al = Get-Module AutomatedLab -ListAvailable |
    Sort-Object Version -Descending |
    Select-Object -First 1

if (-not $al) {
    throw "AutomatedLab module not found after vendored install. Check lib\AutomatedLab\"
}
Write-Status "AutomatedLab v$($al.Version) (vendored fork)"

# ── 3. LabSources folder ────────────────────────────────────────────────────

Write-Host "`n=== Step 3: LabSources ===" -ForegroundColor Cyan

Import-Module AutomatedLab -ErrorAction Stop

$labSources = Get-LabSourcesLocation -ErrorAction SilentlyContinue
if (-not $labSources -or -not (Test-Path $labSources)) {
    Write-Host "  Creating LabSources folder on C:..." -ForegroundColor Yellow
    New-LabSourcesFolder -DriveLetter C
    $labSources = Get-LabSourcesLocation
}
Write-Status "LabSources: $labSources"

# Create SoftwarePackages subdirectories
$subDirs = @(
    'SoftwarePackages\CM'
    'SoftwarePackages\ADK'
    'SoftwarePackages\ADKPE'
    'SoftwarePackages\ODBC'
    'SoftwarePackages\VCRedist'
    'SoftwarePackages\CMPrereqs'
)
foreach ($sub in $subDirs) {
    $fullPath = Join-Path $labSources $sub
    if (-not (Test-Path $fullPath)) {
        New-Item -Path $fullPath -ItemType Directory -Force | Out-Null
        Write-Status "Created: $fullPath" -Level INFO
    }
}

# ── 4. ISO and download checklist ────────────────────────────────────────────

Write-Host "`n=== Step 4: Download Checklist ===" -ForegroundColor Cyan
Write-Host ""

$isoPath = Join-Path $labSources 'ISOs'
$swPkg   = Join-Path $labSources 'SoftwarePackages'

$downloads = @(
    @{
        Name   = 'Windows Server 2025 Evaluation ISO'
        Path   = $isoPath
        URL    = 'https://www.microsoft.com/en-us/evalcenter/evaluate-windows-server-2025'
        Exists = [bool](Get-ChildItem $isoPath -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match 'SERVER_EVAL' -and $_.Extension -eq '.iso' })
    }
    @{
        Name   = 'SQL Server 2022 Evaluation ISO'
        Path   = $isoPath
        URL    = 'https://www.microsoft.com/en-us/evalcenter/download-sql-server-2022'
        Exists = [bool](Get-ChildItem $isoPath -Filter '*SQL*2022*' -ErrorAction SilentlyContinue)
    }
    @{
        Name   = 'ConfigMgr 2509 Baseline (extract contents into CM folder)'
        Path   = (Join-Path $swPkg 'CM')
        URL    = 'https://www.microsoft.com/en-us/evalcenter/download-microsoft-endpoint-configuration-manager'
        Exists = [bool](Get-ChildItem (Join-Path $swPkg 'CM') -ErrorAction SilentlyContinue |
            Where-Object { $_.PSIsContainer })
    }
    @{
        Name   = 'Windows ADK (Dec 2024) - adksetup.exe'
        Path   = (Join-Path $swPkg 'ADK')
        URL    = 'https://go.microsoft.com/fwlink/?linkid=2289980'
        Exists = [bool](Get-ChildItem (Join-Path $swPkg 'ADK') -Filter 'adksetup*' -ErrorAction SilentlyContinue)
    }
    @{
        Name   = 'Windows PE add-on - adkwinpesetup.exe'
        Path   = (Join-Path $swPkg 'ADKPE')
        URL    = 'https://go.microsoft.com/fwlink/?linkid=2289981'
        Exists = [bool](Get-ChildItem (Join-Path $swPkg 'ADKPE') -Filter 'adkwinpe*' -ErrorAction SilentlyContinue)
    }
    @{
        Name   = 'Windows 11 Enterprise ISO (optional, for client VM)'
        Path   = $isoPath
        URL    = 'https://www.microsoft.com/en-us/evalcenter/evaluate-windows-11-enterprise'
        Exists = [bool](Get-ChildItem $isoPath -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match 'CLIENTENTERPRISE' -and $_.Extension -eq '.iso' })
    }
)

$allFound = $true
foreach ($dl in $downloads) {
    if ($dl.Exists) {
        Write-Status $dl.Name -Level OK
    } else {
        $allFound = $false
        Write-Status $dl.Name -Level FAIL
        Write-Host "         Path: $($dl.Path)" -ForegroundColor DarkGray
        Write-Host "         URL:  $($dl.URL)" -ForegroundColor DarkGray
    }
}

# ── 5. RAM check ─────────────────────────────────────────────────────────────

Write-Host "`n=== Host Resources ===" -ForegroundColor Cyan

$totalRam = (Get-CimInstance Win32_PhysicalMemory | Measure-Object Capacity -Sum).Sum
$totalRamGB = [math]::Round($totalRam / 1GB, 0)
if ($totalRamGB -ge 64) {
    Write-Status "Host RAM: ${totalRamGB}GB (excellent)"
} elseif ($totalRamGB -ge 32) {
    Write-Status "Host RAM: ${totalRamGB}GB (sufficient, 64GB recommended)" -Level WARN
} else {
    Write-Status "Host RAM: ${totalRamGB}GB (minimum 32GB required)" -Level FAIL
}

# ── Summary ──────────────────────────────────────────────────────────────────

Write-Host ""
if ($allFound) {
    Write-Host "All prerequisites found! Next step:" -ForegroundColor Green
    Write-Host "  .\02-Download-Offline.ps1" -ForegroundColor White
} else {
    Write-Host "Download the missing items above, then run this script again." -ForegroundColor Yellow
    Write-Host "Once all items show [OK], proceed to:" -ForegroundColor Yellow
    Write-Host "  .\02-Download-Offline.ps1" -ForegroundColor White
}
Write-Host ""
