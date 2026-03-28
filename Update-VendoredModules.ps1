#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Copies the AutomatedLab fork modules into this repo's lib\AutomatedLab\ directory.

.DESCRIPTION
    Run this after making changes to the AutomatedLab fork (c:\projects\AutomatedLab\)
    to update the vendored copy used by Deploy-HomeLab.ps1.

    Copies only the modules needed at runtime (excludes Recipe, Ships, Test).
    Removes the existing vendored copy first to avoid stale files.

.PARAMETER ForkPath
    Path to the AutomatedLab fork repo. Default: c:\projects\AutomatedLab

.EXAMPLE
    .\Update-VendoredModules.ps1

.EXAMPLE
    .\Update-VendoredModules.ps1 -ForkPath D:\repos\AutomatedLab
#>

param(
    [string]$ForkPath = 'c:\projects\AutomatedLab'
)

$ErrorActionPreference = 'Stop'

$vendoredDest = Join-Path $PSScriptRoot 'lib\AutomatedLab'

# Modules to vendor (excludes Recipe, Ships, Test -- they cause parse errors)
$modules = @(
    'AutomatedLab'
    'AutomatedLabCore'
    'AutomatedLab.Common'
    'AutomatedLabDefinition'
    'AutomatedLabNotifications'
    'AutomatedLabUnattended'
    'AutomatedLabWorker'
)

# Verify fork exists
if (-not (Test-Path (Join-Path $ForkPath 'AutomatedLabCore'))) {
    throw "AutomatedLab fork not found at '$ForkPath'. Specify -ForkPath."
}

# Check if vendored modules have been built (published)
# The fork uses PSGallery-format modules installed to PSModulePath.
# If modules are installed there, copy from the installed location.
# Otherwise, copy source directories directly from the fork.

$installedPath = Join-Path $env:ProgramFiles 'WindowsPowerShell\Modules'

Write-Host "Updating vendored modules from: $ForkPath" -ForegroundColor Cyan
Write-Host "Destination: $vendoredDest" -ForegroundColor Cyan

foreach ($mod in $modules) {
    $src = Join-Path $ForkPath $mod
    $dest = Join-Path $vendoredDest $mod

    if (-not (Test-Path $src)) {
        # Try installed location (PSGallery-format after build)
        $src = Join-Path $installedPath $mod
    }

    if (-not (Test-Path $src)) {
        Write-Warning "Module '$mod' not found in fork or installed modules -- skipping"
        continue
    }

    # Remove existing and copy fresh
    if (Test-Path $dest) { Remove-Item $dest -Recurse -Force }
    Copy-Item $src $dest -Recurse -Force
    Write-Host "  Updated: $mod" -ForegroundColor Green
}

# Verify manifest is clean (no Recipe/Ships/Test references)
$manifestPath = Join-Path $vendoredDest 'AutomatedLab\AutomatedLab.psd1'
if (Test-Path $manifestPath) {
    $content = Get-Content $manifestPath -Raw
    $dirty = $false
    foreach ($bad in @('AutomatedLab.Recipe', 'AutomatedLab.Ships', 'AutomatedLabTest')) {
        if ($content -match $bad) {
            $content = $content -replace ".*$bad.*\r?\n", ''
            $dirty = $true
        }
    }
    if ($dirty) {
        Set-Content $manifestPath -Value $content
        Write-Host '  Patched manifest: removed Recipe/Ships/Test references' -ForegroundColor Yellow
    }
}

Write-Host "`nVendored modules updated." -ForegroundColor Green
Write-Host 'Run Pester tests to verify: Invoke-Pester ./Tests/' -ForegroundColor DarkGray
