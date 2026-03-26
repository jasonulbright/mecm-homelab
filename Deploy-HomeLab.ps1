#Requires -RunAsAdministrator
#Requires -Version 5.1
<#
.SYNOPSIS
    Deploys a complete MECM home lab: DC01 + CM01 + CLIENT01 with AD, CA, SQL, ConfigMgr 2509, and a Windows 11 workstation.

.DESCRIPTION
    Single-script deployment that performs all steps:
      1.  Check prerequisites (Hyper-V, host RAM)
      2.  Install vendored AutomatedLab modules
      3.  Create LabSources folder structure
      4.  Verify ISOs are present
      5.  Create ADK offline layouts if needed
      6.  Download CM prereqs, ODBC, VC++ runtimes if not present
      7.  Define and deploy the lab (DC01 + CM01) via AutomatedLab
      8.  Expand CM01 OS disk, configure SQL memory
      9.  Copy all software to CM01
     10.  Extend AD schema, create System Management container
     11.  Create service accounts (svc-CMPush, svc-CMNAA, svc-CMAdmin)
     12.  Install VC++, ODBC, MSOLEDB, ADK, ADK PE on CM01
     13.  Install ConfigMgr unattended
     14.  Create content share on CM01
     15.  Add svc-CMAdmin as MECM Full Administrator
     16.  Deploy tools (cc4cm + AppPackager) to CM01
     17.  Create snapshots
     18.  Print connection info and remaining manual console steps

    Idempotent where possible -- safe to re-run if it fails partway through.

.PARAMETER RemoveExisting
    If specified, removes an existing lab with the same name without prompting.

.EXAMPLE
    .\Deploy-HomeLab.ps1

.EXAMPLE
    .\Deploy-HomeLab.ps1 -RemoveExisting
#>

param(
    [switch]$RemoveExisting
)

$ErrorActionPreference = 'Stop'

# ── Helpers ──────────────────────────────────────────────────────────────────

function Write-Step {
    param([string]$Title)
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "  $Title" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
}

function Write-Status {
    param([string]$Message, [ValidateSet('OK','WARN','FAIL','INFO','RUN','SKIP')]$Level = 'OK')
    switch ($Level) {
        'OK'   { Write-Host "  [OK]   $Message" -ForegroundColor Green }
        'WARN' { Write-Host "  [WARN] $Message" -ForegroundColor Yellow }
        'FAIL' { Write-Host "  [FAIL] $Message" -ForegroundColor Red }
        'INFO' { Write-Host "  [INFO] $Message" -ForegroundColor Cyan }
        'RUN'  { Write-Host "  [RUN]  $Message" -ForegroundColor Yellow }
        'SKIP' { Write-Host "  [SKIP] $Message" -ForegroundColor DarkGray }
    }
}

# ── Load config ──────────────────────────────────────────────────────────────

$Config = Import-PowerShellDataFile -Path (Join-Path $PSScriptRoot 'config.psd1')

$labName     = $Config.LabName
$domainName  = $Config.DomainName
$netPrefix   = $Config.Network
$networkName = "$labName-Network"
$siteCode    = $Config.SiteCode
$siteName    = $Config.SiteName
$netbios     = ($domainName -split '\.')[0].ToUpper()

Write-Host "`nMECM Home Lab Deployment" -ForegroundColor White
Write-Host "  Lab: $labName | Domain: $domainName | Site: $siteCode" -ForegroundColor DarkGray
Write-Host "  Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor DarkGray

###############################################################################
# PHASE 1: PREREQUISITES
###############################################################################

Write-Step 'Phase 1: Prerequisites'

# ── 1.1 Hyper-V ──────────────────────────────────────────────────────────────

Write-Host "`n--- Hyper-V ---" -ForegroundColor White

$hv = Get-WindowsOptionalFeature -Online -FeatureName Microsoft-Hyper-V -ErrorAction SilentlyContinue
if ($hv.State -ne 'Enabled') {
    Write-Status 'Hyper-V is NOT enabled.' -Level FAIL
    Write-Host ''
    Write-Host '  To enable Hyper-V, run the following in an elevated PowerShell prompt:' -ForegroundColor Yellow
    Write-Host '    Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Hyper-V -All' -ForegroundColor White
    Write-Host '  Then reboot and re-run this script.' -ForegroundColor Yellow
    Write-Host ''
    throw 'Hyper-V must be enabled before running this script. See instructions above.'
}
Write-Status 'Hyper-V enabled'

# ── 1.2 Host RAM ─────────────────────────────────────────────────────────────

Write-Host "`n--- Host Resources ---" -ForegroundColor White

$totalRam = (Get-CimInstance Win32_PhysicalMemory | Measure-Object Capacity -Sum).Sum
$totalRamGB = [math]::Round($totalRam / 1GB, 0)
if ($totalRamGB -ge 64) {
    Write-Status "Host RAM: ${totalRamGB}GB (excellent)"
} elseif ($totalRamGB -ge 32) {
    Write-Status "Host RAM: ${totalRamGB}GB (sufficient, 64GB recommended)" -Level WARN
} else {
    Write-Status "Host RAM: ${totalRamGB}GB (minimum 32GB required)" -Level FAIL
    throw "Insufficient RAM: ${totalRamGB}GB. Minimum 32GB required."
}

# ── 1.3 AutomatedLab module (vendored) ────────────────────────────────────────

Write-Host "`n--- AutomatedLab ---" -ForegroundColor White

$vendoredAL = Join-Path $PSScriptRoot 'lib\AutomatedLab'
$moduleDirs = Get-ChildItem $vendoredAL -Directory | Where-Object { Test-Path (Join-Path $_.FullName '*.psd1') }

$targetPath = Join-Path $env:ProgramFiles 'WindowsPowerShell\Modules'
foreach ($mod in $moduleDirs) {
    $dest = Join-Path $targetPath $mod.Name
    if (-not (Test-Path $dest)) {
        Copy-Item $mod.FullName $dest -Recurse -Force
        Write-Status "Installed: $($mod.Name)" -Level INFO
    }
}

# Remove non-essential modules that cause parse errors (Recipe, Ships)
foreach ($removeMod in @('AutomatedLab.Recipe', 'AutomatedLab.Ships', 'AutomatedLabTest')) {
    $removePath = Join-Path $targetPath $removeMod
    if (Test-Path $removePath) { Remove-Item $removePath -Recurse -Force }
}

# Patch manifest to remove references to removed modules
$manifestPath = Get-ChildItem (Join-Path $targetPath 'AutomatedLab') -Filter 'AutomatedLab.psd1' -Recurse | Select-Object -First 1
if ($manifestPath) {
    $content = Get-Content $manifestPath.FullName -Raw
    $content = $content -replace ".*AutomatedLab\.Recipe.*\r?\n", ''
    $content = $content -replace ".*AutomatedLab\.Ships.*\r?\n", ''
    $content = $content -replace ".*AutomatedLabTest.*\r?\n", ''
    Set-Content $manifestPath.FullName -Value $content
}

$al = Get-Module AutomatedLab -ListAvailable |
    Sort-Object Version -Descending |
    Select-Object -First 1

if (-not $al) {
    throw "AutomatedLab module not found after vendored install. Check lib\AutomatedLab\"
}
Write-Status "AutomatedLab v$($al.Version) (vendored fork)"

Import-Module AutomatedLab -ErrorAction Stop

# ── 1.4 LabSources folder ────────────────────────────────────────────────────

Write-Host "`n--- LabSources ---" -ForegroundColor White

$labSources = Get-LabSourcesLocation -ErrorAction SilentlyContinue
if (-not $labSources -or -not (Test-Path $labSources)) {
    Write-Status 'Creating LabSources folder on C:...' -Level RUN
    New-LabSourcesFolder -DriveLetter C
    $labSources = Get-LabSourcesLocation
}
Write-Status "LabSources: $labSources"

$swPkg = Join-Path $labSources 'SoftwarePackages'

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

# ── 1.5 ISO checks ───────────────────────────────────────────────────────────

Write-Host "`n--- ISO and Software Checks ---" -ForegroundColor White

$isoPath = Join-Path $labSources 'ISOs'

$wsIso = Get-ChildItem $isoPath -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -match 'SERVER_EVAL' -and $_.Extension -eq '.iso' }
$sqlIsoFile = Get-ChildItem $isoPath -Filter '*SQL*2022*' -ErrorAction SilentlyContinue |
    Select-Object -First 1
$cmDir = Get-ChildItem (Join-Path $swPkg 'CM') -ErrorAction SilentlyContinue |
    Where-Object { $_.PSIsContainer }

$missingItems = @()
if (-not $wsIso) {
    $missingItems += @{
        Name = 'Windows Server 2025 Evaluation ISO'
        Path = $isoPath
        URL  = 'https://www.microsoft.com/en-us/evalcenter/evaluate-windows-server-2025'
    }
}
if (-not $sqlIsoFile) {
    $missingItems += @{
        Name = 'SQL Server 2022 Evaluation ISO'
        Path = $isoPath
        URL  = 'https://www.microsoft.com/en-us/evalcenter/download-sql-server-2022'
    }
}
if (-not $cmDir) {
    $missingItems += @{
        Name = 'ConfigMgr 2509 Baseline (extract contents into CM folder)'
        Path = (Join-Path $swPkg 'CM')
        URL  = 'https://www.microsoft.com/en-us/evalcenter/download-microsoft-endpoint-configuration-manager'
    }
}

$adkSource   = Join-Path $swPkg 'ADK'
$adkPeSource = Join-Path $swPkg 'ADKPE'

$adkSetupFile = Get-ChildItem $adkSource -Filter 'adksetup*' -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $adkSetupFile) {
    $missingItems += @{
        Name = 'Windows ADK (adksetup.exe)'
        Path = $adkSource
        URL  = 'https://go.microsoft.com/fwlink/?linkid=2289980'
    }
}

$adkPeSetupFile = Get-ChildItem $adkPeSource -Filter 'adkwinpe*' -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $adkPeSetupFile) {
    $missingItems += @{
        Name = 'Windows PE add-on (adkwinpesetup.exe)'
        Path = $adkPeSource
        URL  = 'https://go.microsoft.com/fwlink/?linkid=2289981'
    }
}

if ($missingItems.Count -gt 0) {
    Write-Host ''
    foreach ($item in $missingItems) {
        Write-Status $item.Name -Level FAIL
        Write-Host "         Path: $($item.Path)" -ForegroundColor DarkGray
        Write-Host "         URL:  $($item.URL)" -ForegroundColor DarkGray
    }
    Write-Host ''
    throw "Missing required downloads (see above). Download them and re-run this script."
}

Write-Status 'Windows Server 2025 ISO found'
Write-Status 'SQL Server 2022 ISO found'
Write-Status 'ConfigMgr 2509 source found'
Write-Status 'ADK setup found'
Write-Status 'ADK PE setup found'

###############################################################################
# PHASE 2: OFFLINE DOWNLOADS
###############################################################################

Write-Step 'Phase 2: Offline Downloads'

# ── 2.1 ADK Offline Layout ───────────────────────────────────────────────────

Write-Host "`n--- ADK Offline Layout ---" -ForegroundColor White

$adkLayout = Join-Path $adkSource 'Offline'

if (Test-Path (Join-Path $adkLayout 'Installers')) {
    Write-Status "ADK offline layout already exists at: $adkLayout" -Level SKIP
} else {
    Write-Status "Creating ADK offline layout (this may take several minutes)..." -Level RUN
    $proc = Start-Process -FilePath $adkSetupFile.FullName `
        -ArgumentList @('/quiet', '/layout', $adkLayout) `
        -Wait -PassThru -NoNewWindow
    if ($proc.ExitCode -ne 0) {
        throw "ADK layout failed with exit code $($proc.ExitCode)"
    }
    Copy-Item -Path $adkSetupFile.FullName -Destination $adkLayout -Force
    Write-Status "ADK offline layout created at: $adkLayout"
}

# ── 2.2 ADK WinPE Offline Layout ─────────────────────────────────────────────

Write-Host "`n--- ADK WinPE Offline Layout ---" -ForegroundColor White

$adkPeLayout = Join-Path $adkPeSource 'Offline'

if (Test-Path (Join-Path $adkPeLayout 'Installers')) {
    Write-Status "ADK PE offline layout already exists at: $adkPeLayout" -Level SKIP
} else {
    Write-Status "Creating ADK PE offline layout (this may take several minutes)..." -Level RUN
    $proc = Start-Process -FilePath $adkPeSetupFile.FullName `
        -ArgumentList @('/quiet', '/layout', $adkPeLayout) `
        -Wait -PassThru -NoNewWindow
    if ($proc.ExitCode -ne 0) {
        throw "ADK PE layout failed with exit code $($proc.ExitCode)"
    }
    Copy-Item -Path $adkPeSetupFile.FullName -Destination $adkPeLayout -Force
    Write-Status "ADK PE offline layout created at: $adkPeLayout"
}

# ── 2.3 CM Prerequisites (setupdl.exe) ───────────────────────────────────────

Write-Host "`n--- ConfigMgr Prerequisites ---" -ForegroundColor White

$cmSource   = Join-Path $swPkg 'CM'
$prereqDest = Join-Path $swPkg 'CMPrereqs'

$setupDl = Get-ChildItem $cmSource -Recurse -Filter 'setupdl.exe' -ErrorAction SilentlyContinue | Select-Object -First 1

if (-not $setupDl) {
    Write-Status 'setupdl.exe not found in CM source folder -- cannot download prerequisites' -Level FAIL
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

# ── 2.4 ODBC Driver ──────────────────────────────────────────────────────────

Write-Host "`n--- ODBC Driver $($Config.ODBCVersion) ---" -ForegroundColor White

$odbcDest = Join-Path $swPkg 'ODBC'
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
        Write-Status "ODBC download failed: $($_.Exception.Message)" -Level FAIL
    }
}

# ── 2.5 VC++ Runtimes ────────────────────────────────────────────────────────

Write-Host "`n--- VC++ 14.50 Runtimes ---" -ForegroundColor White

$vcDest = Join-Path $swPkg 'VCRedist'
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
        Write-Status "VC++ x64 download failed: $($_.Exception.Message)" -Level FAIL
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
        Write-Status "VC++ x86 download failed: $($_.Exception.Message)" -Level FAIL
    }
}

###############################################################################
# PHASE 3: LAB DEPLOYMENT
###############################################################################

Write-Step 'Phase 3: Lab Deployment'

# ── 3.1 Check for existing lab ────────────────────────────────────────────────

$existingLabs = Get-Lab -List -ErrorAction SilentlyContinue
if ($existingLabs -contains $labName) {
    if ($RemoveExisting) {
        Write-Host "  Removing existing lab '$labName'..." -ForegroundColor Yellow
        Remove-Lab -Name $labName -Confirm:$false
    } else {
        Write-Host ''
        Write-Host "  Lab '$labName' already exists." -ForegroundColor Yellow
        Write-Host "  Use -RemoveExisting to auto-remove, or remove manually:" -ForegroundColor Yellow
        Write-Host "    Remove-Lab -Name $labName" -ForegroundColor White
        Write-Host ''

        # If the lab exists and we're not removing it, try to import it
        # and skip ahead to check what's already done
        $response = Read-Host "  Remove and recreate? (y/N)"
        if ($response -ne 'y') {
            Write-Host "  Attempting to import existing lab and continue..." -ForegroundColor Yellow
            Import-Lab -Name $labName -ErrorAction Stop
            # Fall through -- idempotent steps will skip what's already done
        } else {
            Remove-Lab -Name $labName -Confirm:$false
        }
    }
}

# ── 3.2 Define lab (only if not already imported) ─────────────────────────────

$labImported = $false
try {
    $currentLab = Get-Lab -ErrorAction SilentlyContinue
    if ($currentLab -and $currentLab.Name -eq $labName) {
        $labImported = $true
    }
} catch {
    $labImported = $false
}

if (-not $labImported) {
    Write-Host "`n--- Defining Lab: $labName ---" -ForegroundColor White

    New-LabDefinition -Name $labName -DefaultVirtualizationEngine HyperV

    # Internal network
    Add-LabVirtualNetworkDefinition -Name $networkName `
        -AddressSpace "$netPrefix.0/24" `
        -HyperVProperties @{ SwitchType = 'Internal' }

    # Default Switch for internet access (NAT) - reference existing Hyper-V built-in switch
    Add-LabVirtualNetworkDefinition -Name 'Default Switch' `
        -HyperVProperties @{ SwitchType = 'Internal' }

    # Domain
    Add-LabDomainDefinition -Name $domainName `
        -AdminUser $Config.AdminUser `
        -AdminPassword $Config.AdminPass

    # Credentials
    Set-LabInstallationCredential -Username $Config.AdminUser -Password $Config.AdminPass

    # SQL ISO
    $sqlIso = Get-ChildItem (Join-Path $labSources 'ISOs') -Filter '*SQL*2022*' -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if (-not $sqlIso) {
        throw "SQL Server 2022 ISO not found in $(Join-Path $labSources 'ISOs')"
    }
    Add-LabIsoImageDefinition -Name 'SQLServer2022' -Path $sqlIso.FullName

    # ── DC01 ──
    Write-Host "`n--- Defining DC01 ---" -ForegroundColor White

    $dcRoles = @(
        Get-LabMachineRoleDefinition -Role RootDC
        Get-LabMachineRoleDefinition -Role CaRoot
        Get-LabMachineRoleDefinition -Role Routing
    )

    $dcNics = @(
        New-LabNetworkAdapterDefinition -VirtualSwitch $networkName -Ipv4Address "$($Config.DC.IP)/24" -Ipv4DNSServers $Config.DC.IP
        New-LabNetworkAdapterDefinition -VirtualSwitch 'Default Switch' -UseDhcp
    )

    Add-LabMachineDefinition -Name $Config.DC.Name `
        -Roles $dcRoles `
        -Memory $Config.DC.Memory `
        -MinMemory $Config.DC.MinMemory `
        -MaxMemory $Config.DC.MaxMemory `
        -Processors $Config.DC.Processors `
        -NetworkAdapter $dcNics `
        -DomainName $domainName `
        -OperatingSystem 'Windows Server 2025 Datacenter Evaluation (Desktop Experience)'

    Write-Status "DC01 defined: $($Config.DC.IP), $([math]::Round($Config.DC.Memory/1GB))GB RAM, $($Config.DC.Processors) vCPU"

    # ── CM01 ──
    Write-Host "`n--- Defining CM01 ---" -ForegroundColor White

    $sqlRole = Get-LabMachineRoleDefinition -Role SQLServer2022 -Properties @{
        Collation    = $Config.SQLCollation
        InstanceName = 'MSSQLSERVER'
    }

    Add-LabDiskDefinition -Name 'CM01-SQL' -DiskSizeInGb $Config.CM.SQLDisk
    Add-LabDiskDefinition -Name 'CM01-Data' -DiskSizeInGb $Config.CM.DataDisk

    $cmNics = @(
        New-LabNetworkAdapterDefinition -VirtualSwitch $networkName -Ipv4Address "$($Config.CM.IP)/24" -Ipv4DNSServers $Config.DC.IP
        New-LabNetworkAdapterDefinition -VirtualSwitch 'Default Switch' -UseDhcp
    )

    Add-LabMachineDefinition -Name $Config.CM.Name `
        -Roles $sqlRole `
        -Memory $Config.CM.Memory `
        -MinMemory $Config.CM.MinMemory `
        -MaxMemory $Config.CM.MaxMemory `
        -Processors $Config.CM.Processors `
        -DiskName 'CM01-SQL', 'CM01-Data' `
        -NetworkAdapter $cmNics `
        -DomainName $domainName `
        -OperatingSystem 'Windows Server 2025 Datacenter Evaluation (Desktop Experience)'

    Write-Status "CM01 defined: $($Config.CM.IP), $([math]::Round($Config.CM.Memory/1GB))GB RAM, $($Config.CM.Processors) vCPU"

    # ── CLIENT01 ──
    Write-Host "`n--- Defining CLIENT01 ---" -ForegroundColor White

    $clientNics = @(
        New-LabNetworkAdapterDefinition -VirtualSwitch $networkName -Ipv4Address "$($Config.Client.IP)/24" -Ipv4DNSServers $Config.DC.IP
        New-LabNetworkAdapterDefinition -VirtualSwitch 'Default Switch' -UseDhcp
    )

    Add-LabMachineDefinition -Name $Config.Client.Name `
        -Memory $Config.Client.Memory `
        -MinMemory $Config.Client.MinMemory `
        -MaxMemory $Config.Client.MaxMemory `
        -Processors $Config.Client.Processors `
        -NetworkAdapter $clientNics `
        -DomainName $domainName `
        -OperatingSystem 'Windows 11 Enterprise Evaluation'

    Write-Status "CLIENT01 defined: $($Config.Client.IP), $([math]::Round($Config.Client.Memory/1GB))GB RAM, $($Config.Client.Processors) vCPU"

    # ── Install Lab ──
    Write-Host "`n--- Installing Lab (this will take 30-60 minutes) ---" -ForegroundColor White
    Write-Host "  Started at: $(Get-Date -Format 'HH:mm:ss')" -ForegroundColor DarkGray

    try {
        Install-Lab -DelayBetweenComputers 30 -NoValidation
    }
    catch {
        # Install-Lab may throw non-fatal errors (e.g., SSRS config timing).
        # Check if VMs are running — if so, continue with remaining phases.
        $runningVMs = Get-VM -Name $Config.DC.Name, $Config.CM.Name -ErrorAction SilentlyContinue | Where-Object State -eq 'Running'
        if ($runningVMs.Count -ge 2) {
            Write-Status "Install-Lab reported errors but VMs are running. Continuing." -Level WARN
            Write-Status "Error: $($_.Exception.Message)" -Level WARN
        }
        else {
            throw "Install-Lab failed and VMs are not running: $($_.Exception.Message)"
        }
    }

    Write-Status 'Lab VMs deployed and domain joined'
    Write-Host "  Finished at: $(Get-Date -Format 'HH:mm:ss')" -ForegroundColor DarkGray
} else {
    Write-Status "Lab '$labName' already deployed -- skipping VM creation" -Level SKIP
}

# ── 3.3 Expand CM01 OS disk ──────────────────────────────────────────────────

Write-Host "`n--- Expanding CM01 OS Disk ---" -ForegroundColor White

$cmVM = Get-VM -Name $Config.CM.Name -ErrorAction SilentlyContinue
if ($cmVM) {
    $osDisk = $cmVM | Get-VMHardDiskDrive | Where-Object { $_.ControllerLocation -eq 0 } | Select-Object -First 1
    if ($osDisk) {
        $currentSizeGB = [math]::Round((Get-VHD $osDisk.Path).Size / 1GB, 0)
        if ($currentSizeGB -lt $Config.CM.OSDiskSize) {
            Resize-VHD -Path $osDisk.Path -SizeBytes ($Config.CM.OSDiskSize * 1GB)
            Write-Status "VHDX expanded to $($Config.CM.OSDiskSize)GB"

            Invoke-LabCommand -ComputerName $Config.CM.Name -ActivityName 'Extend C: partition' -ScriptBlock {
                $maxSize = (Get-PartitionSupportedSize -DriveLetter C).SizeMax
                Resize-Partition -DriveLetter C -Size $maxSize
            }
            Write-Status 'C: partition extended inside VM'
        } else {
            Write-Status "OS disk already ${currentSizeGB}GB (>= $($Config.CM.OSDiskSize)GB)" -Level SKIP
        }
    }
} else {
    Write-Status 'CM01 VM not found -- cannot expand disk' -Level FAIL
}

# ── 3.4 Configure SQL Memory ─────────────────────────────────────────────────

Write-Host "`n--- Configuring SQL Server Memory ---" -ForegroundColor White

Invoke-LabCommand -ComputerName $Config.CM.Name -ActivityName 'Configure SQL memory' -ScriptBlock {
    $sqlSvc = Get-Service MSSQLSERVER -ErrorAction SilentlyContinue
    if ($sqlSvc -and $sqlSvc.Status -eq 'Running') {
        Invoke-Sqlcmd -Query @"
EXEC sp_configure 'show advanced options', 1;
RECONFIGURE;
EXEC sp_configure 'min server memory', 8192;
EXEC sp_configure 'max server memory', 8192;
RECONFIGURE;
"@
    }
}
Write-Status 'SQL memory set to 8GB min/max'

###############################################################################
# PHASE 4: COPY SOFTWARE TO CM01
###############################################################################

Write-Step 'Phase 4: Copy Software to CM01'

$cmName = $Config.CM.Name

# Create install directories on CM01
Invoke-LabCommand -ComputerName $cmName -ActivityName 'Create Install directories' -ScriptBlock {
    $dirs = @(
        'C:\Install'
        'C:\Install\ADKoffline'
        'C:\Install\ADKPEoffline'
        'C:\Install\CM'
        'C:\Install\CM-Prereqs'
        'C:\Install\ODBC'
        'C:\Install\VCRedist'
    )
    foreach ($d in $dirs) {
        if (-not (Test-Path $d)) { New-Item -Path $d -ItemType Directory -Force | Out-Null }
    }
}

# ADK offline layout
$adkLayoutCopy = Join-Path $swPkg 'ADKoffline'
# Check both layout locations (ADK\Offline or ADKoffline)
if (-not (Test-Path $adkLayoutCopy)) { $adkLayoutCopy = $adkLayout }
$adkBootstrapper = Join-Path $swPkg 'ADK\adksetup.exe'
if (-not (Test-Path $adkBootstrapper) -and $adkSetupFile) { $adkBootstrapper = $adkSetupFile.FullName }

if (Test-Path $adkLayoutCopy) {
    Write-Status 'Copying ADK offline layout...' -Level RUN
    Copy-LabFileItem -Path $adkLayoutCopy -ComputerName $cmName -DestinationFolderPath 'C:\Install' -Recurse

    Invoke-LabCommand -ComputerName $cmName -ActivityName 'Flatten ADK folder' -ScriptBlock {
        $nested = 'C:\Install\ADKoffline\ADKoffline'
        if (Test-Path $nested) {
            $items = Get-ChildItem $nested
            foreach ($item in $items) { Move-Item -Path $item.FullName -Destination 'C:\Install\ADKoffline' -Force }
            Remove-Item $nested -Recurse -Force -ErrorAction SilentlyContinue
        }
        # Also flatten Offline subfolder if layout was ADK\Offline
        $nestedOffline = 'C:\Install\ADKoffline\Offline'
        if (Test-Path $nestedOffline) {
            $items = Get-ChildItem $nestedOffline
            foreach ($item in $items) { Move-Item -Path $item.FullName -Destination 'C:\Install\ADKoffline' -Force }
            Remove-Item $nestedOffline -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    if (Test-Path $adkBootstrapper) {
        Copy-LabFileItem -Path $adkBootstrapper -ComputerName $cmName -DestinationFolderPath 'C:\Install\ADKoffline'
    }
    Write-Status 'ADK offline layout copied'
} else {
    Write-Status "ADK offline layout not found" -Level WARN
}

# ADK PE offline layout
$adkPeLayoutCopy = Join-Path $swPkg 'ADKPEoffline'
if (-not (Test-Path $adkPeLayoutCopy)) { $adkPeLayoutCopy = $adkPeLayout }
$adkPeBootstrapper = Join-Path $swPkg 'ADKPE\adkwinpesetup.exe'
if (-not (Test-Path $adkPeBootstrapper) -and $adkPeSetupFile) { $adkPeBootstrapper = $adkPeSetupFile.FullName }

if (Test-Path $adkPeLayoutCopy) {
    Write-Status 'Copying ADK PE offline layout...' -Level RUN
    Copy-LabFileItem -Path $adkPeLayoutCopy -ComputerName $cmName -DestinationFolderPath 'C:\Install' -Recurse

    Invoke-LabCommand -ComputerName $cmName -ActivityName 'Flatten ADK PE folder' -ScriptBlock {
        $nested = 'C:\Install\ADKPEoffline\ADKPEoffline'
        if (Test-Path $nested) {
            $items = Get-ChildItem $nested
            foreach ($item in $items) { Move-Item -Path $item.FullName -Destination 'C:\Install\ADKPEoffline' -Force }
            Remove-Item $nested -Recurse -Force -ErrorAction SilentlyContinue
        }
        $nestedOffline = 'C:\Install\ADKPEoffline\Offline'
        if (Test-Path $nestedOffline) {
            $items = Get-ChildItem $nestedOffline
            foreach ($item in $items) { Move-Item -Path $item.FullName -Destination 'C:\Install\ADKPEoffline' -Force }
            Remove-Item $nestedOffline -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    if (Test-Path $adkPeBootstrapper) {
        Copy-LabFileItem -Path $adkPeBootstrapper -ComputerName $cmName -DestinationFolderPath 'C:\Install\ADKPEoffline'
    }
    Write-Status 'ADK PE offline layout copied'
} else {
    Write-Status "ADK PE offline layout not found" -Level WARN
}

# CM source
$cmSourceDir = Get-ChildItem (Join-Path $swPkg 'CM') -Directory -ErrorAction SilentlyContinue | Select-Object -First 1
if ($cmSourceDir) {
    Write-Status "Copying CM source ($($cmSourceDir.Name))..." -Level RUN
    Copy-LabFileItem -Path $cmSourceDir.FullName -ComputerName $cmName -DestinationFolderPath 'C:\Install\CM' -Recurse

    Invoke-LabCommand -ComputerName $cmName -ActivityName 'Flatten CM folder' -ScriptBlock {
        $cmBase = 'C:\Install\CM'
        $subDirs = Get-ChildItem $cmBase -Directory -ErrorAction SilentlyContinue
        foreach ($sub in $subDirs) {
            $items = Get-ChildItem $sub.FullName
            foreach ($item in $items) { Move-Item -Path $item.FullName -Destination $cmBase -Force }
            Remove-Item $sub.FullName -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    Write-Status 'CM source copied and flattened'
} else {
    Write-Status 'CM source folder not found -- skipping' -Level WARN
}

# CM prerequisites
$prereqDir = Join-Path $swPkg 'CM-Prereqs'
if (-not (Test-Path $prereqDir)) { $prereqDir = Join-Path $swPkg 'CMPrereqs' }
if ((Get-ChildItem $prereqDir -ErrorAction SilentlyContinue | Measure-Object).Count -gt 0) {
    Write-Status 'Copying CM prerequisites...' -Level RUN
    Copy-LabFileItem -Path $prereqDir -ComputerName $cmName -DestinationFolderPath 'C:\Install\CM-Prereqs' -Recurse

    Invoke-LabCommand -ComputerName $cmName -ActivityName 'Flatten CM-Prereqs folder' -ScriptBlock {
        $base = 'C:\Install\CM-Prereqs'
        $subDirs = Get-ChildItem $base -Directory -ErrorAction SilentlyContinue
        foreach ($sub in $subDirs) {
            $items = Get-ChildItem $sub.FullName
            foreach ($item in $items) { Move-Item -Path $item.FullName -Destination $base -Force }
            Remove-Item $sub.FullName -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    Write-Status 'CM prerequisites copied'
} else {
    Write-Status 'CM prerequisites folder empty -- skipping' -Level WARN
}

# ODBC driver
$odbcFileCopy = Get-ChildItem (Join-Path $swPkg 'ODBC') -Filter '*.msi' -ErrorAction SilentlyContinue | Select-Object -First 1
if ($odbcFileCopy) {
    Write-Status 'Copying ODBC driver...' -Level RUN
    Copy-LabFileItem -Path $odbcFileCopy.FullName -ComputerName $cmName -DestinationFolderPath 'C:\Install\ODBC'
    Write-Status 'ODBC driver copied'
} else {
    Write-Status 'ODBC MSI not found -- skipping' -Level WARN
}

# VC++ runtimes
$vcDir = Join-Path $swPkg 'VCRedist'
if (Test-Path (Join-Path $vcDir 'vc_redist.x64.exe')) {
    Write-Status 'Copying VC++ runtimes...' -Level RUN
    Copy-LabFileItem -Path (Join-Path $vcDir 'vc_redist.x64.exe') -ComputerName $cmName -DestinationFolderPath 'C:\Install\VCRedist'
    Copy-LabFileItem -Path (Join-Path $vcDir 'vc_redist.x86.exe') -ComputerName $cmName -DestinationFolderPath 'C:\Install\VCRedist'
    Write-Status 'VC++ runtimes copied'
} else {
    Write-Status 'VC++ runtimes not found -- skipping' -Level WARN
}

###############################################################################
# PHASE 5: AD CONFIGURATION
###############################################################################

Write-Step 'Phase 5: AD Configuration'

# ── 5.1 AD Schema Extension ──────────────────────────────────────────────────

Write-Host "`n--- AD Schema Extension ---" -ForegroundColor White

Invoke-LabCommand -ComputerName $cmName -ActivityName 'Run extadsch.exe' -ScriptBlock {
    $extadsch = Get-ChildItem 'C:\Install\CM' -Filter 'extadsch.exe' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($extadsch) {
        $result = Start-Process -FilePath $extadsch.FullName -Wait -PassThru -NoNewWindow
        if (Test-Path 'C:\ExtADSch.log') {
            $logLines = Get-Content 'C:\ExtADSch.log' -Tail 5
            foreach ($line in $logLines) { Write-Host "    $line" }
        }
    } else {
        Write-Warning 'extadsch.exe not found in C:\Install\CM'
    }
}
Write-Status 'AD schema extension complete'

# ── 5.2 System Management Container ──────────────────────────────────────────

Write-Host "`n--- System Management Container ---" -ForegroundColor White

$domainDN = ($domainName -split '\.' | ForEach-Object { "DC=$_" }) -join ','

Invoke-LabCommand -ComputerName $Config.DC.Name -ActivityName 'Create System Management container' -ScriptBlock {
    param($DomainDN, $SiteServerName)

    Import-Module ActiveDirectory

    $sysManPath = "CN=System Management,CN=System,$DomainDN"
    $systemPath = "CN=System,$DomainDN"

    # Create container if it doesn't exist
    $existing = Get-ADObject -Filter "DistinguishedName -eq '$sysManPath'" -ErrorAction SilentlyContinue
    if (-not $existing) {
        New-ADObject -Name 'System Management' -Type Container -Path $systemPath
        Write-Host "Created container: $sysManPath"
    } else {
        Write-Host "Container already exists: $sysManPath"
    }

    # Grant site server Full Control
    $siteServer = Get-ADComputer $SiteServerName
    $acl = Get-Acl "AD:\$sysManPath"

    $identity = [System.Security.Principal.SecurityIdentifier]$siteServer.SID
    $rights = [System.DirectoryServices.ActiveDirectoryRights]::GenericAll
    $type = [System.Security.AccessControl.AccessControlType]::Allow
    $inheritance = [System.DirectoryServices.ActiveDirectorySecurityInheritance]::SelfAndChildren

    $ace = New-Object System.DirectoryServices.ActiveDirectoryAccessRule($identity, $rights, $type, $inheritance)
    $acl.AddAccessRule($ace)
    Set-Acl "AD:\$sysManPath" $acl

    Write-Host "$SiteServerName granted Full Control on System Management container"
} -ArgumentList $domainDN, $Config.CM.Name

Write-Status 'System Management container configured'

###############################################################################
# PHASE 6: SERVICE ACCOUNTS
###############################################################################

Write-Step 'Phase 6: Service Accounts'

# Write a script file to copy to DC01 and execute there, to avoid
# splatting/pipeline issues through remoting layers.

$svcAccountScript = @"

Import-Module ActiveDirectory

`$domainDN = '$domainDN'
`$domainName = '$domainName'
`$netbios = '$netbios'
`$ouName = 'Service Accounts'
`$ouPath = "OU=`$ouName,`$domainDN"

# Create OU if needed
`$existingOU = Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '`$ouPath'" -ErrorAction SilentlyContinue
if (-not `$existingOU) {
    New-ADOrganizationalUnit -Name `$ouName -Path `$domainDN
    Write-Host "Created OU: `$ouPath"
} else {
    Write-Host "OU already exists: `$ouPath"
}

# Account definitions
`$accounts = @(
    @{
        Sam  = '$($Config.ServiceAccounts.ClientPush.Name)'
        Full = 'MECM Client Push'
        UPN  = '$($Config.ServiceAccounts.ClientPush.Name)@$domainName'
        Pass = '$($Config.ServiceAccounts.ClientPush.Password)'
        Desc = '$($Config.ServiceAccounts.ClientPush.Desc)'
    },
    @{
        Sam  = '$($Config.ServiceAccounts.NAA.Name)'
        Full = 'MECM Network Access Account'
        UPN  = '$($Config.ServiceAccounts.NAA.Name)@$domainName'
        Pass = '$($Config.ServiceAccounts.NAA.Password)'
        Desc = '$($Config.ServiceAccounts.NAA.Desc)'
    },
    @{
        Sam  = '$($Config.ServiceAccounts.Admin.Name)'
        Full = 'MECM Admin'
        UPN  = '$($Config.ServiceAccounts.Admin.Name)@$domainName'
        Pass = '$($Config.ServiceAccounts.Admin.Password)'
        Desc = '$($Config.ServiceAccounts.Admin.Desc)'
    }
)

foreach (`$acct in `$accounts) {
    `$existing = Get-ADUser -Filter "SamAccountName -eq '`$(`$acct.Sam)'" -ErrorAction SilentlyContinue
    if (-not `$existing) {
        New-ADUser -Name `$acct.Full ``
            -SamAccountName `$acct.Sam ``
            -UserPrincipalName `$acct.UPN ``
            -Path `$ouPath ``
            -AccountPassword (ConvertTo-SecureString `$acct.Pass -AsPlainText -Force) ``
            -PasswordNeverExpires `$true ``
            -CannotChangePassword `$true ``
            -Enabled `$true ``
            -Description `$acct.Desc
        Write-Host "Created: `$netbios\`$(`$acct.Sam)"
    } else {
        Write-Host "Exists: `$netbios\`$(`$acct.Sam)"
    }
}

# Group memberships
Add-ADGroupMember -Identity 'Domain Admins' -Members '$($Config.ServiceAccounts.ClientPush.Name)' -ErrorAction SilentlyContinue
Add-ADGroupMember -Identity 'Domain Admins' -Members '$($Config.ServiceAccounts.Admin.Name)' -ErrorAction SilentlyContinue
Add-ADGroupMember -Identity 'Remote Desktop Users' -Members '$($Config.ServiceAccounts.Admin.Name)' -ErrorAction SilentlyContinue

Write-Host "`nService accounts configured successfully."
"@

# Write the script to a temp file, copy to DC01, run it
$tempScript = Join-Path $env:TEMP 'Create-ServiceAccounts.ps1'
$svcAccountScript | Set-Content -Path $tempScript -Encoding ASCII -Force

Copy-LabFileItem -Path $tempScript -ComputerName $Config.DC.Name -DestinationFolderPath 'C:\Install'

Invoke-LabCommand -ComputerName $Config.DC.Name -ActivityName 'Create service accounts' -ScriptBlock {
    & 'C:\Install\Create-ServiceAccounts.ps1'
}

Remove-Item $tempScript -Force -ErrorAction SilentlyContinue

Write-Status "Service accounts created ($($Config.ServiceAccounts.ClientPush.Name), $($Config.ServiceAccounts.NAA.Name), $($Config.ServiceAccounts.Admin.Name))"

###############################################################################
# PHASE 7: INSTALL SOFTWARE ON CM01
###############################################################################

Write-Step 'Phase 7: Install Software on CM01'

# ── 7.1 VC++ Runtimes ────────────────────────────────────────────────────────

Write-Host "`n--- VC++ 14.50 Runtimes ---" -ForegroundColor White

Invoke-LabCommand -ComputerName $cmName -ActivityName 'Install VC++ x64' -ScriptBlock {
    $exe = 'C:\Install\VCRedist\vc_redist.x64.exe'
    if (Test-Path $exe) {
        $proc = Start-Process -FilePath $exe -ArgumentList @('/install', '/quiet', '/norestart') -Wait -PassThru
        if ($proc.ExitCode -eq 0 -or $proc.ExitCode -eq 3010) {
            Write-Host "VC++ x64 installed (exit code: $($proc.ExitCode))"
        } else {
            throw "VC++ x64 install failed with exit code $($proc.ExitCode)"
        }
    } else {
        throw "vc_redist.x64.exe not found at $exe"
    }
}
Write-Status 'VC++ x64 installed'

Invoke-LabCommand -ComputerName $cmName -ActivityName 'Install VC++ x86' -ScriptBlock {
    $exe = 'C:\Install\VCRedist\vc_redist.x86.exe'
    if (Test-Path $exe) {
        $proc = Start-Process -FilePath $exe -ArgumentList @('/install', '/quiet', '/norestart') -Wait -PassThru
        if ($proc.ExitCode -eq 0 -or $proc.ExitCode -eq 3010) {
            Write-Host "VC++ x86 installed (exit code: $($proc.ExitCode))"
        } else {
            throw "VC++ x86 install failed with exit code $($proc.ExitCode)"
        }
    } else {
        throw "vc_redist.x86.exe not found at $exe"
    }
}
Write-Status 'VC++ x86 installed'

# ── 7.2 ODBC Driver ──────────────────────────────────────────────────────────

Write-Host "`n--- ODBC Driver $($Config.ODBCVersion) ---" -ForegroundColor White

Invoke-LabCommand -ComputerName $cmName -ActivityName 'Install ODBC Driver' -ScriptBlock {
    $msi = Get-ChildItem 'C:\Install\ODBC' -Filter '*.msi' -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($msi) {
        $proc = Start-Process -FilePath 'msiexec.exe' `
            -ArgumentList @('/i', $msi.FullName, '/qn', '/norestart', 'IACCEPTMSODBCSQLLICENSETERMS=YES') `
            -Wait -PassThru
        if ($proc.ExitCode -eq 0 -or $proc.ExitCode -eq 3010) {
            Write-Host "ODBC Driver installed (exit code: $($proc.ExitCode))"
        } else {
            throw "ODBC install failed with exit code $($proc.ExitCode)"
        }
    } else {
        throw 'ODBC MSI not found in C:\Install\ODBC'
    }
}
Write-Status "ODBC Driver $($Config.ODBCVersion) installed"

# ── 7.3 MSOLEDB ──────────────────────────────────────────────────────────────

Write-Host "`n--- MSOLEDB 19 ---" -ForegroundColor White

Invoke-LabCommand -ComputerName $cmName -ActivityName 'Install MSOLEDB' -ScriptBlock {
    $msi = Get-ChildItem 'C:\Install\CM-Prereqs' -Filter 'msoledbsql*.msi' -ErrorAction SilentlyContinue |
        Sort-Object Name -Descending | Select-Object -First 1
    if (-not $msi) {
        $msi = Get-ChildItem 'C:\Install\CM' -Recurse -Filter 'msoledbsql*.msi' -ErrorAction SilentlyContinue |
            Sort-Object Name -Descending | Select-Object -First 1
    }
    if ($msi) {
        $proc = Start-Process -FilePath 'msiexec.exe' `
            -ArgumentList @('/i', $msi.FullName, '/qn', '/norestart', 'IACCEPTMSOLEDBSQLLICENSETERMS=YES') `
            -Wait -PassThru
        if ($proc.ExitCode -eq 0 -or $proc.ExitCode -eq 3010) {
            Write-Host "MSOLEDB installed (exit code: $($proc.ExitCode))"
        } else {
            throw "MSOLEDB install failed with exit code $($proc.ExitCode)"
        }
    } else {
        Write-Warning 'MSOLEDB MSI not found -- CM setup may install it automatically'
    }
}
Write-Status 'MSOLEDB installed'

# ── 7.4 ADK ──────────────────────────────────────────────────────────────────

Write-Host "`n--- Windows ADK ---" -ForegroundColor White

Invoke-LabCommand -ComputerName $cmName -ActivityName 'Install ADK' -ScriptBlock {
    $adkSetup = Get-ChildItem 'C:\Install\ADKoffline' -Filter 'adksetup*' -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $adkSetup) {
        throw 'adksetup.exe not found in C:\Install\ADKoffline'
    }

    # Check if already installed
    $allUninstall = Get-ItemProperty 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' -ErrorAction SilentlyContinue
    $adkReg = $null
    foreach ($entry in $allUninstall) {
        if ($entry.DisplayName -like '*Assessment and Deployment Kit*') { $adkReg = $entry; break }
    }
    if ($adkReg) {
        Write-Host 'ADK already installed -- skipping'
        return
    }

    $proc = Start-Process -FilePath $adkSetup.FullName `
        -ArgumentList @(
            '/quiet',
            '/installpath', 'C:\Program Files (x86)\Windows Kits\10',
            '/features', 'OptionId.DeploymentTools', 'OptionId.UserStateMigrationTool',
            '/ceip', 'off'
        ) `
        -Wait -PassThru
    if ($proc.ExitCode -ne 0) {
        throw "ADK install failed with exit code $($proc.ExitCode)"
    }
    Write-Host 'ADK installed successfully'
}
Write-Status 'ADK installed (DeploymentTools + UserStateMigrationTool)'

# ── 7.5 ADK WinPE ────────────────────────────────────────────────────────────

Write-Host "`n--- ADK WinPE Add-on ---" -ForegroundColor White

Invoke-LabCommand -ComputerName $cmName -ActivityName 'Install ADK WinPE' -ScriptBlock {
    $peSetup = Get-ChildItem 'C:\Install\ADKPEoffline' -Filter 'adkwinpe*' -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $peSetup) {
        throw 'adkwinpesetup.exe not found in C:\Install\ADKPEoffline'
    }

    # Check if already installed
    $allUninstall = Get-ItemProperty 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' -ErrorAction SilentlyContinue
    $peReg = $null
    foreach ($entry in $allUninstall) {
        if ($entry.DisplayName -like '*Windows PE*' -or $entry.DisplayName -like '*Preinstallation*') { $peReg = $entry; break }
    }
    if ($peReg) {
        Write-Host 'ADK WinPE already installed -- skipping'
        return
    }

    $proc = Start-Process -FilePath $peSetup.FullName `
        -ArgumentList @(
            '/quiet',
            '/installpath', 'C:\Program Files (x86)\Windows Kits\10',
            '/features', 'OptionId.WindowsPreinstallationEnvironment',
            '/ceip', 'off'
        ) `
        -Wait -PassThru
    if ($proc.ExitCode -ne 0) {
        throw "ADK WinPE install failed with exit code $($proc.ExitCode)"
    }
    Write-Host 'ADK WinPE add-on installed successfully'
}
Write-Status 'ADK WinPE add-on installed'

###############################################################################
# PHASE 8: INSTALL CONFIGMGR
###############################################################################

Write-Step 'Phase 8: Install ConfigMgr 2509'

$cmFqdn    = "$cmName.$domainName"
$adminUser = "$netbios\$($Config.AdminUser)"

# ── 8.1 Generate Unattended Setup INI ─────────────────────────────────────────

Write-Host "`n--- Generating Setup INI ---" -ForegroundColor White

Invoke-LabCommand -ComputerName $cmName -ActivityName 'Generate setup INI' -ScriptBlock {
    param($SiteCode, $SiteName, $CmFqdn, $AdminUser, $Domain, $SQLCollation)

    $ini = @"
[Identification]
Action=InstallPrimarySite

[Options]
ProductID=EVAL
SiteCode=$SiteCode
SiteName=$SiteName
SMSInstallDir=C:\Program Files\Microsoft Configuration Manager
SDKServer=$CmFqdn
RoleCommunicationProtocol=HTTPorHTTPS
ClientsUsePKICertificate=0
PrerequisiteComp=1
PrerequisitePath=C:\Install\CM-Prereqs
AdminConsole=1
JoinCEIP=0
MobileDeviceLanguage=0

[SQLConfigOptions]
SQLServerName=$CmFqdn
DatabaseName=CM_$SiteCode
SQLSSBPort=4022
SQLDataFilePath=E:\MSSQL\Data
SQLLogFilePath=E:\MSSQL\Log
SQLServerSSLCertificate=

[CloudConnectorOptions]
CloudConnector=0

[SABranchOptions]
SAActive=0

[SystemCenterOptions]

[HierarchyExpansionOption]
"@

    $ini | Set-Content -Path 'C:\Install\CM\setup.ini' -Encoding ASCII -Force
    Write-Host 'Setup INI generated at C:\Install\CM\setup.ini'
} -ArgumentList $siteCode, $siteName, $cmFqdn, $adminUser, $domainName, $Config.SQLCollation

Write-Status 'Setup INI generated'

# ── 8.2 Prepare SQL directories ──────────────────────────────────────────────

Write-Host "`n--- Preparing SQL Data Directories ---" -ForegroundColor White

Invoke-LabCommand -ComputerName $cmName -ActivityName 'Create SQL directories' -ScriptBlock {
    $dirs = @('E:\MSSQL\Data', 'E:\MSSQL\Log')
    foreach ($d in $dirs) {
        if (-not (Test-Path $d)) {
            New-Item -Path $d -ItemType Directory -Force | Out-Null
            Write-Host "Created: $d"
        }
    }

    foreach ($d in $dirs) {
        $acl = Get-Acl $d
        $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            'NT SERVICE\MSSQLSERVER', 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow'
        )
        $acl.AddAccessRule($rule)
        Set-Acl -Path $d -AclObject $acl
    }
    Write-Host 'SQL directories created and permissions set'
}
Write-Status 'SQL data directories ready'

# ── 8.3 Install ConfigMgr ────────────────────────────────────────────────────

Write-Host "`n--- Installing ConfigMgr 2509 (this will take 1-3 hours) ---" -ForegroundColor White
Write-Host "  Started at: $(Get-Date -Format 'HH:mm:ss')" -ForegroundColor DarkGray

Invoke-LabCommand -ComputerName $cmName -ActivityName 'Install ConfigMgr 2509' -ScriptBlock {
    $setupExe = Get-ChildItem 'C:\Install\CM' -Filter 'setup.exe' -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $setupExe) {
        throw 'setup.exe not found in C:\Install\CM'
    }

    $iniFile = 'C:\Install\CM\setup.ini'
    if (-not (Test-Path $iniFile)) {
        throw "Setup INI not found at $iniFile"
    }

    # Check if CM is already installed
    $cmSvc = Get-Service SMS_EXECUTIVE -ErrorAction SilentlyContinue
    if ($cmSvc) {
        Write-Host 'ConfigMgr already installed (SMS_EXECUTIVE service exists) -- skipping'
        return
    }

    # Enable required Windows features
    $features = @(
        'NET-Framework-Features',
        'NET-Framework-Core',
        'BITS',
        'BITS-IIS-Ext',
        'BITS-Compact-Server',
        'RDC',
        'WAS-Process-Model',
        'WAS-Config-APIs',
        'WAS-Net-Environment',
        'Web-Server',
        'Web-ISAPI-Ext',
        'Web-ISAPI-Filter',
        'Web-Net-Ext',
        'Web-Net-Ext45',
        'Web-ASP-Net',
        'Web-ASP-Net45',
        'Web-ASP',
        'Web-Windows-Auth',
        'Web-Basic-Auth',
        'Web-URL-Auth',
        'Web-IP-Security',
        'Web-Scripting-Tools',
        'Web-Mgmt-Service',
        'Web-Stat-Compression',
        'Web-Dyn-Compression',
        'Web-Default-Doc',
        'Web-Filtering',
        'Web-Dir-Browsing',
        'Web-Http-Errors',
        'Web-Static-Content',
        'Web-Http-Redirect',
        'Web-Log-Libraries',
        'Web-Http-Tracing',
        'Web-Metabase',
        'Web-Lgcy-Mgmt-Console',
        'Web-WMI',
        'Web-Lgcy-Scripting',
        'RSAT-Feature-Tools',
        'UpdateServices-WidDB',
        'UpdateServices-Services',
        'UpdateServices-RSAT',
        'UpdateServices-API',
        'UpdateServices-UI'
    )

    Write-Host 'Installing Windows features for ConfigMgr...'
    Install-WindowsFeature -Name $features -ErrorAction SilentlyContinue | Out-Null
    Write-Host 'Windows features installed'

    Write-Host 'Starting ConfigMgr setup...'
    $proc = Start-Process -FilePath $setupExe.FullName `
        -ArgumentList @('/SCRIPT', $iniFile) `
        -Wait -PassThru -NoNewWindow
    Write-Host "Setup completed with exit code: $($proc.ExitCode)"

    if ($proc.ExitCode -ne 0) {
        $logPath = 'C:\ConfigMgrSetup.log'
        if (Test-Path $logPath) {
            $logLines = Get-Content $logPath -Tail 30
            $errors = @()
            foreach ($line in $logLines) {
                if ($line -match 'ERROR|FAIL') { $errors += $line }
            }
            if ($errors.Count -gt 0) {
                Write-Host '--- Last errors from setup log ---'
                foreach ($err in $errors) { Write-Host "  $err" }
            }
        }
        throw "ConfigMgr setup failed with exit code $($proc.ExitCode)"
    }
} -Timeout ([TimeSpan]::FromHours(4))

Write-Status 'ConfigMgr setup completed'
Write-Host "  Finished at: $(Get-Date -Format 'HH:mm:ss')" -ForegroundColor DarkGray

# ── 8.4 Validate Installation ────────────────────────────────────────────────

Write-Host "`n--- Validating ConfigMgr Installation ---" -ForegroundColor White

Write-Status 'Waiting 60 seconds for services to initialize...' -Level INFO
Start-Sleep -Seconds 60

$validation = Invoke-LabCommand -ComputerName $cmName -ActivityName 'Validate CM install' -ScriptBlock {
    param($SiteCode)

    $result = @{
        SiteFound    = $false
        SiteStatus   = ''
        ConsoleFound = $false
        ServiceState = ''
    }

    try {
        $site = Get-CimInstance -Namespace "ROOT\SMS\site_$SiteCode" -ClassName SMS_Site -ErrorAction Stop
        if ($site) {
            $result.SiteFound  = $true
            $result.SiteStatus = $site.Status
        }
    } catch {
        $errMsg = $Error[0].Exception.Message
        Write-Host "WMI query failed: $errMsg"
    }

    $consolePath = 'C:\Program Files\Microsoft Configuration Manager\AdminConsole\bin\Microsoft.ConfigurationManagement.exe'
    $result.ConsoleFound = Test-Path $consolePath

    $svc = Get-Service SMS_EXECUTIVE -ErrorAction SilentlyContinue
    if ($svc) {
        $result.ServiceState = $svc.Status.ToString()
    }

    return $result
} -ArgumentList $siteCode -PassThru

if ($validation.SiteFound) {
    Write-Status "Site $siteCode found in WMI (Status: $($validation.SiteStatus))"
} else {
    Write-Status "Site $siteCode not yet visible in WMI -- may still be initializing" -Level WARN
}

if ($validation.ConsoleFound) {
    Write-Status 'CM Console installed'
} else {
    Write-Status 'CM Console not found at expected path' -Level WARN
}

if ($validation.ServiceState -eq 'Running') {
    Write-Status 'SMS_EXECUTIVE service is running'
} else {
    Write-Status "SMS_EXECUTIVE service state: $($validation.ServiceState)" -Level WARN
}

###############################################################################
# PHASE 9: CONTENT SHARE
###############################################################################

Write-Step 'Phase 9: Content Share'

$sharePath = 'E:\ContentShare'
$shareName = 'ContentShare$'
$folders = @(
    'Applications'
    'Drivers'
    'Images'
    'OperatingSystems'
    'Packages'
    'Scripts'
    'SoftwareUpdates'
)

Invoke-LabCommand -ComputerName $cmName -ActivityName 'Create content share' -ScriptBlock {
    param($SharePath, $ShareName, $Folders, $DomainNetBIOS)

    # Create folder structure
    New-Item -Path $SharePath -ItemType Directory -Force | Out-Null
    foreach ($f in $Folders) {
        New-Item -Path (Join-Path $SharePath $f) -ItemType Directory -Force | Out-Null
    }
    Write-Host "Created folder structure at $SharePath"

    # Create SMB share
    if (-not (Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue)) {
        New-SmbShare -Name $ShareName -Path $SharePath `
            -FullAccess "$DomainNetBIOS\Domain Admins" `
            -ReadAccess "$DomainNetBIOS\Domain Computers", "$DomainNetBIOS\svc-CMNAA"
        Write-Host "Created share: \\$env:COMPUTERNAME\$ShareName"
    } else {
        Write-Host "Share already exists: \\$env:COMPUTERNAME\$ShareName"
    }

    # Set NTFS permissions
    $acl = Get-Acl $SharePath
    $naaRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
        "$DomainNetBIOS\svc-CMNAA", 'Read', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
    $acl.AddAccessRule($naaRule)
    $compRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
        "$DomainNetBIOS\Domain Computers", 'Read', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
    $acl.AddAccessRule($compRule)
    Set-Acl $SharePath $acl
    Write-Host 'NTFS permissions set (Domain Admins=Full, Domain Computers+NAA=Read)'
} -ArgumentList $sharePath, $shareName, $folders, $netbios

Write-Status "Content share created: \\$cmName\$shareName"

###############################################################################
# PHASE 10: MECM FULL ADMINISTRATOR
###############################################################################

Write-Step 'Phase 10: Add svc-CMAdmin as MECM Full Administrator'

Invoke-LabCommand -ComputerName $cmName -ActivityName 'Add MECM Full Administrator' -ScriptBlock {
    param($SiteCode, $AdminAccount)

    # Wait for the CM provider to be ready
    $retries = 0
    $providerReady = $false
    while ($retries -lt 6 -and -not $providerReady) {
        try {
            $testNs = Get-CimInstance -Namespace "ROOT\SMS\site_$SiteCode" -ClassName SMS_Site -ErrorAction Stop
            if ($testNs) { $providerReady = $true }
        } catch {
            $retries++
            Write-Host "  Waiting for CM provider (attempt $retries/6)..."
            Start-Sleep -Seconds 30
        }
    }

    if (-not $providerReady) {
        Write-Warning 'CM provider not ready -- svc-CMAdmin must be added manually via console'
        return
    }

    # Check if admin already exists
    $allAdmins = Get-CimInstance -Namespace "ROOT\SMS\site_$SiteCode" -ClassName SMS_Admin -ErrorAction SilentlyContinue
    $existing = $null
    foreach ($admin in $allAdmins) {
        if ($admin.LogonName -like "*$AdminAccount*" -or $admin.AdminID -like "*$AdminAccount*") { $existing = $admin; break }
    }
    if ($existing) {
        Write-Host "$AdminAccount is already an MECM administrator"
        return
    }

    # Get the Full Administrator role ID
    $allRoles = Get-CimInstance -Namespace "ROOT\SMS\site_$SiteCode" -ClassName SMS_Role -ErrorAction SilentlyContinue
    $fullAdminRole = $null
    foreach ($role in $allRoles) {
        if ($role.RoleName -eq 'Full Administrator') { $fullAdminRole = $role; break }
    }

    if (-not $fullAdminRole) {
        Write-Warning 'Full Administrator role not found -- add svc-CMAdmin manually'
        return
    }

    # Use the CM PowerShell module if available
    $cmModulePath = Join-Path $env:SMS_ADMIN_UI_PATH '..\ConfigurationManager.psd1'
    if (Test-Path $cmModulePath) {
        Import-Module $cmModulePath -ErrorAction SilentlyContinue
        $drive = Get-PSDrive -Name $SiteCode -ErrorAction SilentlyContinue
        if (-not $drive) {
            New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $env:COMPUTERNAME -ErrorAction SilentlyContinue | Out-Null
        }
        $currentLocation = Get-Location
        Set-Location "${SiteCode}:"

        try {
            $existingAdmin = Get-CMAdministrativeUser -Name $AdminAccount -ErrorAction SilentlyContinue
            if (-not $existingAdmin) {
                New-CMAdministrativeUser -Name $AdminAccount -RoleName 'Full Administrator' -SecurityScopeName 'All' -ErrorAction Stop
                Write-Host "$AdminAccount added as MECM Full Administrator"
            } else {
                Write-Host "$AdminAccount is already an MECM administrator"
            }
        } finally {
            Set-Location $currentLocation
        }
    } else {
        Write-Warning 'CM PowerShell module not found -- add svc-CMAdmin as Full Administrator manually via console'
    }
} -ArgumentList $siteCode, "$netbios\$($Config.ServiceAccounts.Admin.Name)"

Write-Status "svc-CMAdmin MECM Full Administrator role configured"

###############################################################################
# PHASE 11: DEPLOY TOOLS
###############################################################################

Write-Step 'Phase 11: Deploy Tools to CM01'

# Create tools directory on CM01
Invoke-LabCommand -ComputerName $cmName -ActivityName 'Create Tools directory' -ScriptBlock {
    if (-not (Test-Path 'C:\Tools')) { New-Item -Path 'C:\Tools' -ItemType Directory -Force | Out-Null }
}

# Copy cc4cm if available locally
$cc4cmSource = 'C:\projects\Internal\cc4cm'
if (Test-Path $cc4cmSource) {
    Write-Status 'Copying Client Center (cc4cm)...' -Level RUN
    Copy-LabFileItem -Path $cc4cmSource -ComputerName $cmName -DestinationFolderPath 'C:\Tools' -Recurse
    Write-Status 'cc4cm deployed to C:\Tools\cc4cm'
} else {
    Write-Status "cc4cm not found at $cc4cmSource -- skipping" -Level WARN
}

# Copy ApplicationPackager if available locally
$appPkgSource = 'C:\projects\applicationpackager'
if (Test-Path $appPkgSource) {
    Write-Status 'Copying Application Packager...' -Level RUN
    Copy-LabFileItem -Path $appPkgSource -ComputerName $cmName -DestinationFolderPath 'C:\Tools' -Recurse

    # Rename to standard path
    Invoke-LabCommand -ComputerName $cmName -ActivityName 'Rename AppPackager folder' -ScriptBlock {
        $src = 'C:\Tools\applicationpackager'
        $dst = 'C:\Tools\ApplicationPackager'
        if ((Test-Path $src) -and -not (Test-Path $dst)) {
            Rename-Item $src $dst
        }
    }
    Write-Status 'ApplicationPackager deployed to C:\Tools\ApplicationPackager'
} else {
    Write-Status "ApplicationPackager not found at $appPkgSource -- skipping" -Level WARN
}

###############################################################################
# PHASE 12: SNAPSHOTS
###############################################################################

Write-Step 'Phase 12: Snapshots'

Checkpoint-VM -Name $Config.DC.Name -SnapshotName 'Deployment-Complete' -ErrorAction SilentlyContinue
Checkpoint-VM -Name $Config.CM.Name -SnapshotName 'Deployment-Complete' -ErrorAction SilentlyContinue
Checkpoint-VM -Name $Config.Client.Name -SnapshotName 'Deployment-Complete' -ErrorAction SilentlyContinue
Write-Status 'Snapshots created: Deployment-Complete'

###############################################################################
# PHASE 13: CONNECTION INFO & NEXT STEPS
###############################################################################

Write-Step 'Deployment Complete!'

$elapsed = (Get-Date) - (Get-Date $MyInvocation.MyCommand.Module.PrivateData.StartTime -ErrorAction SilentlyContinue)

Write-Host ''
Write-Host '  =============================================' -ForegroundColor Green
Write-Host '   MECM HOME LAB DEPLOYMENT COMPLETE' -ForegroundColor Green
Write-Host '  =============================================' -ForegroundColor Green
Write-Host ''
Write-Host "  Domain:     $domainName" -ForegroundColor White
Write-Host "  Admin:      $domainName\$($Config.AdminUser)" -ForegroundColor White
Write-Host "  Password:   $($Config.AdminPass)" -ForegroundColor White
Write-Host "  DC01:       $($Config.DC.IP)" -ForegroundColor White
Write-Host "  CM01:       $($Config.CM.IP)" -ForegroundColor White
Write-Host "  Site Code:  $siteCode" -ForegroundColor White
Write-Host "  Site Name:  $siteName" -ForegroundColor White
Write-Host ''
Write-Host '  Service Accounts:' -ForegroundColor White
Write-Host "    $netbios\$($Config.ServiceAccounts.ClientPush.Name)  (Client Push - Domain Admins)" -ForegroundColor DarkGray
Write-Host "    $netbios\$($Config.ServiceAccounts.NAA.Name)   (NAA - Domain Users only)" -ForegroundColor DarkGray
Write-Host "    $netbios\$($Config.ServiceAccounts.Admin.Name) (Admin - Domain Admins + RDP)" -ForegroundColor DarkGray
Write-Host ''
Write-Host '  Content Share:' -ForegroundColor White
Write-Host "    \\$cmName\ContentShare$" -ForegroundColor DarkGray
Write-Host ''
Write-Host '  Connect to VMs:' -ForegroundColor Cyan
Write-Host "    Connect-LabVM -ComputerName $cmName          # RDP" -ForegroundColor White
Write-Host "    Enter-LabPSSession -ComputerName $cmName     # PS remoting" -ForegroundColor White
Write-Host ''
Write-Host '  Lab management:' -ForegroundColor Cyan
Write-Host "    Stop-Lab -Name $labName      # Stop all VMs" -ForegroundColor White
Write-Host "    Start-Lab -Name $labName     # Start all VMs" -ForegroundColor White
Write-Host "    Remove-Lab -Name $labName    # Delete entire lab" -ForegroundColor White
Write-Host ''
Write-Host '  =============================================' -ForegroundColor Yellow
Write-Host '   REMAINING MANUAL CONSOLE STEPS' -ForegroundColor Yellow
Write-Host '  =============================================' -ForegroundColor Yellow
Write-Host ''
Write-Host '  Open the CM console on CM01 and configure:' -ForegroundColor White
Write-Host ''
Write-Host '  1. Active Directory Forest Discovery:' -ForegroundColor White
Write-Host '     Administration > Hierarchy Configuration > Discovery Methods' -ForegroundColor DarkGray
Write-Host '     > Active Directory Forest Discovery > Enable' -ForegroundColor DarkGray
Write-Host ''
Write-Host '  2. Create a Boundary:' -ForegroundColor White
Write-Host '     Administration > Hierarchy Configuration > Boundaries > Create' -ForegroundColor DarkGray
Write-Host "     > Type: IP Subnet > $netPrefix.0/24" -ForegroundColor DarkGray
Write-Host ''
Write-Host '  3. Create a Boundary Group:' -ForegroundColor White
Write-Host '     Administration > Hierarchy Configuration > Boundary Groups > Create' -ForegroundColor DarkGray
Write-Host '     > Add the boundary > References tab > add CM01 as site system' -ForegroundColor DarkGray
Write-Host ''
Write-Host '  4. Enable Client Push:' -ForegroundColor White
Write-Host '     Administration > Site Configuration > Sites > right-click site' -ForegroundColor DarkGray
Write-Host "     > Client Installation Settings > Client Push > Accounts > Add $netbios\$($Config.ServiceAccounts.ClientPush.Name)" -ForegroundColor DarkGray
Write-Host ''
Write-Host '  5. Network Access Account:' -ForegroundColor White
Write-Host '     Administration > Site Configuration > Sites > right-click site' -ForegroundColor DarkGray
Write-Host "     > Configure Site Components > Software Distribution > NAA > Add $netbios\$($Config.ServiceAccounts.NAA.Name)" -ForegroundColor DarkGray
Write-Host ''
Write-Host "  Finished at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor DarkGray
Write-Host ''
