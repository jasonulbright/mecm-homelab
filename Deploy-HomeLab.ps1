#Requires -RunAsAdministrator
#Requires -Version 5.1
<#
.SYNOPSIS
    Deploys a complete MECM home lab: DC01 + CM01 + CLIENT01 with AD, CA, SQL, ConfigMgr 2509, and a Windows 11 workstation.

.DESCRIPTION
    Single-script deployment that performs all steps:
      1.  Check prerequisites (Hyper-V, host RAM)
      2.  Install vendored AutomatedLab modules
      3.  Create LabSources folder structure, verify ISOs, create ADK offline layouts
      4.  Define and deploy the lab (DC01 + CM01 + CLIENT01) via AutomatedLab
          - AutomatedLab handles: SQL, VC++, ODBC, MSOLEDB, ADK, AD schema, CM install
      5.  Expand CM01 OS disk, configure SQL memory
      6.  Create service accounts (svc-CMPush, svc-CMNAA, svc-CMAdmin)
      7.  Create content share on CM01
      8.  Add svc-CMAdmin as MECM Full Administrator
      9.  Deploy tools (cc4cm + AppPackager) to CM01
     10.  Create snapshots
     11.  Print connection info and remaining manual console steps

    CM01 has both SQLServer2022 and ConfigurationManager roles, letting
    AutomatedLab do the heavy lifting for prerequisite installs and CM setup.
    Only CM01 has a second NIC on the Default Switch for internet access.

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
    # Always overwrite to ensure vendored fixes are applied
    if (Test-Path $dest) { Remove-Item $dest -Recurse -Force }
    Copy-Item $mod.FullName $dest -Recurse -Force
    Write-Status "Installed: $($mod.Name)" -Level INFO
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

# Force our vendored VC++ URLs (PSFConfig -Initialize won't overwrite existing values)
Set-PSFConfig -Module 'AutomatedLab' -Name cppredist64_2017 -Value 'https://aka.ms/vs/18/release/vc_redist.x64.exe'
Set-PSFConfig -Module 'AutomatedLab' -Name cppredist32_2017 -Value 'https://aka.ms/vs/18/release/vc_redist.x86.exe'
Set-PSFConfig -Module 'AutomatedLab' -Name cppredist64_2015 -Value 'https://aka.ms/vs/18/release/vc_redist.x64.exe'
Set-PSFConfig -Module 'AutomatedLab' -Name cppredist32_2015 -Value 'https://aka.ms/vs/18/release/vc_redist.x86.exe'

# Remove stale cached old VC++ runtimes (PSGallery version downloaded 2015/2017 binaries
# which conflict with latest 14.50). Get-LabInternetFile won't re-download if file exists.
$labSrc = Get-LabSourcesLocation -ErrorAction SilentlyContinue
if ($labSrc) {
    foreach ($stale in @('vcredist_x64_2015.exe','vcredist_x86_2015.exe','vcredist_x64_2017.exe','vcredist_x86_2017.exe')) {
        $stalePath = Join-Path $labSrc "SoftwarePackages\$stale"
        if (Test-Path $stalePath) { Remove-Item $stalePath -Force }
    }
}

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
    $cmRole = Get-LabMachineRoleDefinition -Role ConfigurationManager -Properties @{
        Version  = '2509'
        Branch   = 'CB'
        SiteCode = $siteCode
        SiteName = $siteName
        Roles    = 'Management Point,Distribution Point'
    }

    Add-LabDiskDefinition -Name 'CM01-SQL' -DiskSizeInGb $Config.CM.SQLDisk
    Add-LabDiskDefinition -Name 'CM01-Data' -DiskSizeInGb $Config.CM.DataDisk

    $cmNics = @(
        New-LabNetworkAdapterDefinition -VirtualSwitch $networkName -Ipv4Address "$($Config.CM.IP)/24" -Ipv4DNSServers $Config.DC.IP
        New-LabNetworkAdapterDefinition -VirtualSwitch 'Default Switch' -UseDhcp
    )

    Add-LabMachineDefinition -Name $Config.CM.Name `
        -Roles $sqlRole, $cmRole `
        -Memory $Config.CM.Memory `
        -MinMemory $Config.CM.MinMemory `
        -MaxMemory $Config.CM.MaxMemory `
        -Processors $Config.CM.Processors `
        -DiskName 'CM01-SQL', 'CM01-Data' `
        -NetworkAdapter $cmNics `
        -DomainName $domainName `
        -OperatingSystem 'Windows Server 2025 Datacenter Evaluation (Desktop Experience)'

    Write-Status "CM01 defined: $($Config.CM.IP), $([math]::Round($Config.CM.Memory/1GB))GB RAM, $($Config.CM.Processors) vCPU, SQL+CM roles"

    # ── CLIENT01 (created but not deployed until after DC+CM are done) ──
    Write-Host "`n--- Defining CLIENT01 (deferred deployment) ---" -ForegroundColor White

    $clientNics = @(
        New-LabNetworkAdapterDefinition -VirtualSwitch $networkName -Ipv4Address "$($Config.Client.IP)/24" -Ipv4DNSServers $Config.DC.IP
    )
    Add-LabMachineDefinition -Name $Config.Client.Name `
        -Memory $Config.Client.Memory `
        -MinMemory $Config.Client.MinMemory `
        -MaxMemory $Config.Client.MaxMemory `
        -Processors $Config.Client.Processors `
        -NetworkAdapter $clientNics `
        -DomainName $domainName `
        -OperatingSystem 'Windows 11 Enterprise Evaluation' `
        -SkipDeployment
    Write-Status "CLIENT01 defined (will deploy after DC+CM)" -Level INFO

    # ── Install Lab (DC01 + CM01 — CLIENT01 skipped via SkipDeployment) ──
    # AutomatedLab handles: AD, CA, SQL, VC++, ODBC, MSOLEDB, ADK, AD schema, CM install
    Write-Host "`n--- Installing Lab - DC01 + CM01 with SQL + ConfigMgr (this will take 2-4 hours) ---" -ForegroundColor White
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

    Write-Status 'DC01 + CM01 deployed (AD, SQL, ConfigMgr installed by AutomatedLab)'
    Write-Host "  Finished at: $(Get-Date -Format 'HH:mm:ss')" -ForegroundColor DarkGray

    # ── CLIENT01: Deploy now that DC+CM are up (avoids RAM contention during AD/SQL) ──
    Write-Host "`n--- Deploying CLIENT01 ---" -ForegroundColor White
    $clientVM = Get-LabVM -ComputerName $Config.Client.Name -ErrorAction SilentlyContinue
    if ($clientVM -and $clientVM.SkipDeployment) {
        $clientVM.SkipDeployment = $false
        Install-Lab -NoValidation
        Write-Status "CLIENT01 deployed: $($Config.Client.IP)"
    } elseif (-not $clientVM) {
        Write-Status 'CLIENT01 not in lab definition — skipping' -Level WARN
    } else {
        Write-Status 'CLIENT01 already deployed' -Level SKIP
    }

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

$cmName = $Config.CM.Name
$domainDN = ($domainName -split '\.' | ForEach-Object { "DC=$_" }) -join ','

###############################################################################
# PHASE 4: SERVICE ACCOUNTS
###############################################################################

Write-Step 'Phase 4: Service Accounts'

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
# PHASE 5: CONTENT SHARE
###############################################################################

Write-Step 'Phase 5: Content Share'

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
# PHASE 6: MECM FULL ADMINISTRATOR
###############################################################################

Write-Step 'Phase 6: Add svc-CMAdmin as MECM Full Administrator'

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
# PHASE 7: DEPLOY TOOLS
###############################################################################

Write-Step 'Phase 7: Deploy Tools to CM01'

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
# PHASE 8: SNAPSHOTS
###############################################################################

Write-Step 'Phase 8: Snapshots'

Checkpoint-VM -Name $Config.DC.Name -SnapshotName 'Deployment-Complete' -ErrorAction SilentlyContinue
Checkpoint-VM -Name $Config.CM.Name -SnapshotName 'Deployment-Complete' -ErrorAction SilentlyContinue
Checkpoint-VM -Name $Config.Client.Name -SnapshotName 'Deployment-Complete' -ErrorAction SilentlyContinue
Write-Status 'Snapshots created: Deployment-Complete'

###############################################################################
# PHASE 9: CONNECTION INFO & NEXT STEPS
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
