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
      5.  Expand CM01 OS disk (if below configured size)
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

.PARAMETER LogFile
    Path to a clean log file. Captures all output without ANSI color codes,
    progress bars, or carriage returns. Suitable for documentation or review.

.EXAMPLE
    .\Deploy-HomeLab.ps1

.EXAMPLE
    .\Deploy-HomeLab.ps1 -RemoveExisting

.EXAMPLE
    .\Deploy-HomeLab.ps1 -LogFile C:\temp\deploy.log
#>

param(
    [switch]$RemoveExisting,
    [string]$LogFile
)

$ErrorActionPreference = 'Stop'

# ── Helpers ──────────────────────────────────────────────────────────────────

# Log file setup -- writes clean text (no ANSI, no progress bars) alongside console output.
# Intercepts Write-Host so ALL output is logged without modifying 115+ call sites.
if ($LogFile) {
    $script:logStream = [System.IO.StreamWriter]::new($LogFile, $false, [System.Text.Encoding]::UTF8)
    $script:logStream.AutoFlush = $true

    # Rename the real Write-Host and replace it with a logging wrapper
    if (-not (Get-Command Write-HostOriginal -ErrorAction SilentlyContinue)) {
        $null = New-Item -Path function: -Name 'script:Write-HostOriginal' -Value (Get-Command Write-Host).ScriptBlock -ErrorAction SilentlyContinue
    }

    function Write-Host {
        param(
            [Parameter(Position=0)][object]$Object,
            [switch]$NoNewline,
            [object]$ForegroundColor,
            [object]$BackgroundColor,
            [string]$Separator = ' '
        )
        # Write to console (with color)
        $params = @{}
        if ($PSBoundParameters.ContainsKey('Object'))          { $params.Object = $Object }
        if ($PSBoundParameters.ContainsKey('NoNewline'))        { $params.NoNewline = $NoNewline }
        if ($PSBoundParameters.ContainsKey('ForegroundColor'))  { $params.ForegroundColor = $ForegroundColor }
        if ($PSBoundParameters.ContainsKey('BackgroundColor'))  { $params.BackgroundColor = $BackgroundColor }
        if ($PSBoundParameters.ContainsKey('Separator'))        { $params.Separator = $Separator }
        Microsoft.PowerShell.Utility\Write-Host @params

        # Write to log (clean text, no color codes)
        if ($script:logStream) {
            $text = if ($null -eq $Object) { '' } else { [string]$Object }
            if ($NoNewline) { $script:logStream.Write($text) } else { $script:logStream.WriteLine($text) }
        }
    }
}

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
$script:deployStartTime = Get-Date
Write-Host "  Started: $($script:deployStartTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor DarkGray

# ── Password security check ──────────────────────────────────────────────────
$defaultPasswords = @('P@ssw0rd!', 'P@ssw0rd!Push1', 'P@ssw0rd!NAA1', 'P@ssw0rd!Admin1')
$allPasswords = @(
    $Config.AdminPass
    $Config.ServiceAccounts.ClientPush.Password
    $Config.ServiceAccounts.NAA.Password
    $Config.ServiceAccounts.Admin.Password
)
$usingDefaults = ($allPasswords | Where-Object { $_ -in $defaultPasswords }).Count
if ($usingDefaults -gt 0) {
    Write-Host ''
    Write-Host '  !! WARNING: DEFAULT PASSWORDS DETECTED !!' -ForegroundColor Red
    Write-Host "  $usingDefaults of $($allPasswords.Count) passwords in config.psd1 are still defaults." -ForegroundColor Red
    Write-Host '  These passwords are published in source control.' -ForegroundColor Red
    Write-Host '  Change them in config.psd1 before deploying to any network.' -ForegroundColor Red
    Write-Host ''
}

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

# Unload any previously imported AL modules (DLLs lock files)
Get-Module AutomatedLab* | Remove-Module -Force -ErrorAction SilentlyContinue
Get-Module PSFramework | Remove-Module -Force -ErrorAction SilentlyContinue

$targetPath = Join-Path $env:ProgramFiles 'WindowsPowerShell\Modules'
foreach ($mod in $moduleDirs) {
    $dest = Join-Path $targetPath $mod.Name
    # Always overwrite to ensure vendored fixes are applied
    if (Test-Path $dest) {
        try {
            Remove-Item $dest -Recurse -Force
        } catch {
            # DLL may still be locked by another process -- overwrite in place
            Write-Status "Could not remove $($mod.Name), overwriting in place" -Level WARN
        }
    }
    Copy-Item $mod.FullName $dest -Recurse -Force
    Write-Status "Installed: $($mod.Name)" -Level INFO
}

# Remove non-essential modules that cause parse errors (Recipe, Ships)
foreach ($removeMod in @('AutomatedLab.Recipe', 'AutomatedLab.Ships', 'AutomatedLabTest')) {
    $removePath = Join-Path $targetPath $removeMod
    if (Test-Path $removePath) { Remove-Item $removePath -Recurse -Force -ErrorAction SilentlyContinue }
}

$al = Get-Module AutomatedLab -ListAvailable |
    Sort-Object Version -Descending |
    Select-Object -First 1

if (-not $al) {
    throw "AutomatedLab module not found after vendored install. Check lib\AutomatedLab\"
}
Write-Status "AutomatedLab v$($al.Version) (vendored fork)"

Import-Module AutomatedLab -ErrorAction Stop

# Force our vendored config overrides (PSFConfig -Initialize won't overwrite persisted values)
Set-PSFConfig -Module 'AutomatedLab' -Name DisableVersionCheck -Value $true
Set-PSFConfig -Module 'AutomatedLab' -Name cppredist64_2017 -Value 'https://aka.ms/vs/18/release/vc_redist.x64.exe'
Set-PSFConfig -Module 'AutomatedLab' -Name cppredist32_2017 -Value 'https://aka.ms/vs/18/release/vc_redist.x86.exe'
Set-PSFConfig -Module 'AutomatedLab' -Name cppredist64_2015 -Value 'https://aka.ms/vs/18/release/vc_redist.x64.exe'
Set-PSFConfig -Module 'AutomatedLab' -Name cppredist32_2015 -Value 'https://aka.ms/vs/18/release/vc_redist.x86.exe'

# Remove stale cached old VC++ runtimes (PSGallery version downloaded version-specific binaries
# which conflict with latest). Get-LabInternetFile won't re-download if file exists.
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

# ── 1.6 Resolve OS edition names from ISOs ──────────────────────────────────

Write-Host "`n--- OS Edition Detection ---" -ForegroundColor White

$availableOS = Get-LabAvailableOperatingSystem -Path $isoPath -NoDisplay -ErrorAction SilentlyContinue
if (-not $availableOS) {
    throw "No operating systems found in ISOs at $isoPath. Verify your ISO files are valid."
}

$serverOS = ($availableOS | Where-Object OperatingSystemName -like $Config.ServerOSFilter |
    Sort-Object Version -Descending | Select-Object -First 1).OperatingSystemName
$clientOS = ($availableOS | Where-Object OperatingSystemName -like $Config.ClientOSFilter |
    Sort-Object Version -Descending | Select-Object -First 1).OperatingSystemName

if (-not $serverOS) {
    Write-Status "No server OS matching '$($Config.ServerOSFilter)'" -Level FAIL
    Write-Host '  Available:' -ForegroundColor DarkGray
    $availableOS | Where-Object OperatingSystemName -like '*Server*' | ForEach-Object {
        Write-Host "    $($_.OperatingSystemName)" -ForegroundColor DarkGray
    }
    throw "No server OS found matching filter '$($Config.ServerOSFilter)'. Update ServerOSFilter in config.psd1."
}
if (-not $clientOS) {
    Write-Status "No client OS matching '$($Config.ClientOSFilter)'" -Level FAIL
    Write-Host '  Available:' -ForegroundColor DarkGray
    $availableOS | Where-Object OperatingSystemName -like '*Windows 1*' | ForEach-Object {
        Write-Host "    $($_.OperatingSystemName)" -ForegroundColor DarkGray
    }
    throw "No client OS found matching filter '$($Config.ClientOSFilter)'. Update ClientOSFilter in config.psd1."
}

Write-Status "Server OS: $serverOS"
Write-Status "Client OS: $clientOS"

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
        # Lab exists -- import it and continue (idempotent phases will skip what's done)
        Write-Status "Lab '$labName' already exists -- importing and continuing" -Level INFO
        Import-Lab -Name $labName -ErrorAction Stop
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
        # Routing role removed — requires 2 NICs. CM01 has its own Default Switch for internet.
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
        -OperatingSystem $serverOS

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
        -OperatingSystem $serverOS

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
        -OperatingSystem $clientOS `
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

    Write-Status 'DC01 + CM01 deployed'
    Write-Host "  Finished at: $(Get-Date -Format 'HH:mm:ss')" -ForegroundColor DarkGray

    # Validate CM — SMS_EXECUTIVE may need time to start after install+reboot
    Write-Status 'Waiting for SMS_EXECUTIVE service...' -Level RUN
    $cmRunning = $false
    for ($attempt = 1; $attempt -le 12; $attempt++) {
        $result = Invoke-LabCommand -ComputerName $Config.CM.Name -ActivityName 'Check SMS_EXECUTIVE service' -PassThru -ScriptBlock {
            (Get-Service SMS_EXECUTIVE -ErrorAction SilentlyContinue).Status -eq 'Running'
        }
        # Invoke-LabCommand -PassThru may return PSObject wrapping the bool — cast explicitly
        $cmRunning = [bool]($result | Select-Object -First 1)
        if ($cmRunning) { break }
        Write-Host "  Attempt $attempt/12 -- waiting 30s..." -ForegroundColor DarkGray
        Start-Sleep -Seconds 30
    }
    if ($cmRunning) {
        Write-Status 'ConfigMgr verified: SMS_EXECUTIVE running' -Level OK
    } else {
        Write-Status 'SMS_EXECUTIVE not running after 6 minutes. Check ConfigMgrSetup.log on CM01.' -Level WARN
    }

    # ── CLIENT01: Deploy now that DC+CM are up (avoids RAM contention during AD/SQL) ──
    Write-Host "`n--- Deploying CLIENT01 ---" -ForegroundColor White
    $clientVM = Get-LabVM -ComputerName $Config.Client.Name -ErrorAction SilentlyContinue
    if ($clientVM -and $clientVM.SkipDeployment) {
        $clientVM.SkipDeployment = $false
        try {
            Install-Lab -NoValidation
        } catch {
            # Install-Lab re-iterates all roles (SQL, CM) for idempotency.
            # CM update validation may fail on fresh labs (no updates synced yet).
            $runningClient = Get-VM -Name $Config.Client.Name -ErrorAction SilentlyContinue | Where-Object State -eq 'Running'
            if ($runningClient) {
                Write-Status "Install-Lab reported errors but CLIENT01 is running. Continuing." -Level WARN
            } else {
                Write-Status "CLIENT01 deployment failed: $($_.Exception.Message)" -Level FAIL
            }
        }
        Write-Status "CLIENT01 deployed: $($Config.Client.IP)"
    } elseif (-not $clientVM) {
        Write-Status 'CLIENT01 not in lab definition — skipping' -Level WARN
    } else {
        Write-Status 'CLIENT01 already deployed' -Level SKIP
    }

} else {
    Write-Status "Lab '$labName' already deployed -- skipping VM creation" -Level SKIP

    # Verify CM is running on existing lab
    Import-Lab -Name $labName -NoValidation
    $result = Invoke-LabCommand -ComputerName $Config.CM.Name -ActivityName 'Check SMS_EXECUTIVE service' -PassThru -ScriptBlock {
        (Get-Service SMS_EXECUTIVE -ErrorAction SilentlyContinue).Status -eq 'Running'
    }
    $cmRunning = [bool]($result | Select-Object -First 1)
    if (-not $cmRunning) {
        Write-Status 'SMS_EXECUTIVE not running -- CM may still be initializing or was never installed. Check CM01 manually.' -Level WARN
    } else {
        Write-Status 'ConfigMgr verified: SMS_EXECUTIVE running' -Level OK
    }
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

$cmName = $Config.CM.Name
$domainDN = ($domainName -split '\.' | ForEach-Object { "DC=$_" }) -join ','

###############################################################################
# PHASE 4: SERVICE ACCOUNTS
###############################################################################

Write-Step 'Phase 4: Service Accounts'

try {

# Build accounts array on the host, pass via -ArgumentList (no string interpolation of passwords)
$svcAccounts = @(
    @{
        Sam  = $Config.ServiceAccounts.ClientPush.Name
        Full = 'MECM Client Push'
        Pass = $Config.ServiceAccounts.ClientPush.Password
        Desc = $Config.ServiceAccounts.ClientPush.Desc
    },
    @{
        Sam  = $Config.ServiceAccounts.NAA.Name
        Full = 'MECM Network Access Account'
        Pass = $Config.ServiceAccounts.NAA.Password
        Desc = $Config.ServiceAccounts.NAA.Desc
    },
    @{
        Sam  = $Config.ServiceAccounts.Admin.Name
        Full = 'MECM Admin'
        Pass = $Config.ServiceAccounts.Admin.Password
        Desc = $Config.ServiceAccounts.Admin.Desc
    }
)

# Group memberships: hashtable of group name -> array of account SAMs to add
$groupMemberships = @{
    'Domain Admins'        = @($Config.ServiceAccounts.ClientPush.Name, $Config.ServiceAccounts.Admin.Name)
    'Remote Desktop Users' = @($Config.ServiceAccounts.Admin.Name)
}

Invoke-LabCommand -ComputerName $Config.DC.Name -ActivityName 'Create service accounts' -ScriptBlock {
    param($DomainDN, $DomainName, $NetBIOS, $Accounts, $GroupMemberships)

    Import-Module ActiveDirectory

    $ouName = 'Service Accounts'
    $ouPath = "OU=$ouName,$DomainDN"

    # Create OU if needed
    $existingOU = Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '$ouPath'" -ErrorAction SilentlyContinue
    if (-not $existingOU) {
        New-ADOrganizationalUnit -Name $ouName -Path $DomainDN
        Write-Host "Created OU: $ouPath"
    } else {
        Write-Host "OU already exists: $ouPath"
    }

    # Create accounts
    foreach ($acct in $Accounts) {
        $existing = Get-ADUser -Filter "SamAccountName -eq '$($acct.Sam)'" -ErrorAction SilentlyContinue
        if (-not $existing) {
            New-ADUser -Name $acct.Full `
                -SamAccountName $acct.Sam `
                -UserPrincipalName "$($acct.Sam)@$DomainName" `
                -Path $ouPath `
                -AccountPassword (ConvertTo-SecureString $acct.Pass -AsPlainText -Force) `
                -PasswordNeverExpires $true `
                -CannotChangePassword $true `
                -Enabled $true `
                -Description $acct.Desc
            Write-Host "Created: $NetBIOS\$($acct.Sam)"
        } else {
            Write-Host "Exists: $NetBIOS\$($acct.Sam)"
        }
    }

    # Group memberships
    foreach ($group in $GroupMemberships.Keys) {
        foreach ($member in $GroupMemberships[$group]) {
            Add-ADGroupMember -Identity $group -Members $member -ErrorAction SilentlyContinue
        }
    }

    Write-Host 'Service accounts configured successfully.'
} -ArgumentList $domainDN, $domainName, $netbios, $svcAccounts, $groupMemberships

Write-Status "Service accounts created ($($Config.ServiceAccounts.ClientPush.Name), $($Config.ServiceAccounts.NAA.Name), $($Config.ServiceAccounts.Admin.Name))"

} catch {
    Write-Status "Service account creation failed: $($_.Exception.Message)" -Level FAIL
    Write-Status 'DC01 may be unresponsive. Verify DC01 is running, then re-run this script.' -Level WARN
}

###############################################################################
# PHASE 5: CONTENT SHARE
###############################################################################

Write-Step 'Phase 5: Content Share'

try {

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

} catch {
    Write-Status "Content share creation failed: $($_.Exception.Message)" -Level FAIL
    Write-Status 'CM01 may be unresponsive. Create the share manually after deployment.' -Level WARN
}

###############################################################################
# PHASE 6: MECM FULL ADMINISTRATOR
###############################################################################

Write-Step 'Phase 6: Add svc-CMAdmin as MECM Full Administrator'

try {

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

} catch {
    Write-Status "MECM admin role assignment failed: $($_.Exception.Message)" -Level FAIL
    Write-Status 'Add svc-CMAdmin as Full Administrator manually via the CM console.' -Level WARN
}

###############################################################################
# PHASE 7: DEPLOY TOOLS
###############################################################################

Write-Step 'Phase 7: Deploy Tools to CM01'

try {

# Create tools directory on CM01
Invoke-LabCommand -ComputerName $cmName -ActivityName 'Create Tools directory' -ScriptBlock {
    if (-not (Test-Path 'C:\Tools')) { New-Item -Path 'C:\Tools' -ItemType Directory -Force | Out-Null }
}

# Copy cc4cm if available locally (check for release zip or build output)
$cc4cmZip = Get-ChildItem 'C:\temp' -Filter 'ClientCenter-*.zip' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($cc4cmZip) {
    Write-Status 'Copying Client Center (cc4cm)...' -Level RUN
    Copy-LabFileItem -Path $cc4cmZip.FullName -ComputerName $cmName -DestinationFolderPath 'C:\temp'
    Invoke-LabCommand -ComputerName $cmName -ActivityName 'Extract cc4cm' -ScriptBlock {
        $zip = Get-ChildItem 'C:\temp' -Filter 'ClientCenter-*.zip' | Select-Object -First 1
        if ($zip) { Expand-Archive -Path $zip.FullName -DestinationPath 'C:\Tools\ClientCenter' -Force }
    }
    Write-Status 'cc4cm deployed to C:\Tools\ClientCenter'
} else {
    Write-Status 'cc4cm release zip not found in C:\temp -- skipping' -Level WARN
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

} catch {
    Write-Status "Tool deployment failed: $($_.Exception.Message)" -Level FAIL
    Write-Status 'Tools are optional. Copy them manually to C:\Tools on CM01.' -Level WARN
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

$elapsed = if ($script:deployStartTime) { (Get-Date) - $script:deployStartTime } else { [TimeSpan]::Zero }

Write-Host ''
Write-Host '  =============================================' -ForegroundColor Green
Write-Host '   MECM HOME LAB DEPLOYMENT COMPLETE' -ForegroundColor Green
Write-Host '  =============================================' -ForegroundColor Green
Write-Host ''
Write-Host "  Elapsed:    $($elapsed.Hours)h $($elapsed.Minutes)m $($elapsed.Seconds)s" -ForegroundColor DarkGray
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

# Close log file
if ($script:logStream) {
    $script:logStream.Close()
    $script:logStream.Dispose()
    Write-Host "  Log saved to: $LogFile" -ForegroundColor DarkGray
}
