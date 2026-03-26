#Requires -RunAsAdministrator
#Requires -Version 5.1
<#
.SYNOPSIS
    Deploys the lab infrastructure: DC01 + CM01 VMs with AD, CA, and SQL.

.DESCRIPTION
    Uses AutomatedLab to create a 2-VM lab:
      - DC01: RootDC + CaRoot + Routing (Windows Server 2025)
      - CM01: SQL Server 2022 (Windows Server 2025)

    Both VMs get a second NIC on the Default Switch for internet access.
    CM01 OS disk is expanded to 150GB. SQL memory is configured.
    ADK, CM source, prerequisites, and runtimes are copied to CM01.
    AD schema extension is run. Both VMs are snapshotted.

    Run after 02-Download-Offline.ps1 has completed.

.PARAMETER RemoveExisting
    If specified, removes an existing lab with the same name without prompting.

.EXAMPLE
    .\03-Deploy-Infrastructure.ps1

.EXAMPLE
    .\03-Deploy-Infrastructure.ps1 -RemoveExisting
#>

param(
    [switch]$RemoveExisting
)

$ErrorActionPreference = 'Stop'

# ── Load config ──────────────────────────────────────────────────────────────

$Config = Import-PowerShellDataFile -Path (Join-Path $PSScriptRoot 'config.psd1')

# ── Helpers ──────────────────────────────────────────────────────────────────

function Write-Step {
    param([string]$Title)
    Write-Host "`n=== $Title ===" -ForegroundColor Cyan
}

function Write-Status {
    param([string]$Message, [ValidateSet('OK','WARN','FAIL','INFO','RUN')]$Level = 'OK')
    switch ($Level) {
        'OK'   { Write-Host "  [OK]   $Message" -ForegroundColor Green }
        'WARN' { Write-Host "  [WARN] $Message" -ForegroundColor Yellow }
        'FAIL' { Write-Host "  [FAIL] $Message" -ForegroundColor Red }
        'INFO' { Write-Host "  [INFO] $Message" -ForegroundColor Cyan }
        'RUN'  { Write-Host "  [RUN]  $Message" -ForegroundColor Yellow }
    }
}

# ── Init ─────────────────────────────────────────────────────────────────────

Import-Module AutomatedLab -ErrorAction Stop

$labSources    = Get-LabSourcesLocation
$swPkg         = Join-Path $labSources 'SoftwarePackages'
$labName       = $Config.LabName
$domainName    = $Config.DomainName
$netPrefix     = $Config.Network
$networkName   = "$labName-Network"

# ── Check for existing lab ───────────────────────────────────────────────────

$existingLabs = Get-Lab -List -ErrorAction SilentlyContinue
if ($existingLabs -contains $labName) {
    if ($RemoveExisting) {
        Write-Host "Removing existing lab '$labName'..." -ForegroundColor Yellow
        Remove-Lab -Name $labName -Confirm:$false
    } else {
        $response = Read-Host "Lab '$labName' already exists. Remove and recreate? (y/N)"
        if ($response -ne 'y') {
            Write-Host "Aborted." -ForegroundColor Yellow
            return
        }
        Remove-Lab -Name $labName -Confirm:$false
    }
}

# ── Lab Definition ───────────────────────────────────────────────────────────

Write-Step "Defining Lab: $labName"

New-LabDefinition -Name $labName -DefaultVirtualizationEngine HyperV

# Internal network
Add-LabVirtualNetworkDefinition -Name $networkName `
    -AddressSpace "$netPrefix.0/24" `
    -HyperVProperties @{ SwitchType = 'Internal' }

# Default Switch for internet access (NAT)
Add-LabVirtualNetworkDefinition -Name 'Default Switch' `
    -HyperVProperties @{ SwitchType = 'External' }

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

# ── DC01: Domain Controller ─────────────────────────────────────────────────

Write-Step 'Defining DC01'

# NO DHCP role -- AutomatedLab throws "not implemented yet"
$dcRoles = @(
    Get-LabMachineRoleDefinition -Role RootDC
    Get-LabMachineRoleDefinition -Role CaRoot
    Get-LabMachineRoleDefinition -Role Routing
)

$dcNics = @(
    New-LabNetworkAdapterDefinition -VirtualSwitch $networkName -Ipv4Address "$($Config.DC.IP)/24"
    New-LabNetworkAdapterDefinition -VirtualSwitch 'Default Switch' -UseDhcp
)

Add-LabMachineDefinition -Name $Config.DC.Name `
    -Roles $dcRoles `
    -Memory $Config.DC.Memory `
    -MaxMemory $Config.DC.MaxMemory `
    -Processors $Config.DC.Processors `
    -NetworkAdapter $dcNics `
    -DnsServer1 $Config.DC.IP `
    -DomainName $domainName `
    -OperatingSystem 'Windows Server 2025 Datacenter Evaluation (Desktop Experience)'

Write-Status "DC01 defined: $($Config.DC.IP), $([math]::Round($Config.DC.Memory/1GB))GB RAM, $($Config.DC.Processors) vCPU"

# ── CM01: SQL Server ─────────────────────────────────────────────────────────

Write-Step 'Defining CM01'

$sqlRole = Get-LabMachineRoleDefinition -Role SQLServer2022 -Properties @{
    Collation    = $Config.SQLCollation
    InstanceName = 'MSSQLSERVER'
}

# Additional disks
Add-LabDiskDefinition -Name 'CM01-SQL' -DiskSizeInGb $Config.CM.SQLDisk
Add-LabDiskDefinition -Name 'CM01-Data' -DiskSizeInGb $Config.CM.DataDisk

$cmNics = @(
    New-LabNetworkAdapterDefinition -VirtualSwitch $networkName -Ipv4Address "$($Config.CM.IP)/24"
    New-LabNetworkAdapterDefinition -VirtualSwitch 'Default Switch' -UseDhcp
)

Add-LabMachineDefinition -Name $Config.CM.Name `
    -Roles $sqlRole `
    -Memory $Config.CM.Memory `
    -MaxMemory $Config.CM.MaxMemory `
    -Processors $Config.CM.Processors `
    -DiskName 'CM01-SQL', 'CM01-Data' `
    -NetworkAdapter $cmNics `
    -DnsServer1 $Config.DC.IP `
    -DomainName $domainName `
    -OperatingSystem 'Windows Server 2025 Datacenter Evaluation (Desktop Experience)'

Write-Status "CM01 defined: $($Config.CM.IP), $([math]::Round($Config.CM.Memory/1GB))GB RAM, $($Config.CM.Processors) vCPU"

# ── Expand CM01 OS Disk ─────────────────────────────────────────────────────

Write-Step 'Expanding CM01 OS Disk'
Write-Status "Will expand to $($Config.CM.OSDiskSize)GB after VM creation" -Level INFO

# ── Install Lab ──────────────────────────────────────────────────────────────

Write-Step 'Installing Lab (this will take 30-60 minutes)'

Install-Lab -DelayBetweenComputers 30 -NoValidation

Write-Status 'Lab VMs deployed and domain joined'

# ── Expand CM01 OS disk now that VM exists ───────────────────────────────────

Write-Step 'Expanding CM01 OS Disk to 150GB'

$cmVM = Get-VM -Name $Config.CM.Name -ErrorAction SilentlyContinue
if ($cmVM) {
    $osDisk = $cmVM | Get-VMHardDiskDrive | Where-Object { $_.ControllerLocation -eq 0 } | Select-Object -First 1
    if ($osDisk) {
        $currentSizeGB = [math]::Round((Get-VHD $osDisk.Path).Size / 1GB, 0)
        if ($currentSizeGB -lt $Config.CM.OSDiskSize) {
            Resize-VHD -Path $osDisk.Path -SizeBytes ($Config.CM.OSDiskSize * 1GB)
            Write-Status "VHDX expanded to $($Config.CM.OSDiskSize)GB"

            # Extend partition inside the VM
            Invoke-LabCommand -ComputerName $Config.CM.Name -ActivityName 'Extend C: partition' -ScriptBlock {
                $maxSize = (Get-PartitionSupportedSize -DriveLetter C).SizeMax
                Resize-Partition -DriveLetter C -Size $maxSize
            }
            Write-Status 'C: partition extended inside VM'
        } else {
            Write-Status "OS disk already ${currentSizeGB}GB (>= $($Config.CM.OSDiskSize)GB)" -Level INFO
        }
    }
}

# ── Configure SQL Memory ────────────────────────────────────────────────────

Write-Step 'Configuring SQL Server Memory (8GB min/max)'

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

# ── Copy Software to CM01 ───────────────────────────────────────────────────

Write-Step 'Copying Software to CM01'

# Create install directory on CM01
Invoke-LabCommand -ComputerName $Config.CM.Name -ActivityName 'Create Install directories' -ScriptBlock {
    $dirs = @(
        'C:\Install'
        'C:\Install\ADKOffline'
        'C:\Install\ADKPEOffline'
        'C:\Install\CM'
        'C:\Install\CMPrereqs'
        'C:\Install\ODBC'
        'C:\Install\VCRedist'
    )
    foreach ($d in $dirs) {
        if (-not (Test-Path $d)) { New-Item -Path $d -ItemType Directory -Force | Out-Null }
    }
}

# ADK offline layout
$adkLayout = Join-Path $swPkg 'ADK\Offline'
if (Test-Path $adkLayout) {
    Write-Status 'Copying ADK offline layout...' -Level RUN
    Copy-LabFileItem -Path $adkLayout -ComputerName $Config.CM.Name -DestinationFolderPath 'C:\Install\ADKOffline' -Recurse

    # Flatten: Copy-LabFileItem nests the folder, so move contents up if needed
    Invoke-LabCommand -ComputerName $Config.CM.Name -ActivityName 'Flatten ADK folder' -ScriptBlock {
        $nested = 'C:\Install\ADKOffline\Offline'
        if (Test-Path $nested) {
            Get-ChildItem $nested | Move-Item -Destination 'C:\Install\ADKOffline' -Force
            Remove-Item $nested -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    Write-Status 'ADK offline layout copied'
} else {
    Write-Status 'ADK offline layout not found - skipping' -Level WARN
}

# ADK PE offline layout
$adkPeLayout = Join-Path $swPkg 'ADKPE\Offline'
if (Test-Path $adkPeLayout) {
    Write-Status 'Copying ADK PE offline layout...' -Level RUN
    Copy-LabFileItem -Path $adkPeLayout -ComputerName $Config.CM.Name -DestinationFolderPath 'C:\Install\ADKPEOffline' -Recurse

    Invoke-LabCommand -ComputerName $Config.CM.Name -ActivityName 'Flatten ADK PE folder' -ScriptBlock {
        $nested = 'C:\Install\ADKPEOffline\Offline'
        if (Test-Path $nested) {
            Get-ChildItem $nested | Move-Item -Destination 'C:\Install\ADKPEOffline' -Force
            Remove-Item $nested -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    Write-Status 'ADK PE offline layout copied'
} else {
    Write-Status 'ADK PE offline layout not found - skipping' -Level WARN
}

# CM source
$cmSourceDir = Get-ChildItem (Join-Path $swPkg 'CM') -Directory -ErrorAction SilentlyContinue | Select-Object -First 1
if ($cmSourceDir) {
    Write-Status "Copying CM source ($($cmSourceDir.Name))..." -Level RUN
    Copy-LabFileItem -Path $cmSourceDir.FullName -ComputerName $Config.CM.Name -DestinationFolderPath 'C:\Install\CM' -Recurse

    # Flatten nested folder (Copy-LabFileItem creates CM\ConfigMgr_2509\)
    Invoke-LabCommand -ComputerName $Config.CM.Name -ActivityName 'Flatten CM folder' -ScriptBlock {
        $cmBase = 'C:\Install\CM'
        $subDirs = Get-ChildItem $cmBase -Directory -ErrorAction SilentlyContinue
        foreach ($sub in $subDirs) {
            Get-ChildItem $sub.FullName | Move-Item -Destination $cmBase -Force
            Remove-Item $sub.FullName -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    Write-Status 'CM source copied and flattened'
} else {
    Write-Status 'CM source folder not found - skipping' -Level WARN
}

# CM prerequisites
$prereqDir = Join-Path $swPkg 'CMPrereqs'
if ((Get-ChildItem $prereqDir -ErrorAction SilentlyContinue | Measure-Object).Count -gt 0) {
    Write-Status 'Copying CM prerequisites...' -Level RUN
    Copy-LabFileItem -Path $prereqDir -ComputerName $Config.CM.Name -DestinationFolderPath 'C:\Install' -Recurse

    Invoke-LabCommand -ComputerName $Config.CM.Name -ActivityName 'Flatten CMPrereqs folder' -ScriptBlock {
        $nested = 'C:\Install\CMPrereqs\CMPrereqs'
        if (Test-Path $nested) {
            Get-ChildItem $nested | Move-Item -Destination 'C:\Install\CMPrereqs' -Force
            Remove-Item $nested -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    Write-Status 'CM prerequisites copied'
} else {
    Write-Status 'CM prerequisites folder empty - skipping' -Level WARN
}

# ODBC driver
$odbcFile = Get-ChildItem (Join-Path $swPkg 'ODBC') -Filter '*.msi' -ErrorAction SilentlyContinue | Select-Object -First 1
if ($odbcFile) {
    Write-Status 'Copying ODBC driver...' -Level RUN
    Copy-LabFileItem -Path $odbcFile.FullName -ComputerName $Config.CM.Name -DestinationFolderPath 'C:\Install\ODBC'
    Write-Status 'ODBC driver copied'
} else {
    Write-Status 'ODBC MSI not found - skipping' -Level WARN
}

# VC++ runtimes
$vcDir = Join-Path $swPkg 'VCRedist'
if (Test-Path (Join-Path $vcDir 'vc_redist.x64.exe')) {
    Write-Status 'Copying VC++ runtimes...' -Level RUN
    Copy-LabFileItem -Path (Join-Path $vcDir 'vc_redist.x64.exe') -ComputerName $Config.CM.Name -DestinationFolderPath 'C:\Install\VCRedist'
    Copy-LabFileItem -Path (Join-Path $vcDir 'vc_redist.x86.exe') -ComputerName $Config.CM.Name -DestinationFolderPath 'C:\Install\VCRedist'
    Write-Status 'VC++ runtimes copied'
} else {
    Write-Status 'VC++ runtimes not found - skipping' -Level WARN
}

# ── AD Schema Extension ─────────────────────────────────────────────────────

Write-Step 'Extending AD Schema for ConfigMgr'

Invoke-LabCommand -ComputerName $Config.CM.Name -ActivityName 'Run extadsch.exe' -ScriptBlock {
    $extadsch = Get-ChildItem 'C:\Install\CM' -Filter 'extadsch.exe' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($extadsch) {
        $result = Start-Process -FilePath $extadsch.FullName -Wait -PassThru -NoNewWindow
        if (Test-Path 'C:\ExtADSch.log') {
            $log = Get-Content 'C:\ExtADSch.log' -Tail 5
            $log | ForEach-Object { Write-Host "    $_" }
        }
    } else {
        Write-Warning 'extadsch.exe not found in C:\Install\CM'
    }
}
Write-Status 'AD schema extension complete'

# ── Snapshot ─────────────────────────────────────────────────────────────────

Write-Step 'Creating Pre-CM-Install Snapshots'

Checkpoint-VM -Name $Config.DC.Name -SnapshotName 'Pre-CM-Install' -ErrorAction SilentlyContinue
Checkpoint-VM -Name $Config.CM.Name -SnapshotName 'Pre-CM-Install' -ErrorAction SilentlyContinue
Write-Status 'Snapshots created: Pre-CM-Install'

# ── Connection Info ──────────────────────────────────────────────────────────

Write-Step 'Deployment Complete'

Write-Host ""
Write-Host "  Domain:     $domainName" -ForegroundColor White
Write-Host "  Admin:      $domainName\$($Config.AdminUser)" -ForegroundColor White
Write-Host "  Password:   $($Config.AdminPass)" -ForegroundColor White
Write-Host "  DC01:       $($Config.DC.IP)" -ForegroundColor White
Write-Host "  CM01:       $($Config.CM.IP)" -ForegroundColor White
Write-Host ""
Write-Host "  Useful commands:" -ForegroundColor Cyan
Write-Host "    Enter-LabPSSession -ComputerName CM01    # PS remoting"
Write-Host "    Connect-LabVM -ComputerName CM01          # RDP"
Write-Host ""
Write-Host "Next step:" -ForegroundColor Green
Write-Host "  .\04-Install-ConfigMgr.ps1" -ForegroundColor White
Write-Host ""
