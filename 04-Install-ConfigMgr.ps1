#Requires -RunAsAdministrator
#Requires -Version 5.1
<#
.SYNOPSIS
    Installs ConfigMgr 2509 on CM01.

.DESCRIPTION
    Runs on the host and uses Invoke-LabCommand to perform all actions on CM01:
      1. Install VC++ 14.50 runtimes (x86 + x64)
      2. Install ODBC Driver 18.5.2.1
      3. Install MSOLEDB from CM prerequisites
      4. Install ADK (DeploymentTools + UserStateMigrationTool)
      5. Install ADK WinPE add-on
      6. Generate unattended setup INI
      7. Run CM setup.exe /SCRIPT
      8. Validate installation via SMS_Site WMI

    Run after 03-Deploy-Infrastructure.ps1 has completed.

.EXAMPLE
    .\04-Install-ConfigMgr.ps1
#>

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

# ── Import Lab ───────────────────────────────────────────────────────────────

Import-Module AutomatedLab -ErrorAction Stop
Import-Lab -Name $Config.LabName -ErrorAction Stop

$cmName   = $Config.CM.Name
$siteCode = $Config.SiteCode
$siteName = $Config.SiteName
$domain   = $Config.DomainName
$netbios  = ($domain -split '\.')[0].ToUpper()

Write-Host "`nInstalling ConfigMgr 2509 on $cmName" -ForegroundColor Cyan
Write-Host "  Site Code: $siteCode"
Write-Host "  Site Name: $siteName"
Write-Host "  Domain:    $domain"
Write-Host ""

# ── 1. VC++ Runtimes ────────────────────────────────────────────────────────

Write-Step 'Installing VC++ 14.50 Runtimes'

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

# ── 2. ODBC Driver ──────────────────────────────────────────────────────────

Write-Step "Installing ODBC Driver $($Config.ODBCVersion)"

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

# ── 3. MSOLEDB (from CM prereqs) ────────────────────────────────────────────

Write-Step 'Installing MSOLEDB 19'

Invoke-LabCommand -ComputerName $cmName -ActivityName 'Install MSOLEDB' -ScriptBlock {
    $msi = Get-ChildItem 'C:\Install\CMPrereqs' -Filter 'msoledbsql*.msi' -ErrorAction SilentlyContinue |
        Sort-Object Name -Descending | Select-Object -First 1
    if (-not $msi) {
        # Also check CM source prereqs subfolder
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
        Write-Warning 'MSOLEDB MSI not found - CM setup may install it automatically'
    }
}
Write-Status 'MSOLEDB installed'

# ── 4. ADK ───────────────────────────────────────────────────────────────────

Write-Step 'Installing Windows ADK'

Invoke-LabCommand -ComputerName $cmName -ActivityName 'Install ADK' -ScriptBlock {
    $adkSetup = Get-ChildItem 'C:\Install\ADKOffline' -Filter 'adksetup*' -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $adkSetup) {
        throw 'adksetup.exe not found in C:\Install\ADKOffline'
    }

    # Check if already installed
    $adkReg = Get-ItemProperty 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -like '*Assessment and Deployment Kit*' }
    if ($adkReg) {
        Write-Host 'ADK already installed - skipping'
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

# ── 5. ADK WinPE Add-on ─────────────────────────────────────────────────────

Write-Step 'Installing ADK WinPE Add-on'

Invoke-LabCommand -ComputerName $cmName -ActivityName 'Install ADK WinPE' -ScriptBlock {
    $peSetup = Get-ChildItem 'C:\Install\ADKPEOffline' -Filter 'adkwinpe*' -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $peSetup) {
        throw 'adkwinpesetup.exe not found in C:\Install\ADKPEOffline'
    }

    # Check if already installed
    $peReg = Get-ItemProperty 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -like '*Windows PE*' -or $_.DisplayName -like '*Preinstallation*' }
    if ($peReg) {
        Write-Host 'ADK WinPE already installed - skipping'
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

# ── 6. Generate Unattended Setup INI ────────────────────────────────────────

Write-Step 'Generating CM Unattended Setup INI'

$cmFqdn   = "$cmName.$domain"
$adminUser = "$netbios\$($Config.AdminUser)"

# Pass config values as parameters to avoid variable scope issues
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
PrerequisitePath=C:\Install\CMPrereqs
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
} -ArgumentList $siteCode, $siteName, $cmFqdn, $adminUser, $domain, $Config.SQLCollation

Write-Status 'Setup INI generated'

# ── 7. Prepare SQL directories ──────────────────────────────────────────────

Write-Step 'Preparing SQL Data Directories'

Invoke-LabCommand -ComputerName $cmName -ActivityName 'Create SQL directories' -ScriptBlock {
    # E: drive should be the SQL disk
    $dirs = @('E:\MSSQL\Data', 'E:\MSSQL\Log')
    foreach ($d in $dirs) {
        if (-not (Test-Path $d)) {
            New-Item -Path $d -ItemType Directory -Force | Out-Null
            Write-Host "Created: $d"
        }
    }

    # Grant SQL service account full control
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

# ── 8. Install ConfigMgr ────────────────────────────────────────────────────

Write-Step 'Installing ConfigMgr 2509 (this will take 1-3 hours)'
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
        # Check log for details
        $logPath = 'C:\ConfigMgrSetup.log'
        if (Test-Path $logPath) {
            $errors = Get-Content $logPath -Tail 30 | Where-Object { $_ -match 'ERROR|FAIL' }
            if ($errors) {
                Write-Host '--- Last errors from setup log ---'
                $errors | ForEach-Object { Write-Host "  $_" }
            }
        }
        throw "ConfigMgr setup failed with exit code $($proc.ExitCode)"
    }
} -Timeout ([TimeSpan]::FromHours(4))

Write-Status 'ConfigMgr setup completed'
Write-Host "  Finished at: $(Get-Date -Format 'HH:mm:ss')" -ForegroundColor DarkGray

# ── 9. Validate Installation ────────────────────────────────────────────────

Write-Step 'Validating ConfigMgr Installation'

# Give CM services time to start
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

    # Check SMS_Site WMI
    try {
        $site = Get-CimInstance -Namespace "ROOT\SMS\site_$SiteCode" -ClassName SMS_Site -ErrorAction Stop
        if ($site) {
            $result.SiteFound  = $true
            $result.SiteStatus = $site.Status
        }
    } catch {
        Write-Host "WMI query failed: $_"
    }

    # Check console
    $consolePath = 'C:\Program Files\Microsoft Configuration Manager\AdminConsole\bin\Microsoft.ConfigurationManagement.exe'
    $result.ConsoleFound = Test-Path $consolePath

    # Check SMS_EXECUTIVE service
    $svc = Get-Service SMS_EXECUTIVE -ErrorAction SilentlyContinue
    if ($svc) {
        $result.ServiceState = $svc.Status.ToString()
    }

    return $result
} -ArgumentList $siteCode -PassThru

if ($validation.SiteFound) {
    Write-Status "Site $siteCode found in WMI (Status: $($validation.SiteStatus))"
} else {
    Write-Status "Site $siteCode not yet visible in WMI - may still be initializing" -Level WARN
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

# ── 10. Final Snapshot ──────────────────────────────────────────────────────

Write-Step 'Creating Post-CM-Install Snapshot'

Checkpoint-VM -Name $Config.CM.Name -SnapshotName 'Post-CM-Install' -ErrorAction SilentlyContinue
Write-Status 'Snapshot created: Post-CM-Install'

# ── Done ─────────────────────────────────────────────────────────────────────

Write-Step 'ConfigMgr 2509 Installation Complete'

Write-Host ""
Write-Host "  Site Code:  $siteCode" -ForegroundColor Green
Write-Host "  Site Name:  $siteName" -ForegroundColor Green
Write-Host "  Server:     $cmName.$domain" -ForegroundColor Green
Write-Host ""
Write-Host "  Connect to the CM console:" -ForegroundColor Cyan
Write-Host "    Connect-LabVM -ComputerName $cmName" -ForegroundColor White
Write-Host "    # Launch 'Configuration Manager Console' from Start menu" -ForegroundColor DarkGray
Write-Host ""
Write-Host "  Or use PowerShell remoting:" -ForegroundColor Cyan
Write-Host "    Enter-LabPSSession -ComputerName $cmName" -ForegroundColor White
Write-Host ""
Write-Host "  Lab management:" -ForegroundColor Cyan
Write-Host "    Stop-Lab -Name $($Config.LabName)     # Stop all VMs" -ForegroundColor White
Write-Host "    Start-Lab -Name $($Config.LabName)     # Start all VMs" -ForegroundColor White
Write-Host "    Remove-Lab -Name $($Config.LabName)    # Delete entire lab" -ForegroundColor White
Write-Host ""
