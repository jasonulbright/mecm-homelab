# Extend AD schema for ConfigMgr and configure System Management container
# Must run from a machine with the CM source files (CM01)
# Requires Schema Admin and Domain Admin rights

param(
    [string]$CMInstallPath = 'C:\Install\CM',
    [string]$SiteServerName = 'CM01',
    [string]$DomainDN = 'DC=contoso,DC=com'
)

Import-Module ActiveDirectory

# --- Extend AD Schema ---
$extadsch = Join-Path $CMInstallPath 'SMSSETUP\BIN\X64\extadsch.exe'
if (-not (Test-Path $extadsch)) {
    Write-Error "extadsch.exe not found at $extadsch"
    return
}

$schemaCheck = Get-ADObject -SearchBase ((Get-ADRootDSE).schemaNamingContext) -Filter {name -eq 'MS-SMS-Site-Code'} -ErrorAction SilentlyContinue
if ($schemaCheck) {
    Write-Host 'AD Schema: Already extended'
} else {
    Write-Host 'Extending AD schema...'
    & $extadsch
    Start-Sleep -Seconds 5
    if (Test-Path 'C:\ExtADSch.log') {
        $result = Get-Content 'C:\ExtADSch.log' | Select-String 'Successfully extended'
        if ($result) { Write-Host 'AD Schema: Extended successfully' }
        else { Write-Warning 'Schema extension may have failed — check C:\ExtADSch.log' }
    }
}

# --- Create System Management container ---
$sysManPath = "CN=System Management,CN=System,$DomainDN"
$systemPath = "CN=System,$DomainDN"

if (-not (Get-ADObject -Filter "DistinguishedName -eq '$sysManPath'" -ErrorAction SilentlyContinue)) {
    New-ADObject -Name 'System Management' -Type Container -Path $systemPath
    Write-Host "Created container: $sysManPath"
} else {
    Write-Host "Container exists: $sysManPath"
}

# --- Grant site server Full Control ---
$siteServer = Get-ADComputer $SiteServerName
$acl = Get-Acl "AD:\$sysManPath"

$identity = [System.Security.Principal.SecurityIdentifier]$siteServer.SID
$rights = [System.DirectoryServices.ActiveDirectoryRights]::GenericAll
$type = [System.Security.AccessControl.AccessControlType]::Allow
$inheritance = [System.DirectoryServices.ActiveDirectorySecurityInheritance]::SelfAndChildren

$ace = New-Object System.DirectoryServices.ActiveDirectoryAccessRule($identity, $rights, $type, $inheritance)
$acl.AddAccessRule($ace)
Set-Acl "AD:\$sysManPath" $acl

Write-Host "${SiteServerName}$ granted Full Control on System Management container"
