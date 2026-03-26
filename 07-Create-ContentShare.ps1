#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Creates the MECM content share on CM01.

.DESCRIPTION
    Creates E:\ContentShare with subfolders for Applications, Drivers, Images,
    OperatingSystems, Packages, Scripts, and SoftwareUpdates. Shares as
    ContentShare$ (hidden) with appropriate NTFS and share permissions.

    Run from the host with AutomatedLab imported, or directly on CM01.
#>

$config = Import-PowerShellDataFile -Path (Join-Path $PSScriptRoot 'config.psd1')
$cmName = $config.CM.Name
$domainName = $config.DomainName
$domainNetBIOS = $domainName.Split('.')[0].ToUpper()

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

Write-Host "=== Creating content share on $cmName ===" -ForegroundColor Cyan

Import-Lab -Name $config.LabName -NoValidation

Invoke-LabCommand -ComputerName $cmName -ScriptBlock {
    param($sharePath, $shareName, $folders, $domainNetBIOS)

    # Create folder structure
    New-Item -Path $sharePath -ItemType Directory -Force | Out-Null
    foreach ($f in $folders) {
        New-Item -Path (Join-Path $sharePath $f) -ItemType Directory -Force | Out-Null
    }
    Write-Host "Created folder structure at $sharePath"

    # Create SMB share
    if (-not (Get-SmbShare -Name $shareName -ErrorAction SilentlyContinue)) {
        New-SmbShare -Name $shareName -Path $sharePath `
            -FullAccess "$domainNetBIOS\Domain Admins" `
            -ReadAccess "$domainNetBIOS\Domain Computers", "$domainNetBIOS\svc-CMNAA"
        Write-Host "Created share: \\$env:COMPUTERNAME\$shareName"
    } else {
        Write-Host "Share already exists: \\$env:COMPUTERNAME\$shareName"
    }

    # Set NTFS permissions
    $acl = Get-Acl $sharePath
    $naaRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
        "$domainNetBIOS\svc-CMNAA", 'Read', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
    $acl.AddAccessRule($naaRule)
    $compRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
        "$domainNetBIOS\Domain Computers", 'Read', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
    $acl.AddAccessRule($compRule)
    Set-Acl $sharePath $acl
    Write-Host "NTFS permissions set (Domain Admins=Full, Domain Computers+NAA=Read)"

    # List results
    Write-Host ""
    Get-ChildItem $sharePath -Directory | ForEach-Object { Write-Host "  $($_.Name)/" }
} -ArgumentList $sharePath, $shareName, $folders, $domainNetBIOS -PassThru

Write-Host ""
Write-Host "Configure in ApplicationPackager: File > Preferences > File Share Root = \\$cmName\$shareName" -ForegroundColor Yellow
