#Requires -Modules ActiveDirectory
<#
.SYNOPSIS
    Creates MECM service accounts in Active Directory.

.DESCRIPTION
    Creates the following accounts in OU=Service Accounts:

    svc-CMPush   - Client Push Installation Account (Domain Admins)
    svc-CMNAA    - Network Access Account (Domain Users only)
    svc-CMAdmin  - MECM Admin / cc4cm interactive account (Domain Admins + RDP)

    Run on DC01 or any domain-joined machine with AD PowerShell module.
#>

Import-Module ActiveDirectory

$ouPath = 'OU=Service Accounts,DC=contoso,DC=com'
if (-not (Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '$ouPath'" -ErrorAction SilentlyContinue)) {
    New-ADOrganizationalUnit -Name 'Service Accounts' -Path 'DC=contoso,DC=com'
    Write-Host 'Created OU: Service Accounts'
} else {
    Write-Host 'OU already exists: Service Accounts'
}

# --- Client Push Installation Account ---
$acctName = 'svc-CMPush'
if (-not (Get-ADUser -Filter "SamAccountName -eq '$acctName'" -ErrorAction SilentlyContinue)) {
    New-ADUser -Name 'MECM Client Push' `
        -SamAccountName $acctName `
        -UserPrincipalName "$acctName@contoso.com" `
        -Path $ouPath `
        -AccountPassword (ConvertTo-SecureString 'P@ssw0rd!Push1' -AsPlainText -Force) `
        -PasswordNeverExpires $true `
        -CannotChangePassword $true `
        -Enabled $true `
        -Description 'MECM Client Push Installation Account - local admin on all domain PCs'
    Write-Host "Created: CONTOSO\$acctName"
} else {
    Write-Host "Exists: CONTOSO\$acctName"
}
Add-ADGroupMember -Identity 'Domain Admins' -Members $acctName -ErrorAction SilentlyContinue

# --- Network Access Account ---
$acctName = 'svc-CMNAA'
if (-not (Get-ADUser -Filter "SamAccountName -eq '$acctName'" -ErrorAction SilentlyContinue)) {
    New-ADUser -Name 'MECM Network Access Account' `
        -SamAccountName $acctName `
        -UserPrincipalName "$acctName@contoso.com" `
        -Path $ouPath `
        -AccountPassword (ConvertTo-SecureString 'P@ssw0rd!NAA1' -AsPlainText -Force) `
        -PasswordNeverExpires $true `
        -CannotChangePassword $true `
        -Enabled $true `
        -Description 'MECM Network Access Account - least privilege, domain user only'
    Write-Host "Created: CONTOSO\$acctName"
} else {
    Write-Host "Exists: CONTOSO\$acctName"
}
# NAA is intentionally Domain Users only — no admin rights

# --- MECM Admin / cc4cm Interactive Account ---
$acctName = 'svc-CMAdmin'
if (-not (Get-ADUser -Filter "SamAccountName -eq '$acctName'" -ErrorAction SilentlyContinue)) {
    New-ADUser -Name 'MECM Admin' `
        -SamAccountName $acctName `
        -UserPrincipalName "$acctName@contoso.com" `
        -Path $ouPath `
        -AccountPassword (ConvertTo-SecureString 'P@ssw0rd!Admin1' -AsPlainText -Force) `
        -PasswordNeverExpires $true `
        -CannotChangePassword $true `
        -Enabled $true `
        -Description 'MECM administration - Domain Admin, interactive logon, cc4cm'
    Write-Host "Created: CONTOSO\$acctName"
} else {
    Write-Host "Exists: CONTOSO\$acctName"
}
Add-ADGroupMember -Identity 'Domain Admins' -Members $acctName -ErrorAction SilentlyContinue
Add-ADGroupMember -Identity 'Remote Desktop Users' -Members $acctName -ErrorAction SilentlyContinue

Write-Host "`n=== Service Accounts ==="
Get-ADUser -SearchBase $ouPath -Filter * -Properties Description, MemberOf |
    Select-Object SamAccountName, Description, Enabled |
    Format-Table -AutoSize
