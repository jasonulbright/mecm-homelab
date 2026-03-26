Import-Module ActiveDirectory

# Create Service Accounts OU
$ouPath = 'OU=Service Accounts,DC=contoso,DC=com'
if (-not (Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '$ouPath'" -ErrorAction SilentlyContinue)) {
    New-ADOrganizationalUnit -Name 'Service Accounts' -Path 'DC=contoso,DC=com'
    Write-Host 'Created OU: Service Accounts'
} else {
    Write-Host 'OU already exists: Service Accounts'
}

# --- Client Push Installation Account ---
$acctName = 'svc-CMPush'
$password = ConvertTo-SecureString 'P@ssw0rd!Push1' -AsPlainText -Force
if (-not (Get-ADUser -Filter "SamAccountName -eq '$acctName'" -ErrorAction SilentlyContinue)) {
    New-ADUser @{
        Name                  = 'MECM Client Push'
        SamAccountName        = $acctName
        UserPrincipalName     = "$acctName@contoso.com"
        Path                  = $ouPath
        AccountPassword       = $password
        PasswordNeverExpires  = $true
        CannotChangePassword  = $true
        Enabled               = $true
        Description           = 'MECM Client Push Installation Account - local admin on all domain PCs'
    }
    Write-Host "Created: CONTOSO\$acctName"
} else {
    Write-Host "Exists: CONTOSO\$acctName"
}
Add-ADGroupMember -Identity 'Domain Admins' -Members $acctName -ErrorAction SilentlyContinue

# --- Network Access Account ---
$acctName = 'svc-CMNAA'
$password = ConvertTo-SecureString 'P@ssw0rd!NAA1' -AsPlainText -Force
if (-not (Get-ADUser -Filter "SamAccountName -eq '$acctName'" -ErrorAction SilentlyContinue)) {
    New-ADUser @{
        Name                  = 'MECM Network Access Account'
        SamAccountName        = $acctName
        UserPrincipalName     = "$acctName@contoso.com"
        Path                  = $ouPath
        AccountPassword       = $password
        PasswordNeverExpires  = $true
        CannotChangePassword  = $true
        Enabled               = $true
        Description           = 'MECM Network Access Account - least privilege, domain user only'
    }
    Write-Host "Created: CONTOSO\$acctName"
} else {
    Write-Host "Exists: CONTOSO\$acctName"
}
# NAA is intentionally Domain Users only — no admin rights

Write-Host "`n=== Service Accounts ==="
Get-ADUser -SearchBase $ouPath -Filter * -Properties Description, MemberOf |
    Select-Object SamAccountName, Description, Enabled |
    Format-Table -AutoSize
