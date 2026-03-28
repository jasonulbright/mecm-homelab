#requires -Module Pester

<#
.SYNOPSIS
    Pester tests for Deploy-HomeLab.ps1 and config.psd1.
    Validates config structure, password security, OS filters,
    script structure, and parameter handling -- all without VMs.
#>

BeforeAll {
    $script:repoRoot = Split-Path -Parent $PSScriptRoot
    $script:configPath = Join-Path $repoRoot 'config.psd1'
    $script:scriptPath = Join-Path $repoRoot 'Deploy-HomeLab.ps1'
    $script:scriptContent = Get-Content $scriptPath -Raw
    $script:config = Import-PowerShellDataFile -Path $configPath
}

Describe 'config.psd1 structure' {

    It 'Config file exists and is valid PowerShell data' {
        { Import-PowerShellDataFile -Path $configPath } | Should -Not -Throw
    }

    It 'Has required top-level keys' {
        $required = @('LabName', 'DomainName', 'SiteCode', 'SiteName', 'Network',
                       'AdminUser', 'AdminPass', 'DC', 'CM', 'Client',
                       'ServiceAccounts', 'ServerOSFilter', 'ClientOSFilter',
                       'ODBCVersion', 'SQLCollation')
        foreach ($key in $required) {
            $config.ContainsKey($key) | Should -BeTrue -Because "config must have '$key'"
        }
    }

    It 'DC has required VM properties' {
        foreach ($prop in @('Name', 'IP', 'Memory', 'MinMemory', 'MaxMemory', 'Processors')) {
            $config.DC.ContainsKey($prop) | Should -BeTrue -Because "DC must have '$prop'"
        }
    }

    It 'CM has required VM properties including disk sizes' {
        foreach ($prop in @('Name', 'IP', 'Memory', 'MinMemory', 'MaxMemory', 'Processors', 'SQLDisk', 'DataDisk', 'OSDiskSize')) {
            $config.CM.ContainsKey($prop) | Should -BeTrue -Because "CM must have '$prop'"
        }
    }

    It 'Client has required VM properties' {
        foreach ($prop in @('Name', 'IP', 'Memory', 'MinMemory', 'MaxMemory', 'Processors')) {
            $config.Client.ContainsKey($prop) | Should -BeTrue -Because "Client must have '$prop'"
        }
    }

    It 'All 3 service accounts defined with Name and Password' {
        foreach ($acct in @('ClientPush', 'NAA', 'Admin')) {
            $config.ServiceAccounts[$acct].Name | Should -Not -BeNullOrEmpty -Because "$acct must have a Name"
            $config.ServiceAccounts[$acct].Password | Should -Not -BeNullOrEmpty -Because "$acct must have a Password"
        }
    }

    It 'MinMemory <= MaxMemory for all VMs' {
        $config.DC.MinMemory | Should -BeLessOrEqual $config.DC.MaxMemory
        $config.CM.MinMemory | Should -BeLessOrEqual $config.CM.MaxMemory
        $config.Client.MinMemory | Should -BeLessOrEqual $config.Client.MaxMemory
    }

    It 'IP addresses are in the configured network prefix' {
        $config.DC.IP | Should -BeLike "$($config.Network).*"
        $config.CM.IP | Should -BeLike "$($config.Network).*"
        $config.Client.IP | Should -BeLike "$($config.Network).*"
    }

    It 'Site code is 3 characters' {
        $config.SiteCode.Length | Should -Be 3
    }
}

Describe 'Password security' {

    It 'Config ships with default passwords (warning expected at runtime)' {
        # This test documents that defaults exist -- the script warns at runtime.
        # When a user changes them, this test will start failing, which is fine.
        $defaults = @('P@ssw0rd!', 'P@ssw0rd!Push1', 'P@ssw0rd!NAA1', 'P@ssw0rd!Admin1')
        $config.AdminPass | Should -BeIn $defaults
    }

    It 'Deploy script contains default password detection logic' {
        $scriptContent | Should -Match 'DEFAULT PASSWORDS DETECTED'
    }

    It 'Config has password-change comment' {
        $configRaw = Get-Content $configPath -Raw
        $configRaw | Should -Match 'CHANGE THESE PASSWORDS'
    }
}

Describe 'OS filter patterns' {

    It 'ServerOSFilter is a wildcard pattern containing Server' {
        $config.ServerOSFilter | Should -BeLike '*Server*'
    }

    It 'ClientOSFilter is a wildcard pattern containing Windows' {
        $config.ClientOSFilter | Should -BeLike '*Windows*'
    }

    It 'Deploy script uses Get-LabAvailableOperatingSystem for OS resolution' {
        $scriptContent | Should -Match 'Get-LabAvailableOperatingSystem'
        $scriptContent | Should -Match 'ServerOSFilter'
        $scriptContent | Should -Match 'ClientOSFilter'
    }

    It 'No hardcoded OS edition names in Add-LabMachineDefinition calls' {
        # Should use $serverOS / $clientOS variables, not literal strings
        $scriptContent | Should -Not -Match "OperatingSystem\s+'Windows Server 2025"
        $scriptContent | Should -Not -Match "OperatingSystem\s+'Windows 11 Enterprise"
    }
}

Describe 'Deploy-HomeLab.ps1 script structure' {

    It 'Script parses without errors' {
        $errors = $null
        [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$null, [ref]$errors)
        $errors.Count | Should -Be 0
    }

    It 'Requires administrator' {
        $scriptContent | Should -Match '#Requires -RunAsAdministrator'
    }

    It 'Has RemoveExisting parameter' {
        $scriptContent | Should -Match '\[switch\]\$RemoveExisting'
    }

    It 'Does not use Read-Host (non-interactive)' {
        $scriptContent | Should -Not -Match 'Read-Host'
    }

    It 'Uses -ArgumentList for service account creation (no string interpolation of passwords)' {
        # The service account Invoke-LabCommand should use -ArgumentList, not a here-string script
        $scriptContent | Should -Not -Match 'svcAccountScript'
        $scriptContent | Should -Match 'Create service accounts'
        $scriptContent | Should -Match '-ArgumentList \$domainDN'
    }

    It 'Casts Invoke-LabCommand -PassThru results to avoid PSObject issues' {
        $scriptContent | Should -Match '\[bool\]\(\$result'
    }

    It 'Does not duplicate SQL memory configuration (handled by AutomatedLab fork)' {
        $scriptContent | Should -Not -Match "sp_configure 'max server memory'"
    }

    It 'Has try/catch on Phases 4 through 7' {
        # Each phase header should be followed by a try block
        $phase4Pos = $scriptContent.IndexOf('PHASE 4:')
        $phase5Pos = $scriptContent.IndexOf('PHASE 5:')
        $phase6Pos = $scriptContent.IndexOf('PHASE 6:')
        $phase7Pos = $scriptContent.IndexOf('PHASE 7:')
        $phase8Pos = $scriptContent.IndexOf('PHASE 8:')

        # Between each phase header and the next, there should be both 'try {' and '} catch'
        foreach ($pair in @(
            @($phase4Pos, $phase5Pos, '4'),
            @($phase5Pos, $phase6Pos, '5'),
            @($phase6Pos, $phase7Pos, '6'),
            @($phase7Pos, $phase8Pos, '7')
        )) {
            $section = $scriptContent.Substring($pair[0], $pair[1] - $pair[0])
            $section | Should -Match 'try\s*\{' -Because "Phase $($pair[2]) should have try block"
            $section | Should -Match '\}\s*catch' -Because "Phase $($pair[2]) should have catch block"
        }
    }

    It 'Displays elapsed time in completion banner' {
        $scriptContent | Should -Match '\$elapsed'
        $scriptContent | Should -Match 'Elapsed:'
    }
}

Describe 'Vendored AutomatedLab modules' {

    It 'Vendored lib directory exists' {
        $vendoredPath = Join-Path $repoRoot 'lib\AutomatedLab'
        Test-Path $vendoredPath | Should -BeTrue
    }

    It 'AutomatedLab manifest has no Recipe/Ships/Test references' {
        $vendoredPath = Join-Path $repoRoot 'lib\AutomatedLab'
        $manifest = Get-Content (Join-Path $vendoredPath 'AutomatedLab\AutomatedLab.psd1') -Raw
        $manifest | Should -Not -Match 'AutomatedLab\.Recipe'
        $manifest | Should -Not -Match 'AutomatedLab\.Ships'
        $manifest | Should -Not -Match 'AutomatedLabTest'
    }

    It 'Deploy script does not regex-patch the manifest at runtime' {
        $scriptContent | Should -Not -Match 'content -replace.*Recipe'
    }

    It 'Key modules are present' {
        $vendoredPath = Join-Path $repoRoot 'lib\AutomatedLab'
        foreach ($mod in @('AutomatedLab', 'AutomatedLabCore', 'AutomatedLabDefinition',
                           'AutomatedLabWorker', 'AutomatedLabUnattended')) {
            Test-Path (Join-Path $vendoredPath $mod) | Should -BeTrue -Because "$mod should be vendored"
        }
    }
}
