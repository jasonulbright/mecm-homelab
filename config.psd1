@{
    LabName        = 'HomeLab'
    DomainName     = 'contoso.com'
    SiteCode       = 'MCM'
    SiteName       = 'Home Lab Primary Site'
    Network        = '192.168.50'

    # ── CHANGE THESE PASSWORDS BEFORE DEPLOYING ──────────────────────────────
    # Default passwords are published in source control. Treat this lab as
    # internet-facing even if it is not -- change every password below.
    AdminUser      = 'LabAdmin'
    AdminPass      = 'P@ssw0rd!'

    # OS filters -- matched against Get-LabAvailableOperatingSystem from your ISOs.
    # Wildcard patterns. The highest-version match wins.
    ServerOSFilter = 'Windows Server 2025*Desktop Experience*'
    ClientOSFilter = 'Windows 11*Enterprise*'

    # Hyper-V auto-start/stop behavior
    # AutoStart: Start, StartIfRunning, Nothing
    # AutoStop:  ShutDown, Save, TurnOff
    AutoStartAction = 'Start'
    AutoStopAction  = 'ShutDown'

    DC = @{
        Name       = 'DC01'
        IP         = '192.168.50.10'
        Memory     = 2GB
        MinMemory  = 1GB
        MaxMemory  = 2GB
        Processors = 2
        AutoStartDelay = 30   # DC starts first (AD/DNS must be up)
    }
    CM = @{
        Name       = 'CM01'
        IP         = '192.168.50.20'
        Memory     = 10GB
        MinMemory  = 4GB
        MaxMemory  = 12GB
        Processors = 4
        SQLDisk    = 50   # GB
        DataDisk   = 50   # GB
        OSDiskSize = 150  # GB
        AutoStartDelay = 90   # Wait for DC to be ready
    }
    Client = @{
        Name       = 'CLIENT01'
        IP         = '192.168.50.100'
        Memory     = 4GB
        MinMemory  = 2GB
        MaxMemory  = 4GB
        Processors = 2
        OSDiskSize = 150  # GB - must be large enough for 117 app install/uninstall cycles
        AutoStartDelay = 180  # Wait for DC+CM to be ready
    }

    # Service accounts — CHANGE THESE PASSWORDS
    ServiceAccounts = @{
        ClientPush = @{
            Name     = 'svc-CMPush'
            Password = 'P@ssw0rd!Push1'
            Desc     = 'MECM Client Push Installation Account'
            Group    = 'Domain Admins'  # Needs local admin on all targets
        }
        NAA = @{
            Name     = 'svc-CMNAA'
            Password = 'P@ssw0rd!NAA1'
            Desc     = 'MECM Network Access Account'
            Group    = $null  # Domain Users only -- least privilege
        }
        Admin = @{
            Name     = 'svc-CMAdmin'
            Password = 'P@ssw0rd!Admin1'
            Desc     = 'MECM administration - Domain Admin, interactive logon, cc4cm'
            Group    = 'Domain Admins'  # Also added to Remote Desktop Users
        }
    }

    # Software versions
    ODBCVersion    = '18.5.2.1'
    ODBCURL        = 'https://go.microsoft.com/fwlink/?linkid=2335671'
    VCRedistX64URL = 'https://aka.ms/vs/18/release/vc_redist.x64.exe'
    VCRedistX86URL = 'https://aka.ms/vs/18/release/vc_redist.x86.exe'
    SQLCollation   = 'SQL_Latin1_General_CP1_CI_AS'
}
