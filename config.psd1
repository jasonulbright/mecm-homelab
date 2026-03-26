@{
    LabName        = 'HomeLab'
    DomainName     = 'contoso.com'
    SiteCode       = 'MCM'
    SiteName       = 'Home Lab Primary Site'
    Network        = '192.168.50'
    AdminUser      = 'LabAdmin'
    AdminPass      = 'P@ssw0rd!'
    DC = @{
        Name       = 'DC01'
        IP         = '192.168.50.10'
        Memory     = 2GB
        MaxMemory  = 2GB
        Processors = 2
    }
    CM = @{
        Name       = 'CM01'
        IP         = '192.168.50.20'
        Memory     = 10GB
        MaxMemory  = 12GB
        Processors = 4
        SQLDisk    = 50   # GB
        DataDisk   = 50   # GB
        OSDiskSize = 150  # GB
    }
    Client = @{
        Name       = 'CLIENT01'
        IP         = '192.168.50.100'
        Memory     = 4GB
        MaxMemory  = 4GB
        Processors = 2
    }
    # Software versions
    ODBCVersion    = '18.5.2.1'
    ODBCURL        = 'https://go.microsoft.com/fwlink/?linkid=2335671'
    VCRedistX64URL = 'https://aka.ms/vs/18/release/vc_redist.x64.exe'
    VCRedistX86URL = 'https://aka.ms/vs/18/release/vc_redist.x86.exe'
    SQLCollation   = 'SQL_Latin1_General_CP1_CI_AS'
}
