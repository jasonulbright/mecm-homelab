# MECM Home Lab Deployment

Automated deployment of a Microsoft Endpoint Configuration Manager (MECM) home lab using a [vendored fork of AutomatedLab](https://github.com/jasonulbright/AutomatedLab) and Hyper-V. No external dependencies — everything is self-contained.

Deploys a fully functional ConfigMgr 2509 primary site with Windows Server 2025, SQL Server 2022, Active Directory, and Certificate Services.

## Hardware Requirements

| Resource | Minimum | Recommended |
|----------|---------|-------------|
| RAM      | 32 GB   | 64 GB       |
| Disk     | 300 GB free (SSD/NVMe) | 500 GB+ |
| CPU      | 4 cores | 8+ cores    |
| OS       | Windows 10/11 Pro (Hyper-V capable) | Windows 11 Pro |

## Lab Architecture

| VM | Role | IP | RAM | vCPU | OS |
|----|------|----|-----|------|----|
| DC01 | Domain Controller, Root CA, Routing | 192.168.50.10 | 2 GB | 2 | Windows Server 2025 |
| CM01 | SQL Server 2022, ConfigMgr 2509 | 192.168.50.20 | 10-12 GB | 4 | Windows Server 2025 |

Both VMs have a second NIC on the Hyper-V `Default Switch` for internet access via NAT.

## Software Prerequisites

Download these before starting:

| Software | Location | Download URL |
|----------|----------|-------------|
| Windows Server 2025 Evaluation ISO | `C:\LabSources\ISOs\` | https://www.microsoft.com/en-us/evalcenter/evaluate-windows-server-2025 |
| SQL Server 2022 Evaluation ISO | `C:\LabSources\ISOs\` | https://www.microsoft.com/en-us/evalcenter/download-sql-server-2022 |
| ConfigMgr 2509 Baseline | `C:\LabSources\SoftwarePackages\CM\` | https://www.microsoft.com/en-us/evalcenter/download-microsoft-endpoint-configuration-manager |
| Windows ADK (adksetup.exe) | `C:\LabSources\SoftwarePackages\ADK\` | https://go.microsoft.com/fwlink/?linkid=2289980 |
| Windows PE add-on (adkwinpesetup.exe) | `C:\LabSources\SoftwarePackages\ADKPE\` | https://go.microsoft.com/fwlink/?linkid=2289981 |

The ODBC Driver and VC++ runtimes are downloaded automatically by the deploy script.

## Quick Start

Must be run as **Administrator**. Enable Hyper-V first (reboot required), then download the ISOs and software listed above into `C:\LabSources\`.

```powershell
# Single command deploys the entire lab (~2-4 hours)
.\Deploy-HomeLab.ps1
```

The script will:
1. Verify prerequisites (Hyper-V, RAM, ISOs)
2. Install AutomatedLab, create ADK offline layouts, download runtimes
3. Deploy DC01 + CM01 VMs with AD, CA, SQL
4. Install ADK, ODBC, VC++, and ConfigMgr 2509 on CM01
5. Extend AD schema, create service accounts, content share
6. Deploy tools (cc4cm + ApplicationPackager) and create snapshots
7. Print connection info and remaining manual console steps

To remove an existing lab and redeploy:
```powershell
.\Deploy-HomeLab.ps1 -RemoveExisting
```

## What the Script Does

`Deploy-HomeLab.ps1` runs through 13 phases:

| Phase | Description |
|-------|-------------|
| 1. Prerequisites | Checks Hyper-V, host RAM, installs vendored AutomatedLab, creates LabSources, verifies ISOs |
| 2. Offline Downloads | Creates ADK/PE offline layouts, downloads CM prereqs, ODBC 18.5.2.1, VC++ 14.50 |
| 3. Lab Deployment | Defines DC01 (RootDC, CaRoot, Routing) + CM01 (SQL 2022), runs `Install-Lab`, expands disk, configures SQL memory |
| 4. Copy Software | Copies ADK, CM source, prereqs, ODBC, VC++ to CM01, flattens nested folders |
| 5. AD Configuration | Extends AD schema (`extadsch.exe`), creates System Management container with site server permissions |
| 6. Service Accounts | Creates svc-CMPush (Domain Admins), svc-CMNAA (Domain Users), svc-CMAdmin (Domain Admins + RDP) |
| 7. Install Software | Installs VC++ x86/x64, ODBC, MSOLEDB, ADK, ADK WinPE on CM01 |
| 8. Install ConfigMgr | Generates unattended INI, installs Windows features, runs CM setup, validates via WMI |
| 9. Content Share | Creates `E:\ContentShare` with SMB share and NTFS permissions |
| 10. MECM Admin | Adds svc-CMAdmin as MECM Full Administrator via CM PowerShell module |
| 11. Deploy Tools | Copies cc4cm and ApplicationPackager to `C:\Tools\` on CM01 |
| 12. Snapshots | Creates "Deployment-Complete" snapshots on both VMs |
| 13. Summary | Prints connection info and remaining manual console steps |

The script is idempotent -- safe to re-run if it fails partway through. Each step checks whether its work has already been done and skips if so.

## Configuration

All configurable values are in `config.psd1`:

- Lab name, domain name, site code, site name
- Network prefix (default: 192.168.50)
- Admin credentials
- VM sizing (RAM, CPU, disk sizes)
- Software download URLs and versions

Edit `config.psd1` before running the scripts to customize your lab.

## After Deployment

### Service Accounts

The script creates the following service accounts in `OU=Service Accounts,DC=contoso,DC=com`:

| Account | Password | Purpose | Permissions |
|---------|----------|---------|-------------|
| `CONTOSO\svc-CMPush` | `P@ssw0rd!Push1` | Client Push Installation | Domain Admins (local admin on all domain PCs) |
| `CONTOSO\svc-CMNAA` | `P@ssw0rd!NAA1` | Network Access Account | Domain Users only (least privilege) |
| `CONTOSO\svc-CMAdmin` | `P@ssw0rd!Admin1` | MECM admin, cc4cm, RDP | Domain Admins + Remote Desktop Users |

Configure these in the MECM console after deployment:

**Client Push Account:**
Administration > Site Configuration > Sites > right-click site > Client Installation Settings > Client Push Installation > Accounts tab > Add `CONTOSO\svc-CMPush`

**Network Access Account:**
Administration > Site Configuration > Sites > right-click site > Configure Site Components > Software Distribution > Network Access Account tab > Add `CONTOSO\svc-CMNAA`

**Admin Account (`svc-CMAdmin`):**
Use this account to RDP into any lab VM (CM01, DC01, CLIENT01) and run tools like the CM console or Client Center (cc4cm). No console configuration needed -- Domain Admins membership provides full MECM and WinRM access.

> **Note:** The NAA test connection may show "access denied" on C$ -- this is expected. NAA only needs read access to the DP content share, not admin shares. It is intentionally least-privilege.

### Console Configuration

Complete these steps from the CM console on CM01:

1. **Enable Active Directory Forest Discovery:** Administration > Hierarchy Configuration > Discovery Methods > Active Directory Forest Discovery > Enable
2. **Create a Boundary:** Administration > Hierarchy Configuration > Boundaries > Create Boundary > Type: IP Subnet > `192.168.50.0/24`
3. **Create a Boundary Group:** Administration > Hierarchy Configuration > Boundary Groups > Create > Add the boundary above > References tab > add CM01 as site system server
4. **Enable Client Push:** Administration > Site Configuration > Sites > right-click site > Client Installation Settings > Client Push Installation > Enable > check "Automatically install..."

### Content Share

A hidden share on CM01 for application content, drivers, images, and packages (created automatically by the deploy script):

| Detail | Value |
|--------|-------|
| UNC Path | `\\CM01\ContentShare$` |
| Local Path | `E:\ContentShare` (SQL data disk) |
| Full Access | Domain Admins |
| Read Access | Domain Computers, svc-CMNAA |

```
\\CM01\ContentShare$\
    Applications\       # AppPackager content (Vendor\App\Version)
    Drivers\
    Images\
    OperatingSystems\
    Packages\
    Scripts\
    SoftwareUpdates\
```

Configure ApplicationPackager to use this share: **File > Preferences > File Share Root** = `\\CM01\ContentShare$`

### Software Update Point Setup

The SUP role was installed during CM setup. WSUS is running on CM01. Configure update synchronization:

1. **Open SUP Properties:** Administration > Site Configuration > Sites > right-click site > Configure Site Components > Software Update Point
2. **Set sync source:** Synchronize from Microsoft Update
3. **Set sync schedule:** Enable scheduled sync (e.g., every 7 days)

4. **Select Products:** Administration > Site Configuration > Sites > right-click site > Configure Site Components > Software Update Point > Products tab. Start small:
   - Windows 11
   - Windows Server 2025
   - Microsoft Edge
   - Office 365 Client (if using M365)
   - Microsoft Defender Antivirus

5. **Select Classifications:**
   - Critical Updates
   - Security Updates
   - Definition Updates
   - Feature Packs (optional)
   - Update Rollups (optional)

6. **Run initial sync:** Software Library > Software Updates > All Software Updates > right-click > Synchronize Software Updates. First sync takes 15-60 minutes depending on product selection.

7. **Monitor sync:** Monitoring > Software Update Point Synchronization Status. Or check `wsyncmgr.log` on CM01:
   ```powershell
   Get-Content 'C:\Program Files\Microsoft Configuration Manager\Logs\wsyncmgr.log' -Tail 20
   ```

> **Tip:** Keep the initial product list small. Each additional product adds significant sync time and disk usage. You can always add more products later.

### Tools on CM01

ApplicationPackager and Client Center are deployed automatically if found on the host:

| Tool | Path on CM01 |
|------|-------------|
| Application Packager | `C:\Tools\ApplicationPackager\start-apppackager.ps1` |
| Client Center (cc4cm) | `C:\Tools\ClientCenter\SCCMCliCtrWPF.exe` |

Run as `CONTOSO\svc-CMAdmin` for full MECM access. Import the ConfigMgr module before using Application Packager:
```powershell
Import-Module (Join-Path $env:SMS_ADMIN_UI_PATH "..\ConfigurationManager.psd1")
cd MCM:
& C:\Tools\ApplicationPackager\start-apppackager.ps1
```

### Connect to the CM Console

```powershell
# RDP into CM01
Connect-LabVM -ComputerName CM01

# Or use PowerShell remoting
Enter-LabPSSession -ComputerName CM01
```

### Lab Management

```powershell
Stop-Lab -Name HomeLab      # Stop all VMs
Start-Lab -Name HomeLab     # Start all VMs
Remove-Lab -Name HomeLab    # Delete entire lab

# Snapshots
Checkpoint-VM -Name DC01, CM01 -SnapshotName 'MySnapshot'
Restore-VMSnapshot -Name 'Pre-CM-Install' -VMName CM01 -Confirm:$false
```

### Interactive Management with Claude Code

Run Claude Code as Administrator and use `Enter-LabPSSession` to manage the lab interactively. Claude Code can help with ConfigMgr configuration, troubleshooting, and automation tasks.

## Troubleshooting

### Hyper-V requires a reboot after enabling

`Deploy-HomeLab.ps1` does not enable Hyper-V automatically. Enable it manually, reboot, then run the script.

### AutomatedLab DHCP role throws "not implemented yet"

The DHCP role in AutomatedLab is not implemented for all OS versions. The deploy script intentionally omits the DHCP role from DC01. If you need DHCP, configure it manually after deployment:

```powershell
Invoke-LabCommand -ComputerName DC01 -ScriptBlock {
    Install-WindowsFeature DHCP -IncludeManagementTools
    Add-DhcpServerv4Scope -Name 'Lab' -StartRange 192.168.50.100 -EndRange 192.168.50.200 -SubnetMask 255.255.255.0
    Set-DhcpServerv4OptionValue -ScopeId 192.168.50.0 -DnsServer 192.168.50.10 -Router 192.168.50.10 -DnsDomain contoso.com
}
```

### VMs have no internet access

The internal Hyper-V switch does not provide internet. Both VMs are configured with a second NIC on the `Default Switch` which provides NAT internet access. If internet is still not working:

```powershell
# Verify the Default Switch NIC exists
Invoke-LabCommand -ComputerName CM01 -ScriptBlock { Get-NetAdapter }

# Check routing
Invoke-LabCommand -ComputerName CM01 -ScriptBlock { Test-NetConnection -ComputerName 8.8.8.8 }
```

### ODBC Driver 18.6.1.1 breaks ConfigMgr (NULL handling regression)

ODBC Driver 18.6.1.1 has a known regression that causes NULL handling issues with ConfigMgr. This project uses 18.5.2.1 specifically. If you accidentally installed 18.6.1.1, uninstall it and install 18.5.2.1 before running CM setup.

### OS disk too small / out of space on CM01

The default VHDX size may not have enough space for CM + SQL + ADK + content. The deploy script expands the OS disk to 150 GB automatically. If you still run out of space, expand manually:

```powershell
# On the host
Stop-VM -Name CM01
Resize-VHD -Path (Get-VMHardDiskDrive -VMName CM01 | Where-Object ControllerLocation -eq 0).Path -SizeBytes 200GB
Start-VM -Name CM01

# Inside the VM
Invoke-LabCommand -ComputerName CM01 -ScriptBlock {
    $max = (Get-PartitionSupportedSize -DriveLetter C).SizeMax
    Resize-Partition -DriveLetter C -Size $max
}
```

### OS name must include "(Desktop Experience)"

AutomatedLab matches OS names from the ISO exactly. The correct name is `Windows Server 2025 Datacenter Evaluation (Desktop Experience)`, not `Windows Server 2025 Datacenter Evaluation`. If you get an OS-not-found error, list available OS names:

```powershell
Get-LabAvailableOperatingSystem -Path C:\LabSources\ISOs | Select-Object OperatingSystemName
```

### Copy-LabFileItem creates nested folders

When copying a folder like `C:\LabSources\SoftwarePackages\CM\ConfigMgr_2509` to `C:\Install\CM`, AutomatedLab creates `C:\Install\CM\ConfigMgr_2509\` (nested). The deploy script handles this by flattening the folder structure after copy.

### ConfigMgr setup fails

Check the setup log:

```powershell
Invoke-LabCommand -ComputerName CM01 -ScriptBlock {
    Get-Content 'C:\ConfigMgrSetup.log' -Tail 50
}
```

Common causes:
- Missing Windows features (the deploy script installs these automatically)
- SQL collation wrong (must be `SQL_Latin1_General_CP1_CI_AS`)
- Prerequisites not fully downloaded (re-run the script -- it is idempotent)
- ADK not installed or wrong version

### ConfigMgr 2509 is not a built-in AutomatedLab role

AutomatedLab's built-in `ConfigurationManager` role only supports up to version 2203. For 2509, the deploy script handles the full installation manually using unattended setup. This is why CM01 only gets the `SQLServer2022` role in the lab definition.

### SQL memory not configured

ConfigMgr recommends dedicated SQL memory. The deploy script sets SQL Server to 8 GB min/max via `sp_configure`. If you have more host RAM, increase these values in the script.

## File Structure

```
homelab/
    README.md              # This file
    CHANGELOG.md           # Version history
    config.psd1            # All configurable values
    Deploy-HomeLab.ps1     # Single-script deployment (run this)
    lib/
        AutomatedLab/      # Vendored AutomatedLab modules (fork)
```

## Estimated Timelines

| Phase | Duration |
|-------|----------|
| Prerequisites + Downloads (Phases 1-2) | 15-30 minutes |
| Lab Deployment + Software Copy (Phases 3-4) | 30-60 minutes |
| AD + Service Accounts (Phases 5-6) | 2 minutes |
| Software Install + ConfigMgr (Phases 7-8) | 1-3 hours |
| Content Share + Tools + Snapshots (Phases 9-12) | 5 minutes |
| Console Configuration (manual) | 5-10 minutes |
| **Total** | **2-4 hours** |

## License

This project is provided as-is for home lab and educational use.
