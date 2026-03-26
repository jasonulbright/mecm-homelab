# MECM Home Lab Deployment

Automated deployment of a Microsoft Endpoint Configuration Manager (MECM) home lab using [AutomatedLab](https://automatedlab.org/) and Hyper-V.

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

The ODBC Driver and VC++ runtimes are downloaded automatically by script 02.

## Quick Start

All scripts must be run as **Administrator**.

```powershell
# Step 1: Install Hyper-V, AutomatedLab, create folder structure
.\01-Install-Prerequisites.ps1
# (reboot if prompted, then run again)

# Step 2: Create ADK offline layouts, download CM prereqs and runtimes
.\02-Download-Offline.ps1

# Step 3: Deploy DC01 + CM01 VMs with AD, CA, SQL
.\03-Deploy-Infrastructure.ps1
# (~30-60 minutes)

# Step 4: Install ADK, prerequisites, and ConfigMgr 2509
.\04-Install-ConfigMgr.ps1
# (~1-3 hours)

# Step 5: Extend AD schema, configure System Management container
.\05-Configure-AD.ps1

# Step 6: Create service accounts (client push + NAA)
.\06-Create-ServiceAccounts.ps1

# Step 7: Open the CM console on CM01 and configure:
#   - Active Directory Forest Discovery (enable)
#   - Boundary (IP subnet 192.168.50.0/24)
#   - Boundary Group (add boundary + CM01 as site system)
#   - Client Push Installation (enable, add svc-CMPush account)
#   - Software Distribution > NAA (add svc-CMNAA account)

# Step 8: Create content share on CM01 (see below)
# Step 9: Configure Software Update Point (see below)
```

## What Each Script Does

### 01-Install-Prerequisites.ps1

- Enables Hyper-V (prompts for reboot if needed)
- Installs the AutomatedLab PowerShell module from PSGallery
- Creates the `C:\LabSources` folder structure with subdirectories for CM, ADK, ODBC, and VC++ runtimes
- Checks for required ISOs and shows a download checklist with URLs
- Validates host RAM is sufficient

### 02-Download-Offline.ps1

- Creates ADK offline layout using `adksetup.exe /quiet /layout` (required because VMs may not have reliable internet during ADK install)
- Creates ADK WinPE offline layout using `adkwinpesetup.exe /quiet /layout`
- Downloads CM prerequisites using `setupdl.exe` from the CM source
- Downloads ODBC Driver 18.5.2.1 MSI (NOT 18.6.1.1 -- see troubleshooting)
- Downloads VC++ 14.50 x64 and x86 runtimes
- Shows status of what is already downloaded vs. what needs downloading
- Safe to re-run (skips items that already exist)

### 03-Deploy-Infrastructure.ps1

- Creates an AutomatedLab definition with an internal Hyper-V switch
- Defines DC01 with roles: RootDC, CaRoot, Routing (no DHCP -- see troubleshooting)
- Defines CM01 with SQL Server 2022 role (no ConfigurationManager role -- handled manually in script 04)
- Adds a Default Switch NIC to both VMs for internet access
- Runs `Install-Lab` to create and configure VMs (~30-60 min)
- Expands CM01 OS disk to 150 GB and extends the C: partition
- Configures SQL Server memory (8 GB min/max)
- Copies ADK layouts, CM source, prerequisites, ODBC, and VC++ runtimes to CM01
- Flattens nested folder paths created by `Copy-LabFileItem`
- Runs AD schema extension (`extadsch.exe`)
- Creates "Pre-CM-Install" snapshots on both VMs

### 04-Install-ConfigMgr.ps1

- Imports the existing lab
- Installs VC++ 14.50 runtimes (x86 + x64) on CM01
- Installs ODBC Driver 18.5.2.1 on CM01
- Installs MSOLEDB 19 from CM prerequisites
- Installs ADK with DeploymentTools and UserStateMigrationTool features
- Installs ADK WinPE add-on
- Installs required Windows Server features (IIS, BITS, WSUS, .NET, etc.)
- Generates an unattended setup INI file
- Runs `setup.exe /SCRIPT` for unattended CM installation
- Validates installation by checking SMS_Site WMI class, console presence, and SMS_EXECUTIVE service
- Creates "Post-CM-Install" snapshot

## Configuration

All configurable values are in `config.psd1`:

- Lab name, domain name, site code, site name
- Network prefix (default: 192.168.50)
- Admin credentials
- VM sizing (RAM, CPU, disk sizes)
- Software download URLs and versions

Edit `config.psd1` before running the scripts to customize your lab.

## After Deployment

### Step 5: Configure AD for ConfigMgr

```powershell
.\05-Configure-AD.ps1
```

Extends the Active Directory schema for ConfigMgr (adds SMS classes) and creates the System Management container with Full Control for the CM01 computer account. This allows the site server to publish site information to AD.

### Step 6: Create Service Accounts

```powershell
.\06-Create-ServiceAccounts.ps1
```

Creates the following service accounts in `OU=Service Accounts,DC=contoso,DC=com`:

| Account | Password | Purpose | Permissions |
|---------|----------|---------|-------------|
| `CONTOSO\svc-CMPush` | `P@ssw0rd!Push1` | Client Push Installation | Domain Admins (local admin on all domain PCs) |
| `CONTOSO\svc-CMNAA` | `P@ssw0rd!NAA1` | Network Access Account | Domain Users only (least privilege) |
| `CONTOSO\svc-CMAdmin` | `P@ssw0rd!Admin1` | MECM admin, cc4cm, RDP | Domain Admins + Remote Desktop Users |

After running the script, configure these in the MECM console:

**Client Push Account:**
Administration > Site Configuration > Sites > right-click site > Client Installation Settings > Client Push Installation > Accounts tab > Add `CONTOSO\svc-CMPush`

**Network Access Account:**
Administration > Site Configuration > Sites > right-click site > Configure Site Components > Software Distribution > Network Access Account tab > Add `CONTOSO\svc-CMNAA`

**Admin Account (`svc-CMAdmin`):**
Use this account to RDP into any lab VM (CM01, DC01, CLIENT01) and run tools like the CM console or Client Center (cc4cm). No console configuration needed — Domain Admins membership provides full MECM and WinRM access.

> **Note:** The NAA test connection may show "access denied" on C$ — this is expected. NAA only needs read access to the DP content share, not admin shares. It is intentionally least-privilege.

### Step 7: Console Configuration

Complete these steps from the CM console on CM01:

1. **Enable Active Directory Forest Discovery:** Administration > Hierarchy Configuration > Discovery Methods > Active Directory Forest Discovery > Enable
2. **Create a Boundary:** Administration > Hierarchy Configuration > Boundaries > Create Boundary > Type: IP Subnet > `192.168.50.0/24`
3. **Create a Boundary Group:** Administration > Hierarchy Configuration > Boundary Groups > Create > Add the boundary above > References tab > add CM01 as site system server
4. **Enable Client Push:** Administration > Site Configuration > Sites > right-click site > Client Installation Settings > Client Push Installation > Enable > check "Automatically install..."

### Step 8: Content Share

A hidden share on CM01 for application content, drivers, images, and packages:

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

Created automatically by `03-Deploy-Infrastructure.ps1`. Configure ApplicationPackager to use this share: **File > Preferences > File Share Root** = `\\CM01\ContentShare$`

### Step 9: Software Update Point Setup

The SUP role was installed during CM setup (script 04). WSUS is running on CM01. Configure update synchronization:

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

### Deploying Tools to CM01

ApplicationPackager and Client Center are pre-installed:

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

This is expected. Script 01 will prompt you. After rebooting, run `01-Install-Prerequisites.ps1` again to continue.

### AutomatedLab DHCP role throws "not implemented yet"

The DHCP role in AutomatedLab is not implemented for all OS versions. Script 03 intentionally omits the DHCP role from DC01. If you need DHCP, configure it manually after deployment:

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

The default VHDX size may not have enough space for CM + SQL + ADK + content. Script 03 expands the OS disk to 150 GB automatically. If you still run out of space, expand manually:

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

When copying a folder like `C:\LabSources\SoftwarePackages\CM\ConfigMgr_2509` to `C:\Install\CM`, AutomatedLab creates `C:\Install\CM\ConfigMgr_2509\` (nested). Script 03 handles this by flattening the folder structure after copy.

### ConfigMgr setup fails

Check the setup log:

```powershell
Invoke-LabCommand -ComputerName CM01 -ScriptBlock {
    Get-Content 'C:\ConfigMgrSetup.log' -Tail 50
}
```

Common causes:
- Missing Windows features (script 04 installs these automatically)
- SQL collation wrong (must be `SQL_Latin1_General_CP1_CI_AS`)
- Prerequisites not fully downloaded (re-run script 02)
- ADK not installed or wrong version

### ConfigMgr 2509 is not a built-in AutomatedLab role

AutomatedLab's built-in `ConfigurationManager` role only supports up to version 2203. For 2509, script 04 handles the full installation manually using unattended setup. This is why CM01 only gets the `SQLServer2022` role in script 03.

### SQL memory not configured

ConfigMgr recommends dedicated SQL memory. Script 03 sets SQL Server to 8 GB min/max via `sp_configure`. If you have more host RAM, increase these values in the script.

## File Structure

```
homelab/
    README.md                    # This file
    CHANGELOG.md                 # Version history
    config.psd1                  # All configurable values
    01-Install-Prerequisites.ps1 # Host: Hyper-V, AutomatedLab, LabSources
    02-Download-Offline.ps1      # Host: ADK layouts, CM prereqs, runtimes
    03-Deploy-Infrastructure.ps1 # AutomatedLab: VMs, AD, CA, SQL
    04-Install-ConfigMgr.ps1     # CM01: ADK, prereqs, CM unattended install
    05-Configure-AD.ps1          # DC01: AD schema extension, System Management container
    06-Create-ServiceAccounts.ps1# DC01: Client push + NAA + admin service accounts
    07-Create-ContentShare.ps1   # CM01: Content share for apps, drivers, images
```

## Estimated Timelines

| Phase | Duration |
|-------|----------|
| Prerequisites + Downloads (01-02) | 15-30 minutes |
| Infrastructure Deployment (03) | 30-60 minutes |
| ConfigMgr Installation (04) | 1-3 hours |
| AD + Service Accounts + Share (05-07) | 2 minutes |
| Console Configuration (08-09) | 5-10 minutes |
| **Total** | **2-4 hours** |

## License

This project is provided as-is for home lab and educational use.
