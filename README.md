# MECM Home Lab Deployment

Automated deployment of a Microsoft Endpoint Configuration Manager (MECM) home lab using a [vendored fork of AutomatedLab](https://github.com/jasonulbright/AutomatedLab) and Hyper-V. No external dependencies -- everything is self-contained.

Deploys a fully functional ConfigMgr 2509 primary site with Windows Server 2025, SQL Server 2022, Active Directory, Certificate Services, and a Windows 11 workstation.

## Hardware Requirements

| Resource | Minimum | Recommended |
|----------|---------|-------------|
| RAM      | 32 GB   | 64 GB       |
| Disk     | 300 GB free (SSD/NVMe) | 500 GB+ |
| CPU      | 4 cores | 8+ cores    |
| OS       | Windows 10/11 Pro (Hyper-V capable) | Windows 11 Pro |

## Lab Architecture

| VM | Role | IP | RAM | vCPU | OS | Internet |
|----|------|----|-----|------|----|----------|
| DC01 | Domain Controller, Root CA | 192.168.50.10 | 1-2 GB | 2 | Windows Server 2025 | No |
| CM01 | SQL Server 2022, ConfigMgr 2509 | 192.168.50.20 | 4-12 GB | 4 | Windows Server 2025 | Yes (NAT) |
| CLIENT01 | Workstation (managed client) | 192.168.50.100 | 2-4 GB | 2 | Windows 11 Enterprise | No |

Only CM01 has internet access (via Hyper-V Default Switch NAT). DC01 and CLIENT01 are internal-only.

## Software Prerequisites

Download these before running the script:

| Software | Location | Download URL |
|----------|----------|-------------|
| Windows Server 2025 Evaluation ISO | `C:\LabSources\ISOs\` | https://www.microsoft.com/en-us/evalcenter/evaluate-windows-server-2025 |
| SQL Server 2022 Evaluation ISO | `C:\LabSources\ISOs\` | https://www.microsoft.com/en-us/evalcenter/download-sql-server-2022 |
| Windows 11 Enterprise Evaluation ISO | `C:\LabSources\ISOs\` | https://www.microsoft.com/en-us/evalcenter/evaluate-windows-11-enterprise |
| ConfigMgr 2509 Baseline (extract with 7-Zip) | `C:\LabSources\SoftwarePackages\CM\` | https://www.microsoft.com/en-us/evalcenter/download-microsoft-endpoint-configuration-manager |
| Windows ADK (adksetup.exe) | `C:\LabSources\SoftwarePackages\ADK\` | https://go.microsoft.com/fwlink/?linkid=2289980 |
| Windows PE add-on (adkwinpesetup.exe) | `C:\LabSources\SoftwarePackages\ADKPE\` | https://go.microsoft.com/fwlink/?linkid=2289981 |

ODBC Driver 18.5.x and VC++ runtimes are downloaded automatically by AutomatedLab during deployment.

OS edition names are detected automatically from your ISOs. The script matches against wildcard filters in `config.psd1` (`ServerOSFilter`, `ClientOSFilter`) so any edition (Standard, Datacenter, Evaluation, retail) works.

## Quick Start

Enable Hyper-V first (reboot required). Download the ISOs and software above. **Change the passwords in `config.psd1`**, then run as **Administrator**:

```powershell
.\Deploy-HomeLab.ps1
```

To remove an existing lab and redeploy:
```powershell
.\Deploy-HomeLab.ps1 -RemoveExisting
```

If the lab already exists and `-RemoveExisting` is not specified, the script imports the existing lab and continues with the remaining idempotent phases.

## What the Script Does

`Deploy-HomeLab.ps1` handles prerequisites and lab definition, then delegates to AutomatedLab for the heavy lifting:

| Phase | Description |
|-------|-------------|
| 1. Prerequisites | Checks Hyper-V, host RAM, installs vendored AutomatedLab, creates LabSources, verifies ISOs, detects OS editions from WIM |
| 2. Downloads | Creates ADK/PE offline layouts if needed |
| 3. Lab Deployment | Defines DC01 + CM01 (with ConfigurationManager role) + CLIENT01, runs `Install-Lab`. AutomatedLab handles AD, CA, SQL, VC++, ODBC, ADK, SQL memory, and CM installation. CLIENT01 deferred until DC+CM are up. |
| 4. Service Accounts | Creates svc-CMPush, svc-CMNAA, svc-CMAdmin (with error recovery) |
| 5. Content Share | Creates `E:\ContentShare` with SMB share on CM01 (with error recovery) |
| 6. MECM Admin | Adds svc-CMAdmin as MECM Full Administrator (with error recovery) |
| 7. Deploy Tools | Copies cc4cm and ApplicationPackager to CM01 (optional, with error recovery) |
| 8. Snapshots | Creates "Deployment-Complete" snapshots on all VMs |

The script is idempotent -- safe to re-run if it fails partway through. Phases 4-7 have try/catch blocks with recovery guidance if a VM is unresponsive.

## Configuration

All configurable values are in `config.psd1`:

| Setting | Description |
|---------|-------------|
| Lab name, domain, site code | Core lab identity |
| VM sizing (RAM, CPU, disks) | Per-VM resource allocation |
| Admin credentials | Domain admin account |
| Service account names/passwords | 3 MECM service accounts |
| OS filters | Wildcard patterns matched against ISO contents |
| Software URLs | ODBC, VC++ download locations |

**Security**: Default passwords are published in source control. The script warns loudly at startup if defaults are still in use. Change them before deploying to any network.

## After Deployment

### Service Accounts

Created automatically in `OU=Service Accounts,DC=contoso,DC=com`:

| Account | Purpose | Permissions |
|---------|---------|-------------|
| `CONTOSO\svc-CMPush` | Client Push Installation | Domain Admins |
| `CONTOSO\svc-CMNAA` | Network Access Account | Domain Users only |
| `CONTOSO\svc-CMAdmin` | MECM admin, cc4cm, RDP | Domain Admins + Remote Desktop Users |

Passwords are set in `config.psd1`. Change them before deployment.

Configure in the MECM console:
- **Client Push**: Administration > Site Config > Sites > Client Installation Settings > Client Push > Accounts > Add `svc-CMPush`
- **NAA**: Administration > Site Config > Sites > Configure Site Components > Software Distribution > NAA > Add `svc-CMNAA`

### Console Configuration

Complete from the CM console on CM01:
1. Enable Active Directory Forest Discovery
2. Create Boundary (IP subnet `192.168.50.0/24`)
3. Create Boundary Group (add boundary + CM01 as site system)
4. Enable Client Push (add svc-CMPush, check "Automatically install...")

### Content Share

Created automatically on CM01:

| Detail | Value |
|--------|-------|
| UNC Path | `\\CM01\ContentShare$` |
| Local Path | `E:\ContentShare` |
| Full Access | Domain Admins |
| Read Access | Domain Computers, svc-CMNAA |

### Tools on CM01

| Tool | Path |
|------|------|
| Application Packager | `C:\Tools\ApplicationPackager\start-apppackager.ps1` |
| Client Center (cc4cm) | `C:\Tools\ClientCenter\SCCMCliCtrWPF.exe` |

### Lab Management

```powershell
Connect-LabVM -ComputerName CM01         # RDP
Enter-LabPSSession -ComputerName CM01    # PS remoting
Stop-Lab -Name HomeLab                   # Stop all VMs
Start-Lab -Name HomeLab                  # Start all VMs
Remove-Lab -Name HomeLab                 # Delete entire lab
```

## Troubleshooting

### Hyper-V requires a reboot after enabling

The script does not enable Hyper-V automatically. Enable it manually, reboot, then run the script.

### ODBC Driver 18.6.1.1 breaks ConfigMgr

ODBC 18.6.1.1 has a NULL handling regression that breaks CM. The vendored AutomatedLab fork installs 18.5.2.1 specifically.

### OS disk too small

The default OS disk is expanded to 150 GB (configurable via `CM.OSDiskSize` in config.psd1). If CM01 still runs out of space:
```powershell
Stop-VM -Name CM01
Resize-VHD -Path (Get-VMHardDiskDrive -VMName CM01 | Where-Object ControllerLocation -eq 0).Path -SizeBytes 200GB
Start-VM -Name CM01
Invoke-LabCommand -ComputerName CM01 -ScriptBlock {
    Resize-Partition -DriveLetter C -Size (Get-PartitionSupportedSize -DriveLetter C).SizeMax
}
```

### ConfigMgr setup fails

Check the setup log:
```powershell
Invoke-LabCommand -ComputerName CM01 -ScriptBlock {
    Get-Content 'C:\ConfigMgrSetup.log' -Tail 50
}
```

### OS edition not found

If the script fails with "No server OS found matching filter", run this to see available editions in your ISOs:
```powershell
Import-Module AutomatedLab
Get-LabAvailableOperatingSystem -Path C:\LabSources\ISOs
```
Then update `ServerOSFilter` or `ClientOSFilter` in `config.psd1` to match.

## File Structure

```
homelab/
    README.md                  # This file
    CHANGELOG.md               # Version history
    config.psd1                # All configurable values (change passwords here)
    Deploy-HomeLab.ps1         # Single-script deployment
    Update-VendoredModules.ps1 # Copy fork modules into lib/
    example_install.log        # Sample deployment output (~1h 46m run)
    Tests/
        Deploy-HomeLab.Tests.ps1  # Pester tests (29 tests)
    lib/
        AutomatedLab/          # Vendored AutomatedLab fork (29 bug fixes for CM 2509)
```

## Tests

```powershell
# Homelab wrapper tests (29 tests -- config, passwords, structure, OS filters)
Invoke-Pester ./Tests/

# AutomatedLab fork tests (172 tests -- CM install, DHCP, paths, INI, prereqs)
Invoke-Pester c:\projects\AutomatedLab\Tests\CommunityFork.Tests.ps1
```

## Estimated Timelines

| Phase | Duration |
|-------|----------|
| Prerequisites + Downloads | 10-15 minutes |
| Lab Deployment (DC + SQL + CM install) | 45-60 minutes |
| Service Accounts + Content Share + Tools | 5 minutes |
| CLIENT01 Deployment | 10-15 minutes |
| Console Configuration (manual) | 5-10 minutes |
| **Total** | **~1.5 hours** (observed on Intel Ultra 7 265K, 32GB RAM, NVMe) |

## License

This project is provided as-is for home lab and educational use.
