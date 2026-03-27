# Changelog

## [2.0.0] - 2026-03-26

### Changed
- **AutomatedLab handles CM installation end-to-end.** Removed all band-aid phases (software copy, ADK install, ODBC install, CM setup) from Deploy-HomeLab.ps1. The vendored AutomatedLab fork handles VC++, ODBC, ADK, MSOLEDB, CM prereqs, and CM unattended setup natively via the ConfigurationManager role.
- Script reduced from 1564 to ~890 lines (-43%)
- CM01 now defined with both `SQLServer2022` and `ConfigurationManager` roles
- CLIENT01 deferred until after DC01+CM01 are built (avoids RAM contention during AD/SQL install)
- NAT internet access restricted to CM01 only. DC01 and CLIENT01 are internal-only.

### Fixed (29 bugs total — all in vendored AutomatedLab fork, fixed at the source)
- DHCP role implemented (was "not implemented" placeholder)
- CM local source auto-detection (no hardcoded version URLs for CM 2509+)
- VC++ runtime URLs updated from 2015/2017 to latest (14.50)
- Stale cached VC++ 2015/2017 binaries cleaned up on startup
- Reboot after VC++ install before ODBC/MSOLEDB (3010 pending reboot blocked MSI)
- MSOLEDB install non-fatal (download may fail, SQL has baseline, CM handles upgrade)
- SSRS install/config non-fatal (fails on some Server 2025 configs, optional for CM)
- ADK arguments fixed (single string format, not array)
- Flatten step idempotent (Copy-Item instead of Move-Item)
- SQL Server 2022 added to auto-detection
- Default disk size increased from 50GB to 100GB
- Install-CMSite hardcoded paths replaced with `$VMInstallDirectory` parameter
- `$VMCMBinariesDirectory` and `$VMCMPreReqsDirectory` derived from parameter, not hardcoded
- Case mismatch `CM-Prereqs` vs `CM-PreReqs` normalized
- INI copy destination and path reference use `$VMInstallDirectory`
- Error messages fixed (undefined variables, wrong format placeholders)
- `VMInstallDirectory` passed to `Install-CMSite` via splatted hashtable
- `$WMIZip` typo fixed to `$WMIv2Zip`
- AV exclusion path `C:\InstallCM\` fixed to `C:\Install\CM\`
- Stale ADK/WinPE AV exclusion paths removed
- `Add-LocalGroupMember` missing `-ErrorAction SilentlyContinue` (already-member crash)
- Routing role removed from DC01 (requires 2 NICs, DC01 is internal-only)
- Default Switch defined as Internal (not External)
- MinMemory required when MaxMemory set
- DNS moved to network adapter definition (parameter set conflict)
- Recipe/Ships/Test modules removed from manifest (parse errors)
- Force vendored module overwrite on install
- PSFConfig VC++ URL override after Import-Module
- PSObject serialization: `$deployDebugPath` and `$exePath` cast to `[string]` (Invoke-LabCommand -PassThru returns PSObject, not string)
- Missing `-Function (Get-Command "Import-CMModule")` on 3 Invoke-LabCommand calls in Install-CMSite
- Double TaskEnd in Install-CMSite pre-req checks region

### Added
- CLIENT01 (Windows 11 Enterprise) workstation VM
- svc-CMAdmin as MECM Full Administrator (automated)
- Tools deployment (cc4cm + ApplicationPackager) to CM01
- Deployment-Complete snapshots on all VMs

---

## [1.0.0] - 2026-03-26

### Changed
- Consolidated all 7 numbered scripts into a single `Deploy-HomeLab.ps1`
- Idempotent design -- safe to re-run if it fails partway through

### Removed
- All 7 numbered scripts (merged into single Deploy-HomeLab.ps1)

---

## [0.2.1] - 2026-03-26

### Added
- `svc-CMAdmin` account — Domain Admin + Remote Desktop Users for interactive logon, CM console, and cc4cm
- Quick Start section updated with all 7 steps including post-deployment console configuration

### Fixed
- `06-Create-ServiceAccounts.ps1` — replaced splatting `@{}` with explicit parameters (splatting breaks through PSRemoting layers)

---

## [0.2.0] - 2026-03-26

### Added
- `05-Configure-AD.ps1` — AD schema extension (`extadsch.exe`) and System Management container with site server permissions
- `06-Create-ServiceAccounts.ps1` — Client Push (`svc-CMPush`, Domain Admins) and NAA (`svc-CMNAA`, Domain Users only) service accounts
- Service account configuration in `config.psd1` (names, passwords, group membership)
- Post-deployment console configuration steps in README (discovery, boundaries, boundary groups, client push)
- NAA least-privilege documentation (Domain Users only, no admin shares needed)

---

## [0.1.0] - 2026-03-25

### Added
- Initial release
- 4-script modular deployment (prerequisites, downloads, infrastructure, ConfigMgr)
- Centralized config.psd1 for all lab parameters
- Windows Server 2025 + SQL Server 2022 + ConfigMgr 2509 support
- Automated ADK offline layout creation
- ODBC 18.5.2.1 compatibility (avoids 18.6.1.1 regression)
- VC++ 14.50 runtime installation
- Internet access via Default Switch NAT
- Comprehensive troubleshooting guide
