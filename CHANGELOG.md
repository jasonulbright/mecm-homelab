# Changelog

## [2.0.0] - 2026-03-26

### Changed
- **AutomatedLab handles CM installation end-to-end.** Removed all band-aid phases (software copy, ADK install, ODBC install, CM setup) from Deploy-HomeLab.ps1. The vendored AutomatedLab fork handles VC++, ODBC, ADK, MSOLEDB, CM prereqs, and CM unattended setup natively via the ConfigurationManager role.
- Script reduced from 1564 to ~890 lines (-43%)
- CM01 now defined with both `SQLServer2022` and `ConfigurationManager` roles
- CLIENT01 deferred until after DC01+CM01 are built (avoids RAM contention during AD/SQL install)
- NAT internet access restricted to CM01 only. DC01 and CLIENT01 are internal-only.

### Fixed (in vendored AutomatedLab fork)
- DHCP role implemented (was "not implemented" placeholder)
- CM local source auto-detection (no hardcoded version URLs for CM 2509+)
- VC++ runtime URLs updated from 2015/2017 to latest (14.50)
- Reboot after VC++ install before ODBC/MSOLEDB (3010 pending reboot blocked MSI)
- MSOLEDB 1603 treated as non-fatal (already installed by SQL setup)
- ADK arguments fixed (single string format, not array)
- Flatten step idempotent (Copy-Item instead of Move-Item)
- SQL Server 2022 added to auto-detection
- setup.exe recursive path search (SMSSETUP\BIN\X64)
- Default disk size increased from 50GB to 100GB

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
