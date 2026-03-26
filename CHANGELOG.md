# Changelog

## [1.0.0] - 2026-03-26

### Changed
- Consolidated all 7 numbered scripts into a single `Deploy-HomeLab.ps1`
- Script runs all phases end-to-end: prerequisites, downloads, lab deployment, AD config, service accounts, CM install, content share, tools, snapshots
- Idempotent design -- safe to re-run if it fails partway through (each step checks if work is already done)
- Hyper-V check now errors with instructions instead of attempting to enable (requires reboot)
- Service account creation uses a script file copied to DC01 instead of inline splatting (avoids remoting serialization issues)

### Added
- Phase 10: Adds svc-CMAdmin as MECM Full Administrator via CM PowerShell module
- Phase 11: Deploys cc4cm and ApplicationPackager to `C:\Tools\` on CM01
- ConfigMgr install now checks for SMS_EXECUTIVE service to skip if already installed
- CM provider readiness check with retry loop before adding admin user

### Removed
- `01-Install-Prerequisites.ps1` (merged into Phase 1)
- `02-Download-Offline.ps1` (merged into Phase 2)
- `03-Deploy-Infrastructure.ps1` (merged into Phases 3-4)
- `04-Install-ConfigMgr.ps1` (merged into Phases 7-8)
- `05-Configure-AD.ps1` (merged into Phase 5)
- `06-Create-ServiceAccounts.ps1` (merged into Phase 6)
- `07-Create-ContentShare.ps1` (merged into Phase 9)

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
