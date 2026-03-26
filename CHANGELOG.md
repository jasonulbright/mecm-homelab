# Changelog

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
