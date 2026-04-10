# Changelog

## [2.3.0] - 2026-04-10

### Added
- **CLIENT01 OS disk expansion** -- configurable `OSDiskSize` in config.psd1 (default 150GB). Expands VHDX and extends C: partition inside the VM. Matches the existing CM01 disk expansion pattern.
- **CLIENT01 re-add on rerun** -- when CLIENT01 is missing from the AutomatedLab definition (partial rerun), the script re-adds it with correct NIC config instead of skipping with a warning.

### Changed
- **Pester test updated** -- `OSDiskSize` added to required Client VM properties.

---

## [2.2.1] - 2026-03-31

### Changed
- **SQL Server download instructions** -- clarified that the evaluation download link provides an EXE bootstrapper, not a direct ISO. Run it and select "Download Media" to retrieve the actual ISO.

---

## [2.2.0] - 2026-03-29

### Fixed
- **CM install timeout** -- set `Timeout_ConfigurationManagerInstallation` to 120 minutes before `Install-Lab` (default 60 was insufficient)
- **Log file close order** -- "Log saved to" message now prints before closing the StreamWriter (was attempting to write through the proxy after stream disposal)
- **ODBC MSI on disk** -- fixed misnamed `ODBC` flat file to `ODBC/msodbcsql.msi` directory structure
- **MSOLEDB MSI on disk** -- same fix for `MSOLEDB` flat file

### Changed
- Vendored AutomatedLab fork synced with all v5.60.3-community fixes (ODBC path, validation retry, SNAC/SSRS error suppression, Write-PSFMessage hang)

## [2.1.0] - 2026-03-28

### Security
- **Default password warning** -- script warns loudly at startup if any passwords in config.psd1 are still defaults. Passwords are published in source control and must be changed before deployment.
- **Service account password injection fixed** -- passwords are now passed via `-ArgumentList` instead of string-interpolated into a here-string script. Passwords containing `'` or `$` no longer break the generated script.

### Changed
- **Dynamic OS resolution** -- OS edition names (e.g., "Windows Server 2025 Datacenter Evaluation") are now detected automatically from ISO WIM contents via `Get-LabAvailableOperatingSystem`. Wildcard filters in config.psd1 (`ServerOSFilter`, `ClientOSFilter`) match any edition.
- **Non-interactive** -- removed `Read-Host` prompt when lab already exists. Script now imports and continues (idempotent). Use `-RemoveExisting` to recreate.
- **SQL memory handled by AutomatedLab** -- removed duplicate SQL memory config from wrapper (Phase 3.4). The fork already configures SQL memory before CM setup.
- **Manifest patching removed** -- vendored manifest is pre-patched. No regex replacement at runtime.
- **Elapsed time displayed** in completion banner.

### Added
- `Update-VendoredModules.ps1` -- copies built modules from the AutomatedLab fork into `lib/AutomatedLab/`
- `Tests/Deploy-HomeLab.Tests.ps1` -- 29 Pester tests covering config structure, password security, OS filters, script structure, and vendored module integrity
- try/catch with recovery guidance on Phases 4-7 (service accounts, content share, MECM admin, tools)
- `ServerOSFilter` and `ClientOSFilter` in config.psd1

### Fixed
- **PSObject serialization on SMS_EXECUTIVE check** -- `Invoke-LabCommand -PassThru` result cast with `[bool]($result | Select-Object -First 1)` to handle PSObject wrapping
- **AV exclusion SQL path** -- `MSSQL14.MSSQLSERVER` (SQL 2017) updated to `MSSQL16.MSSQLSERVER` (SQL 2022) in AutomatedLab fork
- **CLIENT01 Install-Lab crash** -- second `Install-Lab` call for CLIENT01 wrapped in try/catch (CM update validation threw on null update packages)
- **DLL lock on module copy** -- `Remove-Item` failure on locked DLLs caught with try/catch, falls back to overwrite-in-place

### Fixed (AutomatedLab fork -- vendored)
- **Update-CMSite null array crash** -- graceful skip when no update packages synced
- **Site-already-installed skip** -- prevents redundant 30-min update polling on re-runs
- **SSRS error noise suppressed** -- `-ErrorAction SilentlyContinue` on SSRS install/config
- **Version check suppressed** -- forced `DisableVersionCheck = $true` overriding persisted values
- **Unnamed activity labels** -- 5 `Invoke-LabCommand` calls given descriptive `-ActivityName`

---

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
