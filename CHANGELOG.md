# Change Log

All notable changes to this project will be documented in this file. The format is based on [Keep a Changelog](https://keepachangelog.com), and this project adheres to [Semantic Versioning](https://semver.org).

## [3.2.3] - 10-08-2026

### Changed
- Updated the sharedMailboxes (separate entitlements) import script:
	- Prevent pagination errors that can occur with `Get-Mailbox -ResultSize Unlimited`:
	- Use `Get-EXOMailbox` instead of `Get-Mailbox` when retrieving user and shared mailboxes.
	- In the sharedMailboxes (separate entitlements) import script, retrieve user and shared mailboxes separately by filtering on `RecipientTypeDetails`.
	- Add `GrantSendOnBehalfTo` to the requested EXO mailbox properties.


## [3.2.2] - 03-08-2026

### Added
- Import script at `correlateOnly` folder

## [3.2.1] - 15-06-2026

### Added
- NEw Workflow files

## [3.2.0] - 13-04-2026

### Added
- Added `importPermission.ps1` and `importSubPermissions.ps1` scripts for Groups permissions
- Added `importPermission.ps1` and `importSubPermissions.ps1` scripts for Shared Mailboxes (legacy and new separate entitlements)
- Added dynamic permissions support for the new Shared Mailboxes separate entitlements implementation
- Added marker fields in resource scripts to support reliable permission imports

### Changed
- Minor improvements and refinements

### Removed
- Litigation Hold scripts [#29](https://github.com/Tools4everBV/HelloID-Conn-Prov-Target-ExchangeOnline/issues/29)

## [3.1.0] - 15-12-2025
- Added certificate support
- Enhanced README [#24](https://github.com/Tools4everBV/HelloID-Conn-Prov-Target-ExchangeOnline/issues/24)
- Fixed removed user or mailbox check [#22](https://github.com/Tools4everBV/HelloID-Conn-Prov-Target-ExchangeOnline/issues/22)

## [3.0.2] - 18-11-2025

### Added
- GitHub workflow for automated release creation with changelog integration

## [3.0.1] - 18-11-2025

### Added
- Enhanced functionality to update **EmailAddresses (proxy addresses)**. The script now ensures that existing proxy addresses are preserved, and new ones are added with the correct primary (SMTP:) and secondary (smtp:) casing

### Changed
- Updated `update.ps1` to merge existing and new email addresses while preserving all prior addresses
- Modified field mapping to support complex email address updates with proper SMTP prefix handling
- Improved audit logging to show only the properties that were updated rather than full before/after values

## [3.0.0] - 12-12-2024

This is the first release of powershell v2

## [2.0.0] - 31-01-2024

Latest release of powershell v1

### Added

### Changed

### Deprecated

### Removed
