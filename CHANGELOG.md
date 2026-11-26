# Change Log

All notable changes to this project will be documented in this file. The format is based on [Keep a Changelog](https://keepachangelog.com), and this project adheres to [Semantic Versioning](https://semver.org).

## [4.0.0] - 26-11-2026

### Breaking changes
- **Distribution groups and mail-enabled security groups are now referenced by `ExchangeObjectId` instead of `Guid`.** This applies to the permission listing (`permissions/groups/permissions.ps1`), the imports and the dynamic sub permissions. In practice both values are the same, so existing entitlements keep working: `grantPermission.ps1`, `revokePermission.ps1` and the sub permissions accept both values (`Identity` takes either). See the README for the upgrade notes.
- Shared mailbox SendAs permissions in `importSubPermissions.ps1` are now retrieved per shared mailbox on the `Guid`, instead of one query grouped on the mailbox name.

### Fixed
- Permission imports with more than 500 accounts per permission: every chunk is now its own object. Before, all chunks contained the accounts of the last chunk (`importPermission.ps1` and `importSubPermissions.ps1` for groups and shared mailboxes).
- Shared mailbox import of SendAs: shared mailboxes with the same name no longer receive each other's permissions.
- Shared mailbox import of SendOnBehalf: a trustee that cannot be resolved is logged with a warning instead of silently skipped; other errors stop the import, so there is no incomplete list. Replaced `Get-Mailbox` (not imported) with `Get-EXOMailbox`.
- Mail-enabled security groups (`permissions/groups/permissions.ps1`) and distribution groups in `resources/groups` are retrieved with `ResultSize Unlimited` (the default is 1000).
- `permissions/groups/importPermission.ps1`: `$commands` was not defined, which is now fixed.
- Granting a shared mailbox permission (FullAccess, SendAs) that the account already has is a success (skipped); revoking a permission that is already gone is skipped.
- Resources (`resources/groups`, `resources/sharedMailboxes`): a half created object (the attributes were not set) is recovered in the next run via its e-mail address, instead of failing on "already exists".
- Error messages in the audit log are readable (`message Error code: code`) instead of raw JSON (`Resolve-ExchangeOnlineError` parsed the wrong object).
- The same error parsing fix is applied to the out of office permission scripts.
- Counts in log lines are correct for a single object in Windows PowerShell 5.1.
- README fully updated: features, lifecycle actions (imports, sub permissions, out of office, regional configuration, new separate entitlements), module requirements (Windows PowerShell 5.1, install for all users), certificate-based authentication, correlation and the used commands.

## [3.5.0] - 09-11-2026

### Added
- Added out of office permission (`permissions/outOfOffice`): grant disables the out of office reply, revoke enables it with a configurable message.

### Changed
- Added the HelloID Icon URL to the README.

## [3.4.1] - 03-11-2026

### Fixed
- Improved function for certificate based connection to Entra ID.

## [3.4.0] - 19-10-2026

### Added
- Added Microsoft Graph integration using the existing certificate authentication to retrieve Entra ID users.
- Added Entra ID personnel numbers (`employeeId`) to account imports in `import.ps1` and `correlateOnly/import.ps1`, including pagination of Entra ID users.

### Changed
- Updated `create.ps1` and `correlateOnly/create.ps1` to correlate against Entra ID user attributes, including `onPremisesExtensionAttributes`, and find the mailbox by `ExternalDirectoryObjectId`.
- Updated account imports to match mailboxes to Entra ID users by `ExternalDirectoryObjectId`, while retaining the mailbox `Guid` as the account reference.
- Updated the README with Entra ID correlation and account import configuration. The app registration now also requires the Microsoft Graph application permission `User.Read.All` with admin consent.

## [3.3.0] - 19-08-2026

### Fixed
- Fallback sendOnBehalfUser failed. Change from `Get-Mailbox` to `Get-EXOMailbox` and added `-RecipientTypeDetails UserMailbox`.

### Removed
- Removed Client Secret support from the repository. Make sure to also update `configuration.json` when updating the connector.

## [3.2.3] - 10-08-2026

### Changed
- Updated the sharedMailboxes import scripts:
	- Prevent pagination errors that can occur with `Get-Mailbox -ResultSize Unlimited`:
	- Use `Get-EXOMailbox` instead of `Get-Mailbox` when retrieving user and shared mailboxes.
	- In the sharedMailboxes import scripts, retrieve user and shared mailboxes separately by filtering on `RecipientTypeDetails`.
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
