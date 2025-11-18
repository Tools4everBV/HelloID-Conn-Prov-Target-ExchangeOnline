# Change Log

All notable changes to this project will be documented in this file. The format is based on [Keep a Changelog](https://keepachangelog.com), and this project adheres to [Semantic Versioning](https://semver.org).

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
