# Changelog

All notable changes to WinTuner GUI are documented here.

## [Unreleased]

### Changed
- Development work is performed on the `Test` branch before promotion to `main`.
- Repository documentation and metadata cleanup started.

## [0.10.13] – 2026-09-08

### Added
- Added modular Intune integration with detected-app caching and throttling handling.
- Added persistent WinGet discovery caching and isolated batch workers for large discovery scans.
- Added cancellation support for active WinGet discovery scans.
- Added version comparison tests and further module separation for core, settings, logging, Intune, and WinGet functionality.

### Improved
- Significantly improved Discovered Apps performance and memory usage.
- Reduced repeated WinGet lookups by using persistent positive and zero-result cache entries.
- Improved Graph API reliability with pagination, retry handling, and detected-app caching.
- Improved duplicate handling for discovery search terms.
- Improved first-login reliability by retrying transient connection verification failures.

### Fixed
- Fixed first-login attempts occasionally failing even though authentication succeeded.
- Fixed asynchronous startup update checks failing because BackgroundWorker threads had no PowerShell runspace.
- Fixed manual update checks using the same broken asynchronous path.
- Fixed false "Update available" prompts when the local and GitHub versions were identical.
- Fixed upgrades from legacy single-file releases by automatically bootstrapping required modules and worker files.

## [0.10.12] – 2026-09-01

### Changed
- Updated the application version to `0.10.12`.
- Improved application self-update check logic and status feedback.
- Hardened update-result handling when GitHub version information is incomplete.

## [0.10.11] – 2026-04-13

### Fixed
- `$progressBar`, `$statusLabel`, and `$outputBox` are now script-scoped to prevent null-control failures in `BackgroundWorker` completion callbacks.

## [0.10.10]

### Fixed
- Hardened `Invoke-AsyncOperation` closure handling for progress controls in `RunWorkerCompleted`.

## [0.10.9]

### Fixed
- Progress-bar maximum/reset handling corrected across asynchronous operations.
- App removal now handles missing items gracefully.

## [0.10.8]

### Fixed
- Update-check status feedback improved.
- Update-check button is reliably re-enabled after asynchronous checks.

## [0.10.7]

### Changed
- Improved error handling, security checks and module import guards.

## [0.10.6] – 2026-04-10

### Changed
- Disk version cache is loaded once per session instead of on every cache miss.
- Added 200 ms filter debouncing for update and discovered-app searches.
- Extracted duplicate update loops into `Invoke-AppUpdateBatch`.
- Moved `Get-StringSimilarity` to a top-level function.
- Kept `$script:updateApps` type-consistent when removing entries.
- Capped Graph API pagination at 100 pages.

## [0.10.5]

### Changed
- Improved GUI usability and batch-update summaries.

### Fixed
- Progress-bar handling during batch operations.

## [0.10.4]

### Changed
- Update checks moved to asynchronous execution.
- Improved disconnect timeout handling.
- Removed obsolete code paths.

## [0.10.3]

### Fixed
- Improved error handling and logging consistency.

## [0.10.2] – 2026-04-10

### Fixed
- Successfully updated apps are immediately removed from the update list.

## [0.10.1]

### Fixed
- Synchronized the Remember Me controls between login and Settings.

## [0.10.0] – 2026-04-10

### Added
- Recent-user history with configurable maximum size.
- `Clear-RecentUsers` functionality and Clear History button.

### Changed
- Username field changed to an editable ComboBox.
- Improved connected/disconnected GUI state handling.
- Increased initial form size to reduce clipping.

## [0.9.0] – 2026-04-09

### Changed
- Translated comments to English.
- Adopted PowerShell-approved verbs throughout the script.
- Added a central configuration block for script-scoped state and constants.

## [0.8.1] – 2026-04-09

### Added
- WinGet version disk cache with a six-hour TTL.
- In-memory version cache.
- Clear Cache button.

### Fixed
- UTC cache timestamp parsing using `RoundtripKind`.

## [0.8.0] – 2026-04-09

### Added
- Optional SHA256 verification for self-update downloads.
- Tooltips for key controls.
- Username persistence.

### Changed
- Increased log-box height.

### Fixed
- Reliable SHA256 mismatch detection.

## [0.7.1] – 2026-04-09

### Fixed
- Cache initialization crashes.
- `BackgroundWorker` disposal leak.
- Thread-unsafe logging from background workers.

## [0.6.x and earlier]

### Added
- Global WinForms/AppDomain exception logging.
- Log rotation.
- Discovered-app matching and consolidation.

### Changed
- Improved discovered-app filtering and name normalization.
- Added UI responsiveness handling for long-running scans.
