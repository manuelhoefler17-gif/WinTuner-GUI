# Changelog

All notable changes to this project will be documented in this file.

## [0.10.11] – 2026-04-13

### Fixed
- `$progressBar`, `$statusLabel`, and `$outputBox` are now declared as `$script:`-scoped variables throughout the entire script. This fixes a fatal `The property 'Style' cannot be found on this object` error that occurred in the `RunWorkerCompleted` callback of `Invoke-AsyncOperation`, where the BackgroundWorker dispatcher invoked the handler outside the normal PowerShell variable scope, causing those variables to resolve as `$null`.

## [0.10.6] – 2026-04-10

### Changed
- Disk version cache is now loaded once at script start instead of on every cache miss (`$script:diskCache` / `$script:diskCacheLoaded`). Cache-clear also resets these variables.
- Filter debouncing (200ms timer) added for `updateFilterBox` and `discoveredAppSearchBox` to prevent lag with 500+ apps.
- Duplicate update-loop code extracted into `Invoke-AppUpdateBatch` helper function; both `updateSelectedButton` and `updateAllButton` handlers now delegate to it.
- `Get-StringSimilarity` moved from nested definition inside `$scanDiscoveredButton.Add_Click` to a top-level function.
- `$script:updateApps` type consistency fixed: removal now uses `List.Remove()` instead of array conversion to preserve the `[List[object]]` type.
- Graph API pagination loop capped at 100 pages to prevent runaway fetches in very large tenants.

## [0.10.2] - 2026-04-10

### Fixed
- Updated apps are now immediately removed from the update list after a successful update, instead of waiting for the background "Search Updates" to finish. This applies to both "Update Checked Apps" and "Update All".

## [Unreleased]

### Changed
- README improvements: fixed WinTuner module hyperlink and added a "Recommended Pre-Flight Checks" section with quick environment validation commands.

## [0.10.0] – 2026-04-10

### Added
- `Add-RecentUser` function: saves used UPNs to settings (max 3, configurable via `MaxRecentUsers`).
- `Clear-RecentUsers` function: clears the saved login history.
- "Clear History" button (🗑 Clear History) next to the username ComboBox, visible only when logged out.
- Settings now persist `RecentUsers` list and `MaxRecentUsers` limit.

### Changed
- Username input field replaced from `TextBox` to `ComboBox` (DropDown style) showing recent logins.
- `Set-ConnectedUIState` updated: hides `$clearHistoryButton` when logged in, shows it when logged out.
- "Clear History" button text made descriptive (`"🗑 Clear History"`) with tooltip `"Clears the list of saved M365 login names"`.
- Initial form window size increased from `900×850` to `960×950` to prevent control clipping.

### Fixed
- "Clear History" button misaligned position – repositioned directly next to username ComboBox.

## [0.9.0] – 2026-04-09

### Changed
- All comments translated to English.
- PowerShell-approved verb names used throughout (`Test-`, `Invoke-`, `Get-`, `New-`, `Update-`, `Set-`, `Switch-`, `Resolve-`, `Save-`).
- Central configuration block added at top of script for all script-scoped constants and mutable state variables.

## [0.8.1] – 2026-04-09

### Added
- Winget version disk cache (JSON file, TTL 6 hours) stored at `%LOCALAPPDATA%\WinTuner_VersionCache.json`.
- In-memory RAM cache for winget version lookups (speeds up repeated queries within session).
- "Clear Cache" button in Settings tab to manually invalidate the version cache.

### Fixed
- UTC timestamp parsing in `Get-VersionDiskCache` using `RoundtripKind` flag.

## [0.8.0] – 2026-04-09

### Added
- SHA256 integrity check for self-update downloads (optional, uses `.sha256` asset from GitHub release).
- Tooltips on key controls for improved usability.
- Username persistence (last used UPN saved and restored on next start).

### Changed
- Log box height increased to 120px for better visibility.

### Fixed
- `hashMismatch` flag variable used instead of string matching for correct SHA256 mismatch detection.

## [0.7.1] – 2026-04-09

### Fixed
- `$script:wingetVersionCache` and `$script:builtVersions` not being initialized before use (cache init crash).
- `BackgroundWorker` not being disposed after completion (BGW Dispose leak).
- Thread-unsafe logging that caused crashes when `Write-Log` was called from background threads.

---

## [0.6.x and earlier]

### Added
- Global error handler for Windows Forms and AppDomain to log unhandled exceptions.
- Log rotation (2MB limit) for the log file.

### Changed
- Improved update list search filtering to use new UI input fields (`discoveredAppSearchBox`, `discoveredPublisherBox`).
- Enhanced regex to remove all text in parentheses and version numbers from Intune discovered app names, improving Winget matching (e.g., for Mozilla Firefox).
- Consolidated discovered apps list so multiple versions of the same app are grouped by Winget PackageID with a summed device count and clean name.

### Fixed
- Fixed UI freezing and "Not Responding" state during long operations (like Discovered Apps scanning) by implementing `[System.Windows.Forms.Application]::DoEvents()`.
- Corrected event listener variable names for Discovered Apps filtering (`$discoveredAppSearchBox` and `$discoveredPublisherBox`) so the UI updates immediately on input.
- Fixed PowerShell stream preferences (`WarningPreference`, etc.) at the top of the script to prevent thread crashes from `WriteObject`/`WriteError` calls.

---

### Added / Changed / Fixed
* [8c27006](https://github.com/manuelhoefler17-gif/WinTuner-GUI/commit/8c2700650dc3303eefc15f707dfe4805065c35a4) - Latest updates by @manuelhoefler17-gif
* [d42ef12](https://github.com/manuelhoefler17-gif/WinTuner-GUI/commit/d42ef12df3ccb00177d9bb60cee3f25d0038e14d) - Update by @manuelhoefler17-gif
* [ba99a49](https://github.com/manuelhoefler17-gif/WinTuner-GUI/commit/ba99a492e499ba9b8c4060bb8d7b27e0ac97db10) - Update by @manuelhoefler17-gif
* [65a92c1](https://github.com/manuelhoefler17-gif/WinTuner-GUI/commit/65a92c1ac55a0500a10e2cc636caab88ea26e71f) - Update by @manuelhoefler17-gif
* [661bd29](https://github.com/manuelhoefler17-gif/WinTuner-GUI/commit/661bd2972cd3c5abb262fbc494d0d84dc34d0a89) - Update by @manuelhoefler17-gif
* [ae08851](https://github.com/manuelhoefler17-gif/WinTuner-GUI/commit/ae08851a1fa046da912b9f339fc5cfbb3c7f33a8) - Update by @manuelhoefler17-gif
* [ae6a07c](https://github.com/manuelhoefler17-gif/WinTuner-GUI/commit/ae6a07c0189e97f72fc2901171ec9f6103e79703) - Update by @manuelhoefler17-gif
* [46a26af](https://github.com/manuelhoefler17-gif/WinTuner-GUI/commit/46a26af1a62a942a6abf28784018312f5993cb02) - Update by @manuelhoefler17-gif
* [6de3919](https://github.com/manuelhoefler17-gif/WinTuner-GUI/commit/6de39190f4fbc0db47dfda853933f5f7cb91b082) - Update by @manuelhoefler17-gif
* [a43c35d](https://github.com/manuelhoefler17-gif/WinTuner-GUI/commit/a43c35d4d705031062dde78f36a39ddb6e92ce93) - Update by @manuelhoefler17-gif
* [eec3443](https://github.com/manuelhoefler17-gif/WinTuner-GUI/commit/eec34435a95cf35918e94cda6f90148ad4495f0e) - Update by @manuelhoefler17-gif
* [e999862](https://github.com/manuelhoefler17-gif/WinTuner-GUI/commit/e999862f06acbe32ec8c9377a3f149d56072413e) - Update by @manuelhoefler17-gif
* [64e5544](https://github.com/manuelhoefler17-gif/WinTuner-GUI/commit/64e5544c8eefd2cca6f7396ebe37013795da0b78) - Update by @manuelhoefler17-gif
* [fdc8335](https://github.com/manuelhoefler17-gif/WinTuner-GUI/commit/fdc8335d99f3e659a6296cd5257328c3e45a5732) - Update by @manuelhoefler17-gif
* [2afa008](https://github.com/manuelhoefler17-gif/WinTuner-GUI/commit/2afa0083aae26edfce92c5691a33db743ba97b7e) - Update by @Copilot
* [ff01005](https://github.com/manuelhoefler17-gif/WinTuner-GUI/commit/ff01005aab623441e63dc10d3378c56a6725af43) - Update by @manuelhoefler17-gif
* [97481fd](https://github.com/manuelhoefler17-gif/WinTuner-GUI/commit/97481fd595905f1414b29cf22e190658596774f1) - Update by @manuelhoefler17-gif
* [e6a01a1](https://github.com/manuelhoefler17-gif/WinTuner-GUI/commit/e6a01a15fb266f7649820cc8769cda16903ab2b1) - Update by @manuelhoefler17-gif
* [0c16057](https://github.com/manuelhoefler17-gif/WinTuner-GUI/commit/0c16057eb64be329aa28abcbb4e89a4ea31a62c0) - Update by @manuelhoefler17-gif
* [a957abe](https://github.com/manuelhoefler17-gif/WinTuner-GUI/commit/a957abecbd28d1e3fdca7aeca4e136361740bfe4) - Update by @manuelhoefler17-gif
* [f075059](https://github.com/manuelhoefler17-gif/WinTuner-GUI/commit/f075059222b7deda9acd3f3d6833eb74c9bd8e42) - Update by @manuelhoefler17-gif
* [27e803b](https://github.com/manuelhoefler17-gif/WinTuner-GUI/commit/27e803b3fd13a1cd3c549a66d729e1187d48421d) - Update by @manuelhoefler17-gif
* [c56bc13](https://github.com/manuelhoefler17-gif/WinTuner-GUI/commit/c56bc136aafe3cae2f917fa0e5e1ceb39877ba54) - Update by @manuelhoefler17-gif
* [48f1dac](https://github.com/manuelhoefler17-gif/WinTuner-GUI/commit/48f1dac5c126fea453c2dced184d41683524d736) - Update by @manuelhoefler17-gif
* [d7b81c6](https://github.com/manuelhoefler17-gif/WinTuner-GUI/commit/d7b81c68a12346eb57babe1f932f9b0800eefbd9) - Update by @manuelhoefler17-gif
* [a75f987](https://github.com/manuelhoefler17-gif/WinTuner-GUI/commit/a75f9873a8c0fc790a24a16e2874f51135564084) - Update by @manuelhoefler17-gif
* [dd16a70](https://github.com/manuelhoefler17-gif/WinTuner-GUI/commit/dd16a7090c4152e30a7861448fa6562137629388) - Update by @manuelhoefler17-gif
* [4f3abfe](https://github.com/manuelhoefler17-gif/WinTuner-GUI/commit/4f3abfec4297dfd2e274a72346f98ccac497de7f) - Update by @manuelhoefler17-gif
* [d8f1c53](https://github.com/manuelhoefler17-gif/WinTuner-GUI/commit/d8f1c53a0c80d70a28cdb92be1571431e61c8583) - Update by @manuelhoefler17-gif
* [bacc708](https://github.com/manuelhoefler17-gif/WinTuner-GUI/commit/bacc708ec74013d3402ddaeb438f8e448b43ef81) - Initial / earlier updates by @manuelhoefler17-gif

---
*Note: You can view the full commit history and find more details on GitHub at [manuelhoefler17-gif/WinTuner-GUI/commits](https://github.com/manuelhoefler17-gif/WinTuner-GUI/commits/main).
