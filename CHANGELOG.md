# Changelog

All notable changes to WinTuner GUI are documented here.

## [0.10.21] – 2026-09-12

### Added
- Added an app-only Microsoft Graph preflight that separately validates managed-app and detected-app access and identifies the exact missing Application permission on 403 responses.
- Extended the read-only tenant E2E runner to support interactive, client-secret and certificate authentication without tenant writes.

### Changed
- App-only login now clears only the dedicated WinTuner client-credential token cache before connecting so newly granted Entra application roles are requested immediately.
- App-only login status and permission errors now distinguish the access being checked and explain that a fresh WinTuner token was requested.

## [0.10.20] – 2026-09-12

### Added
- Added app-only Microsoft Entra authentication through a customer-owned App Registration using either a client secret or a certificate.
- Added an Authentication settings dialog for selecting the sign-in mode and storing tenant, client-ID and certificate-thumbprint configuration.
- Added centralized authentication configuration, certificate availability and Microsoft Graph context validation with automated positive, negative and security regression coverage.

### Changed
- Interactive user sign-in remains the default while tenant workflows and Discovery now reuse the validated Graph context for the selected authentication mode.
- Client secrets can optionally be saved as Windows DPAPI CurrentUser-protected data for automatic reuse by the same Windows user on the same computer. Without secure storage they are requested in a masked one-login dialog. Plaintext secrets are never persisted or logged and are cleared from temporary connection state after each attempt.
- The authentication dialog now shows mode-specific permission requirements, while the general in-app summary and README clearly separate interactive User Login permissions from the two Application permissions required on a customer-owned Entra App Registration.
- App-only Intune 403 responses now show the missing application permission, admin-consent and tenant-license checks instead of raw service JSON.

## [0.10.19] – 2026-09-12

### Added
- Added Graph retrieval time and cache-age metadata to Discovery scan results.
- Added a read-only two-pass tenant E2E mode that verifies a fresh Microsoft Graph request is reused from the persistent detected-app cache.

### Changed
- Discovery now shows whether the latest Graph inventory was fresh or cached, including the age of cached data in the tab header, status, and log summary.
- The Updates tab now keeps the latest batch result visible after completion and explicitly states when failed candidates remain available for retry.
- Discovery failures now identify the processing stage that failed.

## [0.10.18] – 2026-09-11

### Added
- Added a Discovery option to bypass the detected-app cache and request fresh Microsoft Graph data for the current scan.
- Added an explicit read-only tenant end-to-end runner for interactive WinTuner and Microsoft Graph validation without package deployment or tenant removal.
- Documented the complete delegated Microsoft Graph permission set requested by WinTuner GUI and WinTuner 1.3.2, including admin-consent and Intune-role requirements.
- Added a Settings dialog that shows the required delegated Graph permissions before tenant sign-in.

## [0.10.17] – 2026-09-11

### Changed
- Microsoft Graph authentication now completes during the initial tenant login and its process-wide context is reused by background Discovery scans.

### Fixed
- Restored the **Login to Tenant** button immediately after logout when the retained username is valid.
- Normalized fresh Microsoft Graph dictionary responses correctly so an uncached Discovery scan no longer produces an empty result set.
- Retried the narrow transient Intune error "Collection was modified; enumeration operation may not execute" while loading existing managed apps instead of discarding the entire Discovery scan.

## [0.10.16] – 2026-09-11

### Added
- Added a dedicated, testable update-scan module with automated coverage for verified versions, tenant-version fallback, cancellation, missing versions, and load failures.
- Added an in-app cancel action and live progress for active update scans.
- Added automated action-state coverage for Superseded search results, stale selections, disconnected tenants, and active operations.
- Added automated action-state coverage for WinGet package search results, stale selections, and active operations.
- Added a dedicated update-deployment module with automated success, missing-ID, invalid-artifact, and partial-failure coverage.
- Added dedicated modules for Discovery scanning, Discovery deployment, Superseded removal, and bounded tenant-connection verification.
- Added isolated-runspace integration coverage for Discovery matching, non-writing deployment/removal callbacks, and connection retries.

### Changed
- Update scans now run in an isolated PowerShell runspace so Intune and WinGet checks no longer block the WinForms UI.
- Canceling an update scan finishes the current lookup safely and discards partial candidates before update actions are re-enabled.
- WinGet package searches now run in an isolated PowerShell runspace and keep the WinForms UI responsive.
- WinGet version lists now load in the background before the version picker opens.
- WinGet package creation now runs in an isolated PowerShell runspace while preserving fallback and exact artifact validation.
- Validated WinGet package uploads now run in the background and revalidate the exact artifact inside the worker before tenant deployment.
- Checked and all-app update deployments now build, revalidate, and deploy packages in an isolated runspace while preserving failed candidates for retry.
- GitHub Actions validation now uses `actions/checkout@v7` with its native Node.js 24 runtime.
- Superseded-app searches now run in an isolated PowerShell runspace and keep the WinForms UI responsive.
- Superseded deletion actions now require a valid current result, and **Delete all Superseded Apps** uses the verified search results instead of fetching a second list.
- Superseded deletions now run in the background with per-app results; successful or already absent apps leave the list while failures remain available for retry.
- The complete Discovery scan now runs outside the WinForms thread, retains cache and pagination behavior, supports cooperative cancellation, and discards partial results on failure.
- Discovery packaging and tenant deployment now run in a background worker with exact artifact validation immediately before each upload.
- Post-authentication tenant verification now retries in a background runspace, and all remaining `DoEvents()` re-entrancy points were removed.
- Large Login, Superseded and Discovery event handlers now delegate to smaller workflow functions.

### Fixed
- Restored **Search Superseded Apps** after a tenant connection is verified.
- Kept the Discovery search, publisher, and sort controls visible when the window is resized or maximized.
- Included every 0.10.16 runtime module in the standalone dependency bootstrap manifest.

## [0.10.15] – 2026-09-10

### Added
- Added safe package reuse across GUI restarts after validating the selected package ID, version, metadata display version, and exact non-empty `.intunewin` file.
- Added automated package-artifact tests for valid builds, missing or malformed metadata, unsafe filenames, wrong file types, missing exact files, empty files, and unsafe package identifiers.
- Added automated update-action-state tests for disconnected, empty, selected, fully selected, and busy states.
- Added automated Discovery action-state tests for disconnected, empty, selected, scanning, canceling, and deploying states.
- Added centralized package-root validation and automated tests for normal, empty, file, drive-root, and protected-system paths.
- Added installed-to-available version details and candidate/checked counts to the Updates tab.
- Added Discovery match confidence to the result list and CSV export.
- Added a shared PowerShell test runner and GitHub Actions validation for pull requests and pushes to Test.

### Changed
- Centralized package-artifact validation across Upload state calculation, package reuse, post-build verification, and click-time upload checks.
- Package-root changes now immediately recalculate whether Upload is safe.
- Centralized Updates action state so scanning, selection, connection, and busy state determine which actions are available.
- Checked update candidates remain selected through filtering and are included when updating checked apps.
- Logging out clears tenant-specific update candidates and keeps update actions disabled until a new scan.
- Centralized Discovery action state around connection, scan and cancellation progress, deployment, results, and checked selections.
- Discovery selections now remain tied to their result objects through filtering and sorting, including duplicate display text.
- Canceled or failed Discovery scans discard partial results, and successful deployments are removed from the candidate list.
- Logging out clears tenant-specific Discovery results; logout remains unavailable during active Update or Discovery operations.
- Update and Discovery deployments now validate the exact generated package artifact before any tenant upload.
- Package creation, Upload, Updates, and Discovery now share the same protected package-root validation.
- Updating checked apps now requires confirmation, and both update actions show the installed and target versions.
- Update All now operates on the current verified scan result instead of fetching a second candidate list.
- Corrected the documented settings path and listed the persistent cache files.
- Main tabs, logs, progress, status, update results, Discovery results, and header actions now adapt to the window width.

## [0.10.14] – 2026-09-09

### Added
- Added reuse of successfully built packages with the same version during the current GUI session. Packages from previous sessions are not automatically trusted.
- Added the time of the last successful discovery scan to the Discovered Apps tab.
- Added discovery status and log summaries showing fresh/cached Graph data and WinGet cache-hit counts.

### Changed
- Update scan status and logs now report both the number of applications checked and the number of update candidates.
- Centralized Upload button state calculation around the tenant connection, selected package/version, built version, package metadata, and expected `.intunewin` file.
- Removed obsolete `.wtpackage` overwrite handling.
- Refreshed the README for the modular architecture, release bootstrap, development checkout behavior, package validation, and session-based package reuse.
- Added repository development guidance and project status documentation for the `Test` -> Pull Request -> `main` workflow.

### Fixed
- Self-update now requires a successful backup before replacing the script and attempts to restore that backup if an error occurs after replacement, including a restart failure.
- Removed the obsolete manual restart prompt after self-update to match the automatic restart behavior.
- Git development checkouts no longer download release dependencies over local `Modules` and `Workers` files; startup stops with a clear error if required local files are missing.
- Upload readiness is recalculated after package/version selection changes, login/logout, successful builds, package reuse, and upload errors.
- Upload validates the package root, readable `win32LobApp.json`, and the exact `.intunewin` filename referenced by its metadata before deployment.
- Upload is no longer unconditionally re-enabled after deployment; it stays disabled after success and is revalidated after an upload error.

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
