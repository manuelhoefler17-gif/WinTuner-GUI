# WinTuner GUI – Project Status

## Current Version

Development version: **0.10.18**

Latest published stable release: **v0.10.17**

Repository:
`manuelhoefler17-gif/WinTuner-GUI`

Branches:

- `main` = stable
- `Test` = active development
- Changes are developed and tested on `Test`
- Changes reach `main` only through Pull Requests

At the current project state, `main` and `Test` are synchronized.

---

## Development Environment

Primary local repository:

`C:\Intune\WinTuner\Entwicklung`

Do not modify:

`C:\Intune\WinTuner\WinTuner_GUI.ps1`

PowerShell version:

PowerShell 7 (`pwsh.exe`)

Run the development GUI with:

`pwsh.exe -File .\WinTuner_GUI.ps1`

A Git checkout is considered a development environment.

Development checkouts use local `Modules` and `Workers` files and must not silently replace them with release files.

---

## Current Architecture

Main application:

- `WinTuner_GUI.ps1`

Modules:

- `Modules/WinTuner.AppUpdate.psm1`
- `Modules/WinTuner.Connection.psm1`
- `Modules/WinTuner.Core.psm1`
- `Modules/WinTuner.DiscoveryDeployment.psm1`
- `Modules/WinTuner.DiscoveryScan.psm1`
- `Modules/WinTuner.Intune.psm1`
- `Modules/WinTuner.Logging.psm1`
- `Modules/WinTuner.PackageBuild.psm1`
- `Modules/WinTuner.PackageUpload.psm1`
- `Modules/WinTuner.Settings.psm1`
- `Modules/WinTuner.SupersededRemoval.psm1`
- `Modules/WinTuner.UpdateScan.psm1`
- `Modules/WinTuner.Winget.psm1`

Workers:

- `Workers/WinTuner.DiscoveryWorker.ps1`

Tests:

- `Tests/`

---

## Completed Work for 0.10.14

### Self Update

The self-update system has been stabilized.

Implemented:

- automatic restart after successful update
- removal of obsolete manual restart prompt
- backup before script replacement
- rollback if update replacement/restart fails
- SHA256 validation support
- temporary download before replacement
- semantic version comparison
- standalone release bootstrap
- development checkout detection

Important behavior:

A Git checkout must not download release dependencies over local development files.

If required local development files are missing, startup stops with a development error.

---

### Discovery

The Discovered Apps workflow has been significantly improved.

Implemented:

- Microsoft Graph pagination
- retry handling
- detected-app cache
- persistent WinGet discovery cache
- isolated PowerShell worker processes
- batched WinGet lookups
- cancellation support
- duplicate search-term handling
- fuzzy matching
- application-name normalization
- cache-hit reporting
- Graph Fresh/Cached status
- last successful discovery timestamp

Repeated discovery scans should make heavy use of cached data when appropriate.

---

### Update Scan UX

Update scanning now reports:

- total applications checked
- update candidates found

Example:

`Update scan complete | Checked: 18 | Candidates: 13`

---

### Package Creation and Upload

Package handling has been hardened.

Implemented:

- session-based same-version package reuse
- selecting another package version invalidates Upload
- successful package builds store the effective built version
- obsolete `.wtpackage` handling removed
- upload validates package root existence
- upload validates `win32LobApp.json`
- upload reads the exact expected `.intunewin` filename from metadata
- upload is blocked if the exact referenced `.intunewin` is missing

Deployment still uses:

`Deploy-WtWin32App -PackageId ... -Version ... -RootPackageFolder ...`

---

### Centralized Upload State

Upload button state is now calculated centrally.

The state considers:

- tenant connection
- selected application
- selected package version
- built version
- package metadata
- expected `.intunewin`

State is recalculated on:

- app selection changes
- version changes
- login
- logout
- successful build
- package reuse
- upload errors

A failed upload must not blindly re-enable the Upload button.

---

### README

README has been rewritten for the current architecture.

It now documents:

- modular architecture
- Modules and Workers
- release bootstrap
- development checkout behavior
- package validation
- session-based package reuse
- discovery pipeline
- cache behavior
- self-update
- Test -> Pull Request -> main workflow
- troubleshooting
- future plans

---

## Important Design Decisions

### Package Reuse

Current behavior:

Package reuse supports GUI restarts, but existing package folders are never trusted solely because they exist.

For the currently selected package ID and version, reuse requires:

- a package directory under the selected package root
- readable `win32LobApp.json` metadata
- a metadata `displayVersion` matching the selected version
- a safe leaf filename ending in `.intunewin`
- the exact non-empty `.intunewin` referenced by the metadata
- no path traversal or reparse-point package file

`$script:builtVersions` remains an in-memory state cache. After a restart it is repopulated only from an artifact that passes the same centralized validation used for Upload state, post-build checks, and click-time upload checks.

Invalid or stale artifacts keep Upload disabled and package creation rebuilds them.

---

### Upload Safety

Never enable Upload only because a directory or `.intunewin` exists.

The selected package/version and the build metadata must match.

Click-time validation remains a final defense even if UI state already indicates the package is ready.

---

### Git Workflow

Never develop directly on `main`.

Normal process:

`Test -> Pull Request -> main`

After a PR merge:

1. pull `main`
2. fast-forward local `Test`
3. push `Test`
4. verify `main` and `Test` are identical

Do not use:

`git add .`

or:

`git add -A`

Stage only explicitly intended files.

---

## 0.10.14 Release Status

Released: **2026-09-09**

Completed:

- finalized `CHANGELOG.md` for 0.10.14
- completed manual E2E testing
- verified version references and release documentation
- validated PowerShell syntax and automated tests
- created tag `v0.10.14`
- published the GitHub release
- attached `WinTuner_GUI.ps1` and its SHA256 checksum
- verified standalone dependency bootstrap
- verified self-update from v0.10.13 to v0.10.14, including automatic restart, backup, dependency refresh, and checksum validation

---

## 0.10.15 Release Status

Released: **2026-09-10**

Completed:

- safe package reuse across GUI restarts
- centralized persistent artifact validation
- package-root state recalculation
- automated positive and negative package-artifact tests
- centralized Updates action state based on connection, scan results, selection, and active operations
- filter-stable checked update candidates
- tenant-specific update candidate cleanup on logout
- automated update-action-state tests
- centralized Discovery action state for connection, scan, cancellation, deployment, results, and selection
- filter- and sort-stable Discovery selections with unique object mapping
- partial-result cleanup after canceled or failed Discovery scans
- tenant-specific Discovery cleanup on logout
- successful Discovery deployment removal from the candidate list
- automated Discovery action-state tests
- exact post-build artifact validation before Update and Discovery deployments
- centralized protected package-root validation across package creation, Upload, Updates, and Discovery
- automated package-root validation tests
- installed-to-available version display, candidate/checked counts, and confirmation for checked updates
- Update All based on the current verified scan result
- Discovery match-confidence display and CSV export
- shared local syntax/Pester test runner
- GitHub Actions checks for pull requests to main and pushes to Test
- corrected settings and cache-path documentation
- responsive main content, Updates, Discovery, and header controls without unused space or button overlap
- finalized CHANGELOG.md for 0.10.15
- validated 14 PowerShell files and 39 automated tests
- created and published tag v0.10.15
- published the GitHub release with WinTuner_GUI.ps1 and its SHA256 checksum
- verified the published assets and checksum through an independent download
- verified standalone dependency bootstrap directly from tag v0.10.15
- verified self-update from v0.10.14 to v0.10.15, including checksum validation, backup, replacement, restart, and dependency bootstrap

---

## 0.10.16 Release Status

Released: **2026-09-11**

Completed:

- moved the Intune and WinGet update scan into an isolated PowerShell runspace
- kept the WinForms UI responsive while the scan loads and checks applications
- changed **Search Updates** into **Cancel Scan** while a scan is active
- added cooperative cancellation after the current WinGet query
- discarded partial candidates on cancellation or scan failure
- kept update and logout actions centrally disabled during the background scan
- added a dedicated update-scan module and automated positive, fallback, cancellation, missing-version, and error tests
- added clean update-scan runspace shutdown when the GUI closes
- moved the Superseded Apps tenant query into an isolated PowerShell runspace
- centralized Superseded search and deletion state around connection, active operations, current results, and valid selection
- changed bulk Superseded deletion to use the current verified search result instead of loading a second tenant list
- added clean Superseded-search runspace shutdown when the GUI closes
- added automated Superseded action-state and cross-workflow busy-state tests
- moved the WinGet Apps package search into an isolated PowerShell runspace
- kept WinGet result ordering and the current query stable while the background search completes
- centralized package-search controls around active operations, current results, and valid selection
- disabled tenant workflows while a WinGet package search owns the shared progress UI
- added clean package-search runspace shutdown when the GUI closes
- added automated WinGet package-search action-state tests
- upgraded GitHub Actions checkout from v4 to v7 for the native Node.js 24 runtime
- moved WinGet version-list retrieval into an isolated PowerShell runspace
- kept the cached version ordering and modal version picker behavior while preventing concurrent workflows
- added clean version-lookup runspace shutdown when the GUI closes
- moved **Create package** into an isolated PowerShell runspace
- preserved package reuse and exact post-build artifact validation on the UI workflow
- kept 404 fallback and hash-mismatch decisions explicit without showing dialogs from the worker runspace
- added clean package-build runspace shutdown when the GUI closes
- added automated package-build fallback and retry tests
- moved **Upload to Tenant** into an isolated PowerShell runspace
- retained click-time artifact validation and added a second exact validation inside the upload worker immediately before deployment
- kept package selection and Upload available after tenant deployment failures while invalid artifacts remain blocked
- added clean package-upload runspace shutdown when the GUI closes
- added automated upload validation, deployment-error, and workflow-routing tests
- moved checked and all-app update deployment into an isolated PowerShell runspace
- kept exact post-build artifact validation immediately before every update deployment
- removed only successful updates from the current candidate list while preserving failed candidates for retry
- added clean update-deployment runspace shutdown when the GUI closes
- added automated update-deployment success, missing-ID, invalid-artifact, partial-failure, and workflow-routing tests
- moved individual and bulk Superseded deletion into an isolated runspace with per-app results and failure-preserving retry behavior
- moved the full Discovery Graph, normalization, cached WinGet matching, deduplication, and cancellation pipeline into an isolated runspace
- moved checked Discovery package creation, exact artifact validation, and tenant deployment into an isolated runspace
- moved post-authentication tenant verification retries out of the WinForms thread
- removed all remaining `DoEvents()` re-entrancy points
- reduced Login, Superseded, and Discovery event handlers to focused workflow entry points
- added module-level negative-path tests and isolated-runspace integration tests for the new workflows

Validation completed:

- PowerShell syntax validation passed for all 34 repository scripts and modules
- all 93 Pester tests passed
- a separate PowerShell runspace loaded the installed WinTuner module and completed the update-scan logic
- cooperative runspace cancellation completed without blocking the caller
- the tenant-connected GUI remained responsive while moving, resizing, and switching tabs during a scan
- canceling a tenant scan worked, and a subsequent full update scan completed successfully
- the tenant-connected Superseded Apps search kept the GUI responsive and returned its current result list with correct action states
- the Superseded search action was enabled after tenant verification, and Discovery search, publisher, and sort controls remained visible through resize and maximization
- the WinGet Apps package search kept the GUI responsive and restored result, version, and package actions after completion
- the WinGet version list loaded without blocking the GUI, and the modal picker restored the expected package actions after selection or cancellation
- package creation completed without blocking window movement, resizing, or tab changes; the exact 7zip.7zip 26.03 artifact passed validation and was safely reused on the second request
- the upload safety module revalidated that exact artifact in an isolated runspace before a non-writing deployment callback
- the update-deployment module revalidated the exact 7zip.7zip 26.03 artifact in an isolated runspace before a non-writing Graph-ID deployment callback
- a real tenant app update completed through the background worker while the GUI remained responsive and the successful candidate was removed from the current update list
- the real tenant upload of 7zip.7zip 26.03 completed successfully in the background while the GUI remained responsive and cleared the package selection afterward
- finalized `CHANGELOG.md` for 0.10.16
- created and pushed annotated tag `v0.10.16` from the synchronized `main` and `Test` commit
- verified standalone dependency bootstrap directly from tag `v0.10.16`, including all 14 runtime files and the release marker
- published the GitHub release with `WinTuner_GUI.ps1` and its SHA256 checksum
- independently downloaded both release assets and verified version, syntax, and SHA256 integrity

Next candidates:

- tenant-connected end-to-end automation where credentials and a safe test tenant can be provided
- further Discovery refinements based on production feedback

---

## 0.10.17 Release Status

Released: **2026-09-11**

Completed:

- restore the Login button immediately after logout when the retained username is valid
- establish the Microsoft Graph context during the initial tenant login and reuse it across background runspaces
- handle fresh Graph dictionary responses in Discovery normalization
- retry the narrow transient Intune collection-enumeration failure while loading existing managed apps
- add regression coverage for fresh Graph data, the transient enumeration retry, logout state, initial Graph authentication, and process-wide context reuse

Automated validation:

- PowerShell syntax validation passed for all 34 repository scripts and modules
- all 98 Pester tests passed

- merged the implementation through Pull Request #125 and the validation status through Pull Request #126
- finalized the 0.10.17 changelog and release version references
- created and pushed annotated tag v0.10.17 from synchronized main and Test
- published the GitHub release with WinTuner_GUI.ps1 and its SHA256 checksum
- independently verified the published assets, checksum, syntax, and standalone dependency bootstrap

---

## 0.10.18 Release Candidate Status

Merged through Pull Request #128 and manually validated:

- added a Discovery checkbox that bypasses the detected-app cache for a deliberately fresh Graph scan
- passed the force-refresh choice explicitly into the isolated Discovery runspace
- kept the refresh choice disabled while a scan or conflicting workflow is active
- added clear status and log reporting for cached-allowed versus forced-fresh scans
- added an interactive read-only tenant E2E runner covering WinTuner authentication, bounded tenant verification, Graph authentication, and detected-app retrieval
- ensured the E2E runner contains no package deployment, creation, or tenant-removal commands and always disconnects both sessions
- updated README test instructions, cache guidance, roadmap, and the previously stale stable-release text
- documented the complete delegated Microsoft Graph permission set requested during login, including its source, purpose, admin-consent requirement, and separate Intune-role requirement
- manually confirmed the forced-fresh Discovery scan returns current tenant results

Additional work merged through Pull Request #129 and manually validated:

- added a Settings dialog that exposes all required delegated Microsoft Graph permissions before sign-in
- manually confirmed the dialog opens before sign-in and displays correctly

Automated validation:

- PowerShell syntax validation passed for all 37 repository scripts and modules
- all 107 Pester tests passed, including focused Discovery, E2E safety, and permission-documentation tests

Pull Requests #128, #129, and #130 were merged. The implementation, permission dialog, and validation status are present on main and Test.

Release preparation completed:

- finalized the 0.10.18 changelog
- verified the application version is 0.10.18
- verified README and project status still identify v0.10.17 as the latest published release
- calculated the release-candidate WinTuner_GUI.ps1 SHA256: E1B0ADD03A4D242625D9AA3F1DB21B0578F668D8111698087BC26CB5E61BD1C1

Remaining release-only actions:

- create and push the annotated v0.10.18 tag after separate authorization
- publish the GitHub release and checksum asset after separate authorization
- verify standalone dependency bootstrap and independently validate the published assets from the tag

---
## Testing Expectations

Before committing significant changes:

- run PowerShell syntax validation
- inspect `git diff`
- run `git diff --check`
- test affected GUI flow
- test negative/error paths where relevant

Do not commit until behavior is confirmed.

---

## Coding / Patch Style

Prefer small, safe, incremental changes.

For automated patches:

- use `$ErrorActionPreference = 'Stop'`
- assert expected matches before replacement
- modify in memory
- parse PowerShell syntax before writing
- write only after validation succeeds
- avoid broad regex replacements
- avoid editing unrelated files
- avoid placeholder code such as `...`

For significant function changes, prefer replacing a complete known block rather than fragile partial edits.

---

## Current Release History Context

Latest published stable release:

**v0.10.17**

Current development version:

**0.10.18**

0.10.17 was published on 2026-09-11.
