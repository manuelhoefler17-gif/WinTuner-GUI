# WinTuner GUI – Project Status

## Current Version

Development version: **0.10.15**

Latest published stable release: **v0.10.14**

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

- `Modules/WinTuner.Core.psm1`
- `Modules/WinTuner.Intune.psm1`
- `Modules/WinTuner.Logging.psm1`
- `Modules/WinTuner.Settings.psm1`
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

## Current 0.10.15 Work

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

Next candidates:

- additional end-to-end automation
- further Discovery and Updates refinements based on production feedback

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

**v0.10.14**

Current development version:

**0.10.15**

0.10.14 was published on 2026-09-09.
