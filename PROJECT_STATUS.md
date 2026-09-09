# WinTuner GUI – Project Status

## Current Version

Development version: **0.10.14**

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

Package reuse is intentionally **session-based only**.

`$script:builtVersions` is reset when the GUI restarts.

Existing package directories from earlier sessions are not automatically trusted.

Cross-session package reuse is deferred until package folders can be validated reliably.

Planned target:

**0.10.15**

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

## Current Next Steps

### 0.10.14 Release Preparation

1. Complete `CHANGELOG.md` for 0.10.14
2. Perform final E2E testing
3. Verify version references
4. Verify README and release documentation
5. Verify syntax/tests
6. Tag `v0.10.14`
7. Create GitHub release
8. Attach release files
9. Generate and attach SHA256 checksum
10. Verify standalone update/bootstrap path

---

## Planned 0.10.15 Work

Potential improvements:

- safe package reuse across GUI restarts
- persistent build validation
- additional automated tests
- further Discovery UX improvements
- further Updates UX improvements

Cross-session reuse must validate existing package metadata before trusting any previous build.

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

**v0.10.13**

Current development version:

**0.10.14**

0.10.14 is not yet published.

Do not create a release until the changelog and final E2E tests are complete.