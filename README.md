# WinTuner GUI

> 🚀 A PowerShell-based graphical interface for packaging, deploying, discovering and updating Microsoft Intune Win32 applications with WinGet and WinTuner.

[![PowerShell Version](https://img.shields.io/badge/PowerShell-7.0%2B-blue.svg)](https://github.com/PowerShell/PowerShell)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Version](https://img.shields.io/badge/version-0.10.16-orange.svg)](CHANGELOG.md)

## 🎯 Overview

**WinTuner GUI** provides a graphical workflow around the WinTuner PowerShell module by Stephan van Rooij.

It simplifies common Microsoft Intune Win32 application tasks:

- Search WinGet packages
- Select current or older package versions
- Create WinTuner packages
- Upload packages to Microsoft Intune
- Scan existing applications for updates
- Discover applications reported by Intune and match them to WinGet
- Bulk-package and update multiple applications
- Cache expensive WinGet and Microsoft Graph lookups
- Keep long-running tenant and WinGet operations responsive, with safe cancellation where supported

WinTuner GUI requires **PowerShell 7**.

---

## ✨ Key Features

### 📦 WinGet packaging and deployment

- Search WinGet packages directly from the GUI
- Select a specific package version when required
- Package applications through WinTuner
- Reuse matching packages across GUI sessions after validating their metadata and exact `.intunewin` file
- Validate package metadata before upload
- Validate the exact `.intunewin` referenced by `win32LobApp.json`
- Prevent stale application or version selections from enabling Upload
- Upload packages directly to the connected Intune tenant

Existing package folders from earlier GUI sessions are reused only when the selected package ID and version match a valid package directory, the `displayVersion` in readable `win32LobApp.json`, and its exact non-empty `.intunewin` file.

### 🔄 Update management

- Scan WinTuner-managed Intune Win32 applications in a background runspace
- Keep the GUI responsive, show live progress and cancel an active scan safely
- Compare deployed versions with available WinGet versions
- Display the number of checked applications and update candidates
- Select individual or multiple applications
- Keep checked candidates selected while filtering the list
- Enable update actions only when their required scan results and selections exist
- Clear tenant-specific candidates on logout
- Package and deploy updates in bulk

### 🔍 Discovered Apps

The **Discovered Apps** workflow reads application inventory from Microsoft Intune and attempts to match discovered applications with WinGet packages.

The discovery pipeline includes:

- Microsoft Graph pagination and retry handling
- Cached Intune detected-app results
- Persistent WinGet discovery cache
- Application-name normalization and fuzzy matching
- Duplicate search-term handling
- Entire Graph retrieval, normalization and matching orchestration in a background runspace
- Isolated PowerShell worker processes for scalable WinGet queries
- Batched WinGet queries
- Cancellation support with partial-result cleanup
- Action-state validation based on connection, scan, deployment, results, and checked selections
- Filter- and sort-stable selection mapping
- Tenant-specific result cleanup on logout
- Cache hit statistics
- Fresh/cached Graph status
- Timestamp of the last successful discovery

Repeated scans can therefore be significantly faster when cached data is still valid.

### 💾 Caching

WinTuner GUI uses multiple caches for different workloads:

- In-memory WinGet version cache
- Persistent WinGet version cache
- Persistent WinGet discovery cache
- Cached Microsoft Intune detected-app results

The Settings tab provides **Clear All Caches** to clear the local caches used by WinTuner GUI.

Cache lifetimes are workload-specific rather than controlled by one global TTL.

### 🔐 Safer package state handling

Upload availability is calculated from the current application state.

It considers:

- Tenant connection
- Current package selection
- Selected package version
- Successfully built version
- Package metadata
- Expected `.intunewin` file

Existing packages can be reused across GUI sessions only after the selected package ID and version resolve to a valid package directory. WinTuner GUI requires matching `displayVersion` in readable `win32LobApp.json` metadata and the exact non-empty `.intunewin` file referenced by its safe filename.

Changing the application, package version, or package root recalculates Upload readiness. Invalid or stale artifacts cannot enable Upload and are rebuilt when package creation is requested.

### 🔄 Self-update

WinTuner GUI can check GitHub for newer releases.

The self-update process includes:

- Automatic and manual update checks
- Semantic version comparison
- Optional SHA256 validation
- Temporary download before replacement
- Backup of the current script
- Automatic restart after successful update
- Rollback when replacement or restart fails
- Bootstrap of required modular runtime files

### ⚙️ Settings and usability

- Persistent application settings
- Remembered username support
- Recent-user history
- Background tenant-connection verification with bounded retries
- Configurable package output folder
- Dark and light themes
- Search and filtering
- Status information for long-running operations
- Log rotation
- Crash and exception logging

---

## 📋 Requirements

- Windows 10/11 or Windows Server
- PowerShell 7.0 or newer
- Internet access
- WinGet
- Microsoft Graph access
- Microsoft Intune permissions required to manage Win32 applications

WinTuner GUI uses the **WinTuner** PowerShell module.

If the module is not available, the application attempts to install it automatically for the current user.

Manual installation:

```powershell
Install-Module -Name WinTuner -Scope CurrentUser
```

---

## 🚀 Installation

### GitHub release

For normal usage, use the latest published GitHub release.

Download `WinTuner_GUI.ps1` and run it with PowerShell 7:

```powershell
pwsh.exe -File .\WinTuner_GUI.ps1
```

Standalone releases bootstrap the additional runtime files required by the application when they are missing or belong to another release version.

### Git checkout

For development or testing:

```powershell
git clone https://github.com/manuelhoefler17-gif/WinTuner-GUI.git
cd WinTuner-GUI
pwsh.exe -File .\WinTuner_GUI.ps1
```

A Git checkout is treated as a development environment.

The GUI uses the local files from `Modules` and `Workers` and does not replace them with files downloaded from the currently published release.

If required development files are missing, startup stops with a development error instead of silently mixing local and release files.

---

## 🗂️ Project Structure

```text
WinTuner-GUI/
├── WinTuner_GUI.ps1
├── Modules/
│   ├── WinTuner.AppUpdate.psm1
│   ├── WinTuner.Connection.psm1
│   ├── WinTuner.Core.psm1
│   ├── WinTuner.DiscoveryDeployment.psm1
│   ├── WinTuner.DiscoveryScan.psm1
│   ├── WinTuner.Intune.psm1
│   ├── WinTuner.Logging.psm1
│   ├── WinTuner.PackageBuild.psm1
│   ├── WinTuner.PackageUpload.psm1
│   ├── WinTuner.Settings.psm1
│   ├── WinTuner.SupersededRemoval.psm1
│   ├── WinTuner.UpdateScan.psm1
│   └── WinTuner.Winget.psm1
├── Workers/
│   └── WinTuner.DiscoveryWorker.ps1
├── Tests/
├── README.md
├── CHANGELOG.md
└── LICENSE
```

### Main components

**`WinTuner_GUI.ps1`**
WinForms user interface, application orchestration, authentication workflow, packaging controls and update handling.

**`WinTuner.AppUpdate.psm1`**
Testable background package-build, exact artifact validation and tenant-deployment batching for existing apps.

**`WinTuner.Connection.psm1`**
Bounded, testable tenant-connection verification used after interactive authentication.

**`WinTuner.Core.psm1`**
Shared core functionality.

**`WinTuner.DiscoveryDeployment.psm1`**
Background package build, exact artifact validation and tenant deployment for checked Discovery results.

**`WinTuner.DiscoveryScan.psm1`**
Testable Discovery orchestration for Graph inventory, normalization, cached WinGet matching, deduplication and cancellation.

**`WinTuner.Intune.psm1`**
Microsoft Intune and detected-app integration.

**`WinTuner.Logging.psm1`**
Logging and log-management functionality.

**`WinTuner.PackageBuild.psm1`**
Testable package creation and WinGet version fallback decisions used by the background build workflow.

**`WinTuner.PackageUpload.psm1`**
Exact artifact revalidation and tenant deployment used by the background upload workflow.

**`WinTuner.Settings.psm1`**
Persistent settings handling.

**`WinTuner.SupersededRemoval.psm1`**
Background deletion batching with per-app results and safe handling of already absent apps.

**`WinTuner.UpdateScan.psm1`**
Testable update candidate scanning, progress and cooperative cancellation logic.

**`WinTuner.Winget.psm1`**
WinGet lookup, version and discovery functionality.

**`WinTuner.DiscoveryWorker.ps1`**
Isolated worker used for scalable WinGet discovery queries.

---

## 💻 Basic Workflow

### 1. Login

1. Enter your Microsoft 365 UPN.
2. Optionally enable remembering the username.
3. Click **Login**.
4. Complete interactive authentication.

WinTuner GUI verifies the tenant connection with bounded retries in a background runspace before enabling tenant-dependent actions.

### 2. Create and deploy a WinGet application

1. Open **WinGet Apps**.
2. Search for an application. The WinGet query runs in the background so the window remains responsive.
3. Select the package after the current search finishes.
4. Optionally choose a specific version. The version list loads in the background before the selection dialog opens.
5. Select the package output folder.
6. Click **Create package**. Package creation runs in the background so the window remains responsive.
7. After a valid build, click **Upload**. The exact artifact is revalidated in the background immediately before tenant deployment.

Changing the selected version invalidates the previous Upload state. The new version must first be built or safely reused.

### 3. Scan for updates

1. Open **Updates**.
2. Click **Search Updates**. The scan runs in the background and reports the current application and progress.
3. To stop a running scan, click **Cancel Scan**. The current WinGet query finishes and partial results are discarded.
4. Review each candidate's installed and available version.
5. Filter the results if required; checked candidates remain selected when hidden by the filter.
6. Select the applications to update.
7. Review the candidate count and checked count.
8. Confirm the checked or all-candidates update operation.

Update actions remain disabled until their required candidates or checked selections exist. Logout is disabled during scans and deployments. Both update actions use the current verified scan result, then build, validate, and deploy each app in an isolated background runspace. Successful apps are removed from the list while failed apps remain available for retry.

### 4. Find superseded Intune applications

1. Open **Updates**.
2. Click **Search Superseded Apps**. The tenant query runs in the background so the window remains responsive.
3. Select one result for individual deletion, or review the full result list before using **Delete all Superseded Apps**.

Deletion actions remain disabled until the current tenant search returns valid results. Individual and bulk deletion run in a background runspace against that displayed result set. Successful or already absent apps are removed from the list, while failures remain available for retry.

### 5. Discover Intune applications

1. Open **Discovered Apps**.
2. Start Discovery. Graph retrieval, normalization and WinGet matching run in the background; use **Cancel Scan** to stop safely.
3. Intune detected applications are collected.
4. Search terms are generated and matched against WinGet.
5. Review the matched packages and their displayed match confidence.
6. Select applications for packaging and background deployment.

Deployment remains disabled until at least one result is checked. Filtering and sorting preserve the checked objects. Each selected app is built, exactly validated and deployed outside the UI thread. Successful apps leave the candidate list; failed apps remain available for retry. Canceled or failed scans discard partial results, and logging out clears the current tenant's Discovery results.

The status indicates whether Graph data was fresh or cached and how many WinGet discovery queries came from cache.

---

## ⚙️ Configuration

Settings are stored in:

```text
%APPDATA%\WinTunerGUI\settings.json
```

Typical settings include the package path, automatic update checking, remembered users and WinGet overrides.

Cache data is stored separately under the current user's local application data directory:

```text
%LOCALAPPDATA%\WinTuner_VersionCache.json
%LOCALAPPDATA%\WinTuner_DiscoveryCache.json
%LOCALAPPDATA%\WinTuner_DetectedAppsCache.json
```

Use **Clear All Caches** in Settings when a completely fresh lookup is required.

---

## 📝 Logging

The primary GUI log is written to:

```text
%LOCALAPPDATA%\WinTuner_GUI.log
```

The application records information about:

- Authentication
- Package creation
- Package reuse
- Deployments
- Update scans
- Discovery summaries
- Cache usage
- Warnings and errors
- Self-update activity

Log rotation prevents the active log from growing indefinitely.

---

## ✅ Recommended Pre-Flight Checks

Confirm PowerShell 7:

```powershell
$PSVersionTable.PSVersion
```

Check the WinTuner module:

```powershell
Get-Module -ListAvailable -Name WinTuner
```

Check Microsoft Graph connectivity:

```powershell
Test-NetConnection graph.microsoft.com -Port 443
```

Check WinGet:

```powershell
winget --version
```

---

## 📸 Screenshots

### WinGet Apps

<img width="886" height="843" alt="WinGet Apps" src="https://github.com/user-attachments/assets/990f0de4-a5d3-4462-851d-686618faa02f" />

### Updates

<img width="886" height="843" alt="Updates" src="https://github.com/user-attachments/assets/ef03ac45-d9ac-49eb-84b6-e1abc4265c96" />

### Discovered Apps

<img width="886" height="843" alt="Discovered Apps" src="https://github.com/user-attachments/assets/c8bdb7ec-476b-465d-83ab-4fa369120a91" />

### Settings

<img width="886" height="843" alt="Settings" src="https://github.com/user-attachments/assets/ad101628-3a72-4b5e-b550-0e531e0e983a" />

---

## 🔧 Troubleshooting

### PowerShell version error

Run WinTuner GUI with `pwsh.exe`, not Windows PowerShell 5.1:

```powershell
pwsh.exe -File .\WinTuner_GUI.ps1
```

### WinTuner module not found

The application attempts installation automatically.

Manual installation:

```powershell
Install-Module -Name WinTuner -Scope CurrentUser
```

### Login fails

Verify Microsoft Intune permissions, Microsoft Graph connectivity and interactive authentication.

Then review:

```text
%LOCALAPPDATA%\WinTuner_GUI.log
```

### Upload remains disabled

Check that:

- You are logged in
- A package is selected
- The selected version has been built
- `win32LobApp.json` exists
- The `.intunewin` referenced by the metadata exists

If the selected application or version changed, create or reuse the matching package again.

### Discovery is much faster on the second scan

This is expected when Graph or WinGet results can be served from cache.

Use **Clear All Caches** if a completely fresh scan is required.

### Development checkout reports missing files

A Git checkout intentionally does not download release dependencies over development files.

Restore the missing repository files before starting the GUI again.

---

## 🧪 Development

Active development is performed on the **`Test`** branch.

The **`main`** branch is kept stable.

```text
Test
  ↓
Pull Request
  ↓
main
```

Changes should be tested on `Test` before being merged into `main`.

A development checkout intentionally uses its local dependency files and is therefore not a substitute for testing the standalone release bootstrap.

---

## 📝 Changelog

See [CHANGELOG.md](CHANGELOG.md) for the complete version history.

Current stable release: **v0.10.15**.

---

## 🗺️ Planned Improvements

Potential future improvements include:

- Further Discovery and Updates UX improvements
- Additional end-to-end and negative-path automation

---

## 🙏 Credits

WinTuner GUI is built on the **WinTuner PowerShell module** created by Stephan van Rooij.

GUI development: **Manuel Höfler**

Special thanks to Julian Hilgenberg for the initial idea, as well as the WinGet, Microsoft Graph and PowerShell communities.

---

## 📄 License

This project is licensed under the [MIT License](LICENSE).

---

## 🤝 Contributing

Contributions are welcome.

Use a development branch and submit a pull request instead of committing experimental changes directly to `main`.

---

## 📞 Support

Use the repository **Issues** area for bugs and feature requests.

When reporting an issue, include relevant WinTuner GUI log entries whenever possible.

---

<div align="center">
  <strong>Made with ❤️ for the Intune community</strong>
  <br>
  <sub>If this tool helped you, consider giving it a ⭐!</sub>
</div>
