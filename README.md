# WinTuner GUI

> 🚀 A modern PowerShell-based graphical user interface for managing Microsoft Intune Win32 applications using WinGet packages.

[![PowerShell Version](https://img.shields.io/badge/PowerShell-7.0%2B-blue.svg)](https://github.com/PowerShell/PowerShell)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Version](https://img.shields.io/badge/version-0.10.13-orange.svg)](CHANGELOG.md)

## 🎯 Overview

**WinTuner GUI** is a graphical interface built on top of the WinTuner PowerShell module by Stephan van Rooij. It simplifies packaging, deploying and updating WinGet applications in Microsoft Intune.

## ✨ Key Features

- 🔍 Search WinGet packages and deploy them to Intune
- 🔄 Scan WinTuner-managed Intune Win32 apps for available updates
- 📊 Discover Intune apps and match them to WinGet packages
- 📦 Bulk package and update applications
- 🧠 Fuzzy matching and app-name normalization
- 💾 RAM and disk cache for WinGet version lookups
- ⚙️ Persistent settings and recent-user history
- 🌗 Dark/light GUI themes
- 📝 Logging, log rotation and crash reporting
- 🤖 Automatic update checks and self-update support with optional SHA256 validation

## 📋 Requirements

- Windows 10/11 or Windows Server 2016+
- PowerShell 7.0+
- Internet access
- Microsoft Intune permissions required to manage Win32 applications
- WinTuner PowerShell module

```powershell
Install-Module -Name WinTuner -Scope CurrentUser
```

## 🚀 Installation

### Direct download

Download `WinTuner_GUI.ps1` and run it with PowerShell 7:

```powershell
.\WinTuner_GUI.ps1
```

### Git

```powershell
git clone https://github.com/manuelhoefler17-gif/WinTuner-GUI.git
cd WinTuner-GUI
.\WinTuner_GUI.ps1
```

## ✅ Recommended Pre-Flight Checks

```powershell
$PSVersionTable.PSVersion
Get-Module -ListAvailable WinTuner
Test-NetConnection graph.microsoft.com -Port 443
```

## 💻 Basic Workflow

### Login

1. Enter your Microsoft 365 UPN.
2. Optionally enable remembering the username.
3. Click **Login** and complete interactive authentication.

### Deploy a WinGet app

1. Open **WinGet Apps**.
2. Search for an application.
3. Select the package and optionally a specific version.
4. Create the package.
5. Deploy it to Intune.

### Update apps

1. Open **Updates**.
2. Click **Search Updates**.
3. Filter or sort the results if needed.
4. Select the apps to update.
5. Start the update operation.

### Discover apps

The **Discovered Apps** tab scans Intune discovered applications and attempts to match them with WinGet packages. Results can be filtered by application or publisher and selected apps can be packaged and deployed.

## ⚙️ Configuration

Settings are stored in:

```text
%LOCALAPPDATA%\WinTuner_Settings.json
```

Example:

```json
{
  "DefaultPackagePath": "C:\\Packages",
  "AutoCheckUpdates": true,
  "RememberMe": true,
  "LastUser": "admin@contoso.com",
  "RecentUsers": ["admin@contoso.com"],
  "MaxRecentUsers": 3,
  "WingetOverrides": {
    "7zip.7zip": "24.07"
  }
}
```

The WinGet version cache is stored separately in `%LOCALAPPDATA%\WinTuner_VersionCache.json` and uses a six-hour TTL.

## 📝 Logging

Logs are written to:

```text
%LOCALAPPDATA%\WinTuner_GUI.log
```

The application logs authentication events, package creation, deployments, update checks, warnings and errors. Log rotation prevents the active log from growing indefinitely.

## 📸 Screenshots

### WinGet Apps
<img width="886" height="843" alt="WinGet Apps" src="https://github.com/user-attachments/assets/990f0de4-a5d3-4462-851d-686618faa02f" />

### Updates
<img width="886" height="843" alt="Updates" src="https://github.com/user-attachments/assets/ef03ac45-d9ac-49eb-84b6-e1abc4265c96" />

### Discovered Apps
<img width="886" height="843" alt="Discovered Apps" src="https://github.com/user-attachments/assets/c8bdb7ec-476b-465d-83ab-4fa369120a91" />

### Settings
<img width="886" height="843" alt="Settings" src="https://github.com/user-attachments/assets/ad101628-3a72-4b5e-b550-0e531e0e983a" />

## 🔧 Troubleshooting

**PowerShell version error**  
Run the application with `pwsh.exe` and PowerShell 7 or newer.

**WinTuner module not found**  
Install the module with `Install-Module -Name WinTuner -Scope CurrentUser`.

**Login fails or hangs**  
Verify Intune permissions and connectivity to Microsoft Graph, then restart the application if necessary.

**Update check does not run automatically**  
Verify that automatic update checking is enabled in Settings and review `WinTuner_GUI.log` for errors.

## 📝 Changelog

See [CHANGELOG.md](CHANGELOG.md) for the detailed version history.

Current application version: **0.10.13**.

## 🙏 Credits

WinTuner GUI is built on the WinTuner PowerShell module created by Stephan van Rooij.

GUI development: **Manuel Höfler**.

Special thanks to Julian Hilgenberg for the initial idea, as well as the WinGet, Microsoft Graph and PowerShell communities.

## 📄 License

This project is licensed under the [MIT License](LICENSE).

## 🤝 Contributing

Contributions are welcome. For development work, create a dedicated branch and submit a pull request rather than committing experimental changes directly to `main`.

## 📞 Support

Use the repository Issues area for bugs and feature requests. See the changelog for known fixes and recent changes.

---

<div align="center">
  <strong>Made with ❤️ for the Intune community</strong>
  <br>
  <sub>If this tool helped you, consider giving it a ⭐!</sub>
</div>
