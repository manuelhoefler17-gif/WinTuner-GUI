# WinTuner GUI

> 🚀 A modern PowerShell-based graphical user interface for managing Microsoft Intune Win32 applications using WinGet packages.

[![PowerShell Version](https://img.shields.io/badge/PowerShell-7.0%2B-blue.svg)](https://github.com/PowerShell/PowerShell)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Version](https://img.shields.io/badge/version-2.0-orange.svg)](CHANGELOG.md)

## 📋 Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Usage](#usage)
- [Configuration](#configuration)
- [Screenshots](#screenshots)
- [Changelog](#changelog)
- [Credits](#credits)
- [License](#license)

## 🎯 Overview

**WinTuner GUI** is a comprehensive graphical interface built on top of the [WinTuner PowerShell module](https://github.com/svrooij/WinTuner) by Sander Rozemuller. It simplifies the process of packaging and deploying WinGet applications to Microsoft Intune, managing app updates, and performing version rollbacks.

### Key Capabilities

- 🔍 **Search & Deploy**: Search WinGet packages and deploy them to Intune
- 🔄 **Update Management**: Check for and deploy updates to existing Intune apps
- ⏮️ **Version Rollback**: Roll back apps to previous versions
- ⚙️ **Persistent Settings**: Save your preferences and configurations
- 🤖 **Auto-Update Check**: Automatically check for updates on login

## ✨ Features

### 📦 WinGet App Management
- Search the WinGet repository for applications
- Select specific versions or use latest
- Create `.wtpackage` files locally
- Deploy packages directly to Microsoft Intune
- Override WinGet versions per app

### 🔄 Updates Tab
- Scan all Intune Win32 apps for available updates
- Filter apps by name
- Bulk select apps for update
- Check All / Uncheck All functionality
- Detailed progress indication
- Auto-check for updates on login (optional)

### ⏮️ Rollback Functionality
- View all versions of an app in Intune
- Select which version to keep
- Two rollback modes:
  - **Old Version Rollback**: Delete all versions and redeploy an older version
  - **Cleanup Mode**: Keep newest version, remove all others
- Smart deletion order (prevents parent-child conflicts)
- Clear warnings about assignment loss
- Only shows apps with multiple versions

### ⚙️ Settings & Configuration
- **Default Package Path**: Set your preferred package storage location
- **Auto-Check Updates**: Enable automatic update checking after login
- **Remember Me**: Save your username for quick login
- **Persistent Storage**: All settings saved to JSON file
- Settings persist across sessions

### 🔐 Authentication
- Interactive Microsoft 365 login
- Remember last username
- Secure credential handling
- Session management
- Automatic reconnection

## 📋 Requirements

### System Requirements
- **Windows 10/11** or **Windows Server 2016+**
- **PowerShell 7.0** or higher ([Download](https://github.com/PowerShell/PowerShell))
- **.NET Framework 4.7.2** or higher
- Internet connection

### PowerShell Modules
- **WinTuner** (auto-installed if missing)
  ```powershell
  Install-Module -Name WinTuner -Scope CurrentUser
  ```

### Permissions
- **Microsoft Intune Administrator** role or equivalent
- Permissions to manage Intune Win32 apps

## 🚀 Installation

### Option 1: Direct Download
1. Download `WinTuner_GUI.ps1` from this repository
2. Save to your preferred location
3. Run with PowerShell 7:
   ```powershell
   .\WinTuner_GUI.ps1
   ```

### Option 2: Git Clone
```powershell
git clone https://github.com/yourusername/WinTuner-GUI.git
cd WinTuner-GUI
.\WinTuner_GUI.ps1
```

### First Run
On first run, the script will:
1. Check for PowerShell 7 (displays error if < 7)
2. Check for WinTuner module
3. Prompt to install WinTuner if missing
4. Create settings file in `%LOCALAPPDATA%\WinTuner_Settings.json`

## 💻 Usage

### Basic Workflow

#### 1. Login
```
1. Enter your M365 UPN (e.g., admin@contoso.com)
2. Check "Remember last username" (optional)
3. Click "Login"
4. Complete interactive authentication
```

#### 2. Deploy New App
```
1. Go to "WinGet Apps" tab
2. Search for application (e.g., "7-Zip")
3. Select from dropdown
4. (Optional) Click "Select Version" to choose specific version
5. Click "Create .wtpackage"
6. Click "Deploy to Intune"
```

#### 3. Check for Updates
```
1. Go to "Updates" tab
2. Click "Search Updates"
3. Review apps with available updates
4. Check apps to update
5. Click "Update Checked Apps"
```

#### 4. Rollback App Version
```
1. Go to "Updates" tab → Rollback section
2. Click "Load Apps for Rollback"
3. Select app from dropdown (only apps with 2+ versions shown)
4. Select version to keep
5. Click "Execute Rollback"
6. Confirm the action
```

### Settings Configuration

Navigate to **Settings** tab:

- **Default Package Folder**: Set where `.wtpackage` files are stored
- **Check for updates on login**: Auto-trigger update search after login
- **Remember last username**: Save username for next session

Click **"Save Settings"** to persist changes.

## ⚙️ Configuration

### Settings File Location
```
%LOCALAPPDATA%\WinTuner_Settings.json
```

### Settings File Structure
```json
{
  "DefaultPackagePath": "C:\\Packages",
  "AutoCheckUpdates": true,
  "RememberMe": true,
  "LastUser": "admin@contoso.com",
  "WingetOverrides": {
    "7zip.7zip": "24.07"
  }
}
```

### Logging
Logs are written to:
```
%LOCALAPPDATA%\WinTuner_GUI.log
```

Log includes:
- Login/logout events
- Package creation
- Deployment operations
- Update checks
- Rollback operations
- Errors and warnings

## 📸 Screenshots

### Main Interface - WinGet Apps Tab
*Search for WinGet packages and deploy to Intune*

### Updates Tab
*Check for and deploy updates to existing apps*

### Rollback Section
*Manage app versions and perform rollbacks*

### Settings Tab
*Configure application preferences*

## 🔧 Troubleshooting

### Common Issues

**Error: "This script requires PowerShell 7 or higher"**
- Download and install PowerShell 7 from: https://github.com/PowerShell/PowerShell
- Run script with `pwsh.exe` instead of `powershell.exe`

**Error: "Module WinTuner not found"**
- The GUI will prompt to install automatically
- Manual install: `Install-Module -Name WinTuner -Scope CurrentUser`

**Login fails or hangs**
- Ensure you have Intune Administrator permissions
- Check network connectivity
- Try closing and reopening the application

**Auto-check for updates not working**
- Ensure "Check for updates on login" is enabled in Settings
- Settings must be saved (click "Save Settings")
- Check log file for errors

**Rollback fails with "Cannot delete parent" error**
- This is a known Intune limitation with supersedence relationships
- The GUI handles this by deleting in correct order (oldest first)
- If still failing, manually delete versions in Intune portal

**Default package path resets to C:\Temp**
- Ensure you clicked "Save Settings" after changing path
- Check if settings file exists: `%LOCALAPPDATA%\WinTuner_Settings.json`
- Verify file permissions on AppData folder

## 📝 Changelog

See [CHANGELOG.md](CHANGELOG.md) for a detailed list of changes between versions.

### Latest Version: 2.0
- Complete rollback system for app versions
- Auto-check for updates on login
- Persistent settings system
- All UI text translated to English
- 10+ bug fixes and improvements

## 🙏 Credits

### WinTuner Module
This GUI is built on top of the excellent [WinTuner](https://github.com/svrooij/WinTuner) PowerShell module by:
- **[Sander Rozemuller](https://github.com/svrooij)** - WinTuner module creator

### GUI Development
- **Manuel Höfler** - WinTuner GUI development

### Special Thanks
- The WinGet community for package repository
- Microsoft Graph API team
- PowerShell community

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

### How to Contribute
1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📞 Support

For issues, questions, or suggestions:
- Open an [Issue](https://github.com/yourusername/WinTuner-GUI/issues)
- Check existing [Discussions](https://github.com/yourusername/WinTuner-GUI/discussions)
- Review the [Changelog](CHANGELOG.md)

## 🔗 Links

- [WinTuner Module](https://github.com/svrooij/WinTuner)
- [WinGet Repository](https://github.com/microsoft/winget-pkgs)
- [Microsoft Intune Documentation](https://learn.microsoft.com/en-us/mem/intune/)
- [PowerShell Documentation](https://learn.microsoft.com/en-us/powershell/)

---

<div align="center">
  <strong>Made with ❤️ for the Intune community</strong>
  <br>
  <sub>If this tool helped you, consider giving it a ⭐!</sub>
</div>
