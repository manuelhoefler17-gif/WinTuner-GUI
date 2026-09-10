# WinTuner GUI by Manuel Höfler
# v0.10.11 – Hotfix: $progressBar/$statusLabel/$outputBox als $script: Variablen für BackgroundWorker-Closures
# v0.10.10 – Hotfix: Invoke-AsyncOperation Closure-Bug – $progressBar war $null in RunWorkerCompleted
# v0.10.9 – Hotfix: ProgressBar Maximum-Reset an allen Stellen + graceful "not found" bei Remove
# v0.10.8 – Hotfix: Update-Check status feedback & checkUpdateButton re-enable after async
# v0.10.7 – Fix: Phase 5 – error handling, security, module import guard
# v0.10.6 – Fix: Phase 4 – performance improvements & code quality
# v0.10.5 – Fix: Phase 3 – UX improvements, ProgressBar crash hotfix, batch update summary
# v0.10.4 – Fix: Phase 2 – async update check, disconnect timeout, dead code removal
# v0.10.3 – Fix: Phase 1 critical bugfixes – error handling & logging consistency
# v0.10.2 – Fix: Remove updated apps immediately from update list
# v0.10.1 – Fix: Synchronize RememberMe checkboxes (login page ↔ Settings tab)
# v0.10.0 – Phase 6: Login/Logout improvements & recent users ComboBox
# --- PowerShell version gate (runs on PS<7 without parsing the main body) ---
try { $psMajor = $PSVersionTable.PSVersion.Major } catch { $psMajor = 0 }
if ($psMajor -lt 7) {
    try { Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue } catch {}
    [void][System.Windows.Forms.MessageBox]::Show(
        "This script requires PowerShell 7 or higher. Please upgrade your PowerShell version to continue.",
        "PowerShell Version Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    return
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Enable visual styles BEFORE creating controls
[System.Windows.Forms.Application]::EnableVisualStyles()

# Configure error handling for WinForms event handlers
# Suppress Write-* cmdlet errors that occur from non-pipeline threads
$WarningPreference = 'SilentlyContinue'
$InformationPreference = 'SilentlyContinue'
$VerbosePreference = 'SilentlyContinue'
$DebugPreference = 'SilentlyContinue'
$ProgressPreference = 'SilentlyContinue'

# Redirect all output streams to prevent threading issues
# Note: '*:ProgressAction' is intentionally omitted here — using a wildcard for ProgressAction
# can corrupt URI parameter binding in Invoke-WebRequest/Invoke-RestMethod on some PS7 builds.
# $ProgressPreference = 'SilentlyContinue' (set above) already suppresses progress output globally.
$PSDefaultParameterValues = @{
  '*:WarningAction' = 'SilentlyContinue'
  '*:InformationAction' = 'SilentlyContinue'
  '*:Verbose' = $false
  '*:Debug' = $false
}

# ============================================================
# Script configuration – central place for all script-scoped
# constants and mutable state variables
# ============================================================

# --- Application metadata ---
$script:appVersion  = "0.10.16"

# Bootstrap release dependencies and keep them synchronized with the GUI release.
$requiredReleaseFiles = @(
  'Modules/WinTuner.AppUpdate.psm1',
  'Modules/WinTuner.Core.psm1',
  'Modules/WinTuner.PackageBuild.psm1',
  'Modules/WinTuner.PackageUpload.psm1',
  'Modules/WinTuner.Winget.psm1',
  'Modules/WinTuner.UpdateScan.psm1',
  'Modules/WinTuner.Settings.psm1',
  'Modules/WinTuner.Logging.psm1',
  'Modules/WinTuner.Intune.psm1',
  'Workers/WinTuner.DiscoveryWorker.ps1'
)

$releaseMarkerPath = Join-Path $PSScriptRoot '.wintuner-release-version'
$installedReleaseVersion = $null

if (Test-Path $releaseMarkerPath) {
  try {
    $installedReleaseVersion = (Get-Content $releaseMarkerPath -Raw -ErrorAction Stop).Trim()
  } catch {
    $installedReleaseVersion = $null
  }
}

$isDevelopmentCheckout = Test-Path (Join-Path $PSScriptRoot '.git')

$missingReleaseFiles = @(
  $requiredReleaseFiles | Where-Object {
    -not (Test-Path (Join-Path $PSScriptRoot $_))
  }
)

if ($isDevelopmentCheckout -and $missingReleaseFiles.Count -gt 0) {
  try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue

    [void][System.Windows.Forms.MessageBox]::Show(
      "WinTuner development checkout is missing required local files:`n`n$($missingReleaseFiles -join "`n")",
      "WinTuner Development Error",
      [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Error
    )
  } catch {}

  return
}

$refreshAllReleaseFiles = (
  -not $isDevelopmentCheckout -and
  ($installedReleaseVersion -ne $script:appVersion)
)

$releaseFilesToDownload = if ($isDevelopmentCheckout) {
  @()
} elseif ($refreshAllReleaseFiles) {
  @($requiredReleaseFiles)
} else {
  @($missingReleaseFiles)
}

if ($releaseFilesToDownload.Count -gt 0) {
  try {
    $releaseTag = "v$($script:appVersion)"
    $rawBaseUrl = "https://raw.githubusercontent.com/manuelhoefler17-gif/WinTuner-GUI/$releaseTag"

    foreach ($relativePath in $releaseFilesToDownload) {
      $targetPath = Join-Path $PSScriptRoot $relativePath
      $targetDirectory = Split-Path -Parent $targetPath

      if (-not (Test-Path $targetDirectory)) {
        New-Item -ItemType Directory -Path $targetDirectory -Force | Out-Null
      }

      $downloadUrl = "$rawBaseUrl/$($relativePath -replace '\\','/')"
      $tempPath = "$targetPath.download-$PID-$([guid]::NewGuid().ToString('N'))"

      try {
        $savedDefaults = $PSDefaultParameterValues.Clone()

        try {
          $PSDefaultParameterValues = @{}

          Invoke-WebRequest `
            -Uri $downloadUrl `
            -OutFile $tempPath `
            -Headers @{ 'User-Agent' = "WinTuner-GUI/$($script:appVersion)" } `
            -TimeoutSec 30 `
            -UseBasicParsing `
            -ErrorAction Stop
        } finally {
          $PSDefaultParameterValues = $savedDefaults
        }

        if (-not (Test-Path $tempPath)) {
          throw "Bootstrap download did not create a file for $relativePath"
        }

        $downloadedFile = Get-Item $tempPath -ErrorAction Stop

        if ($downloadedFile.Length -lt 100) {
          throw "Bootstrap download for $relativePath is unexpectedly small ($($downloadedFile.Length) bytes)"
        }

        $downloadedContent = Get-Content $tempPath -Raw -ErrorAction Stop
        $parseTokens = $null
        $parseErrors = $null

        [void][System.Management.Automation.Language.Parser]::ParseInput(
          $downloadedContent,
          [ref]$parseTokens,
          [ref]$parseErrors
        )

        if ($parseErrors.Count -gt 0) {
          throw "Bootstrap download for $relativePath contains invalid PowerShell syntax: $($parseErrors[0].Message)"
        }

        Move-Item -Path $tempPath -Destination $targetPath -Force
      } catch {
        Remove-Item $tempPath -Force -ErrorAction SilentlyContinue
        throw
      }
    }

    # Only mark the release as installed after every dependency completed successfully.
    $markerTempPath = "$releaseMarkerPath.download-$PID-$([guid]::NewGuid().ToString('N'))"

    try {
      [System.IO.File]::WriteAllText(
        $markerTempPath,
        $script:appVersion,
        [System.Text.UTF8Encoding]::new($false)
      )

      Move-Item -Path $markerTempPath -Destination $releaseMarkerPath -Force
    } catch {
      Remove-Item $markerTempPath -Force -ErrorAction SilentlyContinue
      throw
    }
  } catch {
    try {
      Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue

      [void][System.Windows.Forms.MessageBox]::Show(
        "WinTuner GUI could not download required release files.`n`n$($_.Exception.Message)",
        "WinTuner Update Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
      )
    } catch {}

    return
  }
}
# Load WinTuner core helpers
$coreModulePath = Join-Path $PSScriptRoot 'Modules\WinTuner.Core.psm1'
Import-Module $coreModulePath -Force
# Load WinTuner WinGet helpers
$wingetModulePath = Join-Path $PSScriptRoot 'Modules\WinTuner.Winget.psm1'
Import-Module $wingetModulePath -Force

# Load WinTuner settings helpers
$settingsModulePath = Join-Path $PSScriptRoot 'Modules\WinTuner.Settings.psm1'
Import-Module $settingsModulePath -Force

# Load WinTuner logging helpers
$loggingModulePath = Join-Path $PSScriptRoot 'Modules\WinTuner.Logging.psm1'
Import-Module $loggingModulePath -Force

# Load WinTuner Intune / Graph helpers
$intuneModulePath = Join-Path $PSScriptRoot 'Modules\WinTuner.Intune.psm1'
Import-Module $intuneModulePath -Force

# Initialize file logging. The GUI output box is attached after it is created.
Initialize-WinTunerLogging -BasePath $PSScriptRoot
$script:repoOwner = "manuelhoefler17-gif"
$script:repoName = "WinTuner-GUI"
$script:githubRepo  = "$($script:repoOwner)/$($script:repoName)"
$script:githubApiUrl = "https://api.github.com/repos/$($script:repoOwner)/$($script:repoName)/releases/latest"
$script:skipLowValueWingetCandidates = $false  # keep all discovered apps; persistent caches handle repeated scans

# --- Runtime state (set during execution) ---
# $script:isConnected      – whether the user is logged in to a tenant
# $script:currentUserUpn   – UPN of the currently logged-in user
# $script:builtVersions    – caches effective or validated package versions per PackageId
# $script:wingetVersionCache – in-memory cache for winget version lookups
# $script:versionCachePath – path to the on-disk version cache JSON file
# $script:isDarkMode       – current theme state (true = dark)
# $script:currentTheme     – active theme hashtable (darkTheme or lightTheme)
# $script:asyncResult      – last result from Invoke-AsyncOperation
# $script:diskCache        – in-memory copy of the on-disk version cache (loaded once)
# $script:diskCacheLoaded  – whether $script:diskCache has been populated from disk

# Version comparison helper: returns $true if Latest > Current



function Test-AppUpdateAvailable {
  $result = [pscustomobject]@{
    UpdateAvailable = $false
    LatestVersion   = $null
    DownloadUrl     = $null
    HashUrl         = $null
    ReleaseUrl      = $null
    ReleaseNotes    = $null
    ErrorMessage    = $null
  }

  try {
    Write-Log "[Update] Prüfe auf neue Version..."

    $releaseInfo = Invoke-RestMethod -Uri $script:githubApiUrl -Method Get -ErrorAction Stop

    $latestVersionTag = [string]$releaseInfo.tag_name
    $latestVersionTag = $latestVersionTag -replace '[^0-9.]', ''
    if ([string]::IsNullOrWhiteSpace($latestVersionTag)) {
      throw "Release enthält keine gültige Versionsnummer (tag_name)."
    }

    $latestVersion = [version]$latestVersionTag
    $currentVersion = [version]$script:appVersion

    $result.LatestVersion = $latestVersion.ToString()
    $result.ReleaseUrl    = $releaseInfo.html_url
    $result.ReleaseNotes  = $releaseInfo.body

    $scriptFileName = if ($PSCommandPath) { [System.IO.Path]::GetFileName($PSCommandPath) } else { 'WinTuner_GUI.ps1' }
    $asset = $releaseInfo.assets | Where-Object { $_.name -like "*$scriptFileName*" } | Select-Object -First 1
    if (-not $asset) {
      $asset = $releaseInfo.assets | Where-Object { $_.name -like '*.ps1' } | Select-Object -First 1
    }
    if ($asset) { $result.DownloadUrl = $asset.browser_download_url }

    $shaAsset = $releaseInfo.assets | Where-Object { $_.name -like '*.sha256' } | Select-Object -First 1
    if ($shaAsset) { $result.HashUrl = $shaAsset.browser_download_url }

    if ($latestVersion -gt $currentVersion) {
      $result.UpdateAvailable = $true
      Write-Log "[*] Neue Version verfügbar: $latestVersion"
    } else {
      Write-Log "[√] Skript ist aktuell."
    }
  } catch {
    $result.ErrorMessage = $_.Exception.Message
    Write-Log "Update-Check konnte nicht durchgeführt werden: $($_.Exception.Message)"
  }

  return $result
}

function Invoke-AsyncUpdateCheck {
  param(
    [Parameter(Mandatory=$true)]
    [scriptblock]$OnComplete,

    [string]$StatusText = "Checking for updates..."
  )

  Update-Status $StatusText
  Write-Log "[Update] Prüfe asynchron auf neue Version..."

  $progressControl = $null
  if ($script:progressBar -is [System.Windows.Forms.ProgressBar] -and -not $script:progressBar.IsDisposed) {
    $progressControl = $script:progressBar
    try {
      $progressControl.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
      $progressControl.MarqueeAnimationSpeed = 30
      $progressControl.Visible = $true
    } catch {
      Write-LogSafe "Async update progress warning: $($_.Exception.Message)"
    }
  }

  $httpClient = [System.Net.Http.HttpClient]::new()
  $httpClient.Timeout = [TimeSpan]::FromSeconds(15)
  $httpClient.DefaultRequestHeaders.UserAgent.ParseAdd("WinTuner-GUI/$($script:appVersion)")

  $requestTask = $httpClient.GetStringAsync($script:githubApiUrl)

  # Capture primitive values before GetNewClosure creates its own script scope.
  $currentAppVersion = [string]$script:appVersion

  $timer = New-Object System.Windows.Forms.Timer
  $timer.Interval = 100

  $tickHandler = {
    param($sender, $e)

    if (-not $requestTask.IsCompleted) {
      return
    }

    $sender.Stop()

    $result = [pscustomobject]@{
      UpdateAvailable = $false
      LatestVersion   = $null
      DownloadUrl     = $null
      HashUrl         = $null
      ReleaseUrl      = $null
      ReleaseNotes    = $null
      ErrorMessage    = $null
    }

    try {
      if ($requestTask.IsCanceled) {
        throw "GitHub update request was canceled."
      }

      if ($requestTask.IsFaulted) {
        $exception = $requestTask.Exception

        if ($exception.InnerException) {
          throw $exception.InnerException
        }

        throw $exception
      }

      $json = $requestTask.GetAwaiter().GetResult()
      $releaseInfo = $json | ConvertFrom-Json -ErrorAction Stop

      $latestVersionTag = [string]$releaseInfo.tag_name
      $latestVersionTag = $latestVersionTag -replace '[^0-9.]', ''

      if ([string]::IsNullOrWhiteSpace($latestVersionTag)) {
        throw "Release enthält keine gültige Versionsnummer (tag_name)."
      }

      $latestVersion = [version]$latestVersionTag
      $currentVersion = [version]$currentAppVersion

      $result.LatestVersion = $latestVersion.ToString()
      $result.ReleaseUrl    = $releaseInfo.html_url
      $result.ReleaseNotes  = $releaseInfo.body

      $scriptFileName = if ($PSCommandPath) {
        [System.IO.Path]::GetFileName($PSCommandPath)
      } else {
        'WinTuner_GUI.ps1'
      }

      $asset = $releaseInfo.assets |
        Where-Object { $_.name -like "*$scriptFileName*" } |
        Select-Object -First 1

      if (-not $asset) {
        $asset = $releaseInfo.assets |
          Where-Object { $_.name -like '*.ps1' } |
          Select-Object -First 1
      }

      if ($asset) {
        $result.DownloadUrl = $asset.browser_download_url
      }

      $shaAsset = $releaseInfo.assets |
        Where-Object { $_.name -like '*.sha256' } |
        Select-Object -First 1

      if ($shaAsset) {
        $result.HashUrl = $shaAsset.browser_download_url
      }

      if ($latestVersion -gt $currentVersion) {
        $result.UpdateAvailable = $true
        Write-Log "[*] Neue Version verfügbar: $latestVersion"
      } else {
        Write-Log "[√] Skript ist aktuell."
      }
    } catch {
      $result.ErrorMessage = $_.Exception.Message
      Write-Log "Async update check failed: $($_.Exception.Message)"
    } finally {
      try {
        $httpClient.Dispose()
      } catch {}

      try {
        if ($progressControl -and -not $progressControl.IsDisposed) {
          $progressControl.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
          $progressControl.Value = 0
          $progressControl.Visible = $false
        }
      } catch {
        Write-LogSafe "Async update progress cleanup warning: $($_.Exception.Message)"
      }

      try {
        & $OnComplete $result
      } catch {
        Write-LogSafe "Async update completion callback error: $($_.Exception.Message)"
      }

      try {
        $sender.Dispose()
      } catch {}
    }
  }.GetNewClosure()

  $timer.Add_Tick($tickHandler)
  $timer.Start()
}
function Invoke-AppSelfUpdate {
  param(
    [Parameter(Mandatory=$true)]
    [string]$DownloadUrl,
    [string]$HashUrl = $null
  )

  $tempFile = $null
  $backupPath = $null
  $scriptReplaced = $false

  try {
    # Determine current script path
    $currentPath = $null
    if ($PSCommandPath) {
      $currentPath = $PSCommandPath
    } elseif ($MyInvocation.ScriptName) {
      $currentPath = $MyInvocation.ScriptName
    } else {
      $sfd = New-Object System.Windows.Forms.SaveFileDialog
      $sfd.Title = "Save updated WinTuner GUI"
      $sfd.Filter = "PowerShell Script (*.ps1)|*.ps1"
      $sfd.FileName = "WinTuner_GUI.ps1"

      if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $currentPath = $sfd.FileName
      } else {
        Write-Log "Update canceled: no save path selected"
        return $false
      }
    }

    $currentPath = [System.IO.Path]::GetFullPath($currentPath)

    Write-Log "Downloading update from: $DownloadUrl"
    Update-Status "Downloading update..."

    $tempFile = [System.IO.Path]::GetTempFileName() + ".ps1"

    # Temporarily clear PSDefaultParameterValues to prevent parameter binding conflicts.
    $savedDefaults = $PSDefaultParameterValues.Clone()
    try {
      $PSDefaultParameterValues = @{}
      $headers = @{ 'User-Agent' = 'WinTuner-GUI-UpdateCheck' }

      Invoke-WebRequest `
        -Uri $DownloadUrl `
        -OutFile $tempFile `
        -Headers $headers `
        -TimeoutSec 60 `
        -UseBasicParsing `
        -ErrorAction Stop
    } finally {
      $PSDefaultParameterValues = $savedDefaults
    }

    # Validate download.
    if (-not (Test-Path $tempFile)) {
      throw "Download failed: temp file not found"
    }

    $fileSize = (Get-Item $tempFile).Length

    if ($fileSize -lt 1000) {
      throw "Download failed: file too small ($fileSize bytes)"
    }

    $content = Get-Content $tempFile -Raw -ErrorAction Stop

    if ($content -notmatch 'WinTuner GUI') {
      throw "Download validation failed: file doesn't appear to be WinTuner GUI"
    }

    # Validate PowerShell syntax before replacing the running script.
    $parseTokens = $null
    $parseErrors = $null

    [void][System.Management.Automation.Language.Parser]::ParseInput(
      $content,
      [ref]$parseTokens,
      [ref]$parseErrors
    )

    if ($parseErrors.Count -gt 0) {
      throw "Downloaded update contains invalid PowerShell syntax: $($parseErrors[0].Message)"
    }

    # SHA256 integrity check.
    if ($HashUrl) {
      $hashMismatch = $false

      try {
        Write-Log "Verifying SHA256 integrity..."

        $savedDefaults2 = $PSDefaultParameterValues.Clone()

        try {
          $PSDefaultParameterValues = @{}

          $expectedHash = (
            Invoke-RestMethod `
              -Uri $HashUrl `
              -TimeoutSec 15 `
              -ErrorAction Stop
          ).Trim().ToUpper()
        } finally {
          $PSDefaultParameterValues = $savedDefaults2
        }

        $expectedHash = ($expectedHash -split '\s+')[0].ToUpper()
        $actualHash = (Get-FileHash $tempFile -Algorithm SHA256).Hash.ToUpper()

        if ($actualHash -ne $expectedHash) {
          $hashMismatch = $true
          throw "SHA256 mismatch: download may be corrupt or tampered! Expected: $expectedHash, Got: $actualHash"
        }

        Write-Log "SHA256 verified OK: $actualHash"
      } catch {
        if ($hashMismatch) {
          throw
        }

        Write-Log "Warning: SHA256 check skipped (could not fetch hash): $($_.Exception.Message)"
      }
    }

    Write-Log "Download complete ($fileSize bytes). Replacing script..."

    # Create backup. A valid backup is required before replacing the script.
    $backupPath = "$currentPath.backup"

    Copy-Item `
      -Path $currentPath `
      -Destination $backupPath `
      -Force `
      -ErrorAction Stop

    Write-Log "Backup created: $backupPath"

    # Replace current script.
    Move-Item `
      -Path $tempFile `
      -Destination $currentPath `
      -Force `
      -ErrorAction Stop

    $tempFile = $null
    $scriptReplaced = $true

    Write-Log "Script replaced successfully. Starting updated version..."
    Update-Status "Update installed. Restarting..."

    # Start the updated script using the same PowerShell 7 host.
    $pwshPath = Join-Path $PSHOME 'pwsh.exe'

    if (-not (Test-Path $pwshPath)) {
      throw "PowerShell 7 executable not found: $pwshPath"
    }

    $startInfo = [System.Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = $pwshPath
    $startInfo.UseShellExecute = $false
    [void]$startInfo.ArgumentList.Add('-NoProfile')
    [void]$startInfo.ArgumentList.Add('-ExecutionPolicy')
    [void]$startInfo.ArgumentList.Add('Bypass')
    [void]$startInfo.ArgumentList.Add('-File')
    [void]$startInfo.ArgumentList.Add($currentPath)

    $newProcess = [System.Diagnostics.Process]::Start($startInfo)

    if (-not $newProcess) {
      throw "Updated WinTuner process could not be started"
    }

    Write-Log "Updated WinTuner started successfully. PID: $($newProcess.Id)"

    return $true

  } catch {
    $updateError = $_.Exception.Message
    $rollbackMessage = $null

    Write-Log "Self-update failed: $updateError"

    if ($tempFile -and (Test-Path $tempFile)) {
      Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
    }

    if ($scriptReplaced) {
      if ($backupPath -and (Test-Path $backupPath)) {
        try {
          Write-Log "Update failed after script replacement. Rolling back from backup..."

          Copy-Item `
            -Path $backupPath `
            -Destination $currentPath `
            -Force `
            -ErrorAction Stop

          $rollbackMessage = "The previous WinTuner version was restored automatically."
          Write-Log "Rollback completed successfully: $currentPath"
        } catch {
          $rollbackError = $_.Exception.Message
          $rollbackMessage = "Automatic rollback also failed: $rollbackError"
          Write-Log "CRITICAL: Automatic rollback failed: $rollbackError"
        }
      } else {
        $rollbackMessage = "The script was replaced, but no backup was available for rollback."
        Write-Log "CRITICAL: Script was replaced but backup is unavailable."
      }
    }

    $errorMessage = "Update failed: $updateError"

    if ($rollbackMessage) {
      $errorMessage += "`n`n$rollbackMessage"
    }

    $errorMessage += "`n`nYou can update manually from:`nhttps://github.com/$($script:githubRepo)/releases/latest"

    [void][System.Windows.Forms.MessageBox]::Show(
      $errorMessage,
      "Update Failed",
      [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Error
    )

    return $false
  }
}

function Invoke-UpdateCheckFeedback {
  param(
    [object]$UpdateResult,
    [ValidateSet('Manual','Startup')]
    [string]$Context = 'Manual'
  )

  # Safe logger: works in normal scope and BackgroundWorker closures where Write-Log may not be visible
  $_SafeLog = {
    param([string]$Msg)
    $logCmd = Get-Command -Name 'Write-Log' -CommandType Function -ErrorAction SilentlyContinue
    if ($logCmd) { & $logCmd $Msg }
    else {
      try {
        $base = if ($PSScriptRoot) { $PSScriptRoot } else { [Environment]::GetFolderPath('LocalApplicationData') }
        $lp   = Join-Path $base 'WinTuner_GUI.log'
        Add-Content -Path $lp -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Msg" -Encoding utf8 -ErrorAction SilentlyContinue
      } catch {}
    }
  }

  $setStatus = {
    param([string]$Text)
    try {
      $statusCmd = Get-Command -Name 'Update-Status' -CommandType Function -ErrorAction SilentlyContinue
      if ($statusCmd) { & $statusCmd $Text }
      elseif (-not [string]::IsNullOrWhiteSpace($Text)) { & $_SafeLog "Status fallback: $Text" }
    } catch {}
  }.GetNewClosure()

  $isManual = ($Context -eq 'Manual')
  $errorDetail = if ($UpdateResult -and $UpdateResult.Error) { $UpdateResult.Error } `
                 elseif ($UpdateResult -and $UpdateResult.ErrorMessage) { $UpdateResult.ErrorMessage } `
                 else { $null }

  if ($errorDetail) {
    if ($isManual) {
      [System.Windows.Forms.MessageBox]::Show(
        "Could not check for updates.`n`nError: $errorDetail`n`nCheck your internet connection and try again.",
        "Update Check Failed",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
      )
    }
    & $setStatus "Update check failed (v$($script:appVersion)) – check internet connection"
    return
  }

  if ($UpdateResult -and $UpdateResult.LatestVersion -and (Test-IsNewerVersion -Latest $UpdateResult.LatestVersion -Current $script:appVersion)) {
    & $setStatus "Update available: v$($UpdateResult.LatestVersion)"
    try {
      $msg  = "A new version of WinTuner GUI is available!`n`n"
      $msg += "Current version: v$($script:appVersion)`n"
      $msg += "Latest version:  v$($UpdateResult.LatestVersion)`n`n"

      if ($UpdateResult.DownloadUrl) {
        $msg += "Do you want to download and install the update now?`n`n"
        $msg += "(A backup of your current version will be created)"

        $answer = [System.Windows.Forms.MessageBox]::Show(
          $msg,
          "Update Available",
          [System.Windows.Forms.MessageBoxButtons]::YesNo,
          [System.Windows.Forms.MessageBoxIcon]::Information
        )

        if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) {
          & $setStatus "Downloading update..."
          [System.Windows.Forms.Application]::DoEvents()

          $success = Invoke-AppSelfUpdate -DownloadUrl $UpdateResult.DownloadUrl -HashUrl $UpdateResult.HashUrl

          if ($success) {
            & $setStatus "Update installed successfully. Restarting WinTuner..."
            $form.Close()
          } else {
            & $setStatus "Update download/install failed. See log for details."
          }
        } else {
          if ($isManual) {
            & $setStatus "Update postponed by user"
          } else {
            & $setStatus "Update available: v$($UpdateResult.LatestVersion) - Go to Settings to update later."
          }
        }
      } else {
        & $setStatus "Update available: v$($UpdateResult.LatestVersion) (manual download required)"
        $msg += "No direct download available for this release.`n"
        $msg += "Please download manually from:`n$($UpdateResult.ReleaseUrl)"

        [System.Windows.Forms.MessageBox]::Show(
          $msg,
          "Update Available",
          [System.Windows.Forms.MessageBoxButtons]::OK,
          [System.Windows.Forms.MessageBoxIcon]::Information
        )
      }
    } catch {
      Write-Log "$Context update dialog error: $($_.Exception.Message)"
      & $setStatus "Update check completed (dialog error). See log for details."
    }
    return
  }

  $latestVer = if ($UpdateResult -and $UpdateResult.LatestVersion) { $UpdateResult.LatestVersion } else { $null }
  if (-not $latestVer) {
    # Retry once synchronously when async callback returned incomplete/empty payload.
    # NOTE: this closure may run inside a BackgroundWorker RunWorkerCompleted delegate where
    # Write-Log is not in scope. Use $_SafeLog (captured via GetNewClosure) for all logging here.
    if (-not ($UpdateResult -and $UpdateResult.ErrorMessage)) {
      & $_SafeLog "Update check returned no version and no error. Retrying once synchronously..."
      $retryResult = $null
      $retryError  = $null
      try { $retryResult = Test-AppUpdateAvailable } catch { $retryError = $_.Exception.Message }
      if ($retryError) {
        & $_SafeLog "Synchronous retry failed: $retryError"
      } elseif ($retryResult -and $retryResult.LatestVersion) {
        & $_SafeLog "Synchronous retry succeeded. GitHub version: v$($retryResult.LatestVersion)"
        # Re-enter with the full fresh result so the UpdateAvailable branch above handles it correctly
        Invoke-UpdateCheckFeedback -UpdateResult $retryResult -Context $Context
        return
      } elseif ($retryResult -and $retryResult.ErrorMessage) {
        $UpdateResult = $retryResult
      }
    }
  }

  $resolvedUpdateAvailable = $false
  if ($latestVer) {
    $resolvedUpdateAvailable = Test-IsNewerVersion -Latest $latestVer -Current $script:appVersion
  }

  if (-not $latestVer) {
    $errText = if ($UpdateResult -and $UpdateResult.ErrorMessage) { $UpdateResult.ErrorMessage } else { 'No version information returned by GitHub.' }
    Write-Log "Update check could not resolve GitHub version: $errText"
    & $setStatus "Update check failed (GitHub version unavailable). Local: v$($script:appVersion)"
  } elseif ($resolvedUpdateAvailable) {
    & $setStatus "Update available: v$latestVer"
  } else {
    $statusMsg = "Up to date – Local: v$($script:appVersion) | GitHub: v$latestVer"
    & $setStatus $statusMsg
  }

  if ($isManual) {
    if (-not $resolvedUpdateAvailable) {
      [System.Windows.Forms.MessageBox]::Show(
        "WinTuner GUI is up to date.`n`nLocal version:  v$($script:appVersion)`nGitHub version: v$(if ($latestVer) { $latestVer } else { 'unavailable' })",
        "No Update Available",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
      )
    }
    # Note: if $resolvedUpdateAvailable is true here without $UpdateResult.UpdateAvailable,
    # the version was detected via Test-IsNewerVersion but DownloadUrl may be missing.
    # The status bar already shows "Update available" — user can use the Check Update button.
  }
}

# Helper: resolve Winget Package Identifier across possible property names
function Resolve-WtWingetId {
    param([object]$AppOrResult)

    if (-not $AppOrResult) { return $null }
    # Do NOT consider Graph 'Id' as Winget Id
    foreach ($prop in 'PackageId','PackageID','WingetId','PackageIdentifier') {
        $p = $AppOrResult.PSObject.Properties[$prop]
        if ($p -and $AppOrResult.$prop) { return [string]$AppOrResult.$prop }
    }
    if ($AppOrResult -is [hashtable]) {
        foreach ($prop in 'PackageId','PackageID','WingetId','PackageIdentifier') {
            if ($AppOrResult.ContainsKey($prop) -and $AppOrResult[$prop]) { return [string]$AppOrResult[$prop] }
        }
    }
    return $null
}

function Resolve-WingetIdForApp {
  param([object]$App)
  $id = Resolve-WtWingetId -AppOrResult $App
  if (-not [string]::IsNullOrWhiteSpace($id)) { return $id }
  try {
    $res = @(Search-WtWinGetPackage -SearchQuery $App.Name)
  } catch { $res = @() }
  if ($res -and $res.Count -gt 0) {
    # Avoid Select-Object -First 1 to prevent WinForms pipeline crash; use array index [0] instead
    $exact = @($res | Where-Object { $_.Name -and ($_.Name -eq $App.Name) })
    if ($exact.Count -gt 0 -and $exact[0].PackageID) { return [string]$exact[0].PackageID }
    
    $first = $res[0]
    if ($first -and $first.PackageID) { return [string]$first.PackageID }
  }
  return $null
}









function Get-StringSimilarity {
  param($str1, $str2)
  if (-not $str1 -or -not $str2) { return 0 }
  $clean1 = $str1.ToLower() -replace '[^\w\s]', ' '
  $clean2 = $str2.ToLower() -replace '[^\w\s]', ' '
  $words1 = @($clean1 -split '\s+' | Where-Object { $_.Trim() -ne '' })
  $words2 = @($clean2 -split '\s+' | Where-Object { $_.Trim() -ne '' })
  if ($words1.Count -eq 0 -or $words2.Count -eq 0) { return 0 }

  $matchCount = 0
  foreach ($w in $words1) { if ($words2 -contains $w) { $matchCount++ } }
  $minWords = [math]::Min($words1.Count, $words2.Count)
  return [math]::Round(($matchCount / $minWords) * 100)
}


function Show-VersionPickerDialog {
  param([string]$Title,[string[]]$Versions)
  $f = New-Object System.Windows.Forms.Form
  $f.Text = $Title
  $f.Size = New-Object System.Drawing.Size(400,500)
  $lb = New-Object System.Windows.Forms.ListBox
  $lb.Location = New-Object System.Drawing.Point(10,10)
  $lb.Size = New-Object System.Drawing.Size(360,400)
  foreach ($v in @($Versions)) { [void]$lb.Items.Add($v) }
  if ($lb.Items.Count -gt 0) { $lb.SelectedIndex = 0 }
  $ok = New-Object System.Windows.Forms.Button
  $ok.Text = 'OK'
  $ok.Location = New-Object System.Drawing.Point(210,420)
  $ok.Add_Click({ $f.Tag = $lb.SelectedItem; $f.DialogResult = [System.Windows.Forms.DialogResult]::OK; $f.Close() })
  $cancel = New-Object System.Windows.Forms.Button
  $cancel.Text = 'Cancel'
  $cancel.Location = New-Object System.Drawing.Point(290,420)
  $cancel.Add_Click({ $f.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $f.Close() })
  $f.Controls.Add($lb)
  $f.Controls.Add($ok)
  $f.Controls.Add($cancel)
  [void]$f.ShowDialog()
  return [string]$f.Tag
}

function New-WingetPackageWithFallback {
  param(
    [string]$PackageId,
    [string]$PackageFolder,
    [string]$DesiredVersion,
    [string]$LatestVersion,
    [string]$InstalledVersion,
    [switch]$AllowUserRetry
  )
  $attemptVersion = $DesiredVersion
  if (-not $attemptVersion) { $attemptVersion = $LatestVersion }
  try {
    if ($attemptVersion) { New-WtWingetPackage -PackageId $PackageId -PackageFolder $PackageFolder -Version $attemptVersion -ErrorAction Stop }
    else { New-WtWingetPackage -PackageId $PackageId -PackageFolder $PackageFolder -ErrorAction Stop }
    return [pscustomobject]@{ Succeeded=$true; EffectiveVersion=$attemptVersion }
  } catch {
    $m = $_.Exception.Message
    if ($m -match '404' -or $m -match 'Not Found') {
      $prev = Get-PreviousWingetVersion -PackageId $PackageId -LatestVersion $attemptVersion
      # Only allow previous if it's newer than current tenant version (if known)
      if ($prev -and ( -not $InstalledVersion -or (Test-IsNewerVersion $prev $InstalledVersion) )) {
        try { New-WtWingetPackage -PackageId $PackageId -PackageFolder $PackageFolder -Version $prev -ErrorAction Stop; return [pscustomobject]@{ Succeeded=$true; EffectiveVersion=$prev } } catch { return [pscustomobject]@{ Succeeded=$false; EffectiveVersion=$null; ErrorMessage=$_.Exception.Message } }
      } else { return [pscustomobject]@{ Succeeded=$false; EffectiveVersion=$null; ErrorMessage=$m } }
    } elseif ($m -match 'Hash mismatch') {
      if ($AllowUserRetry) {
        $res = [System.Windows.Forms.MessageBox]::Show("Hash mismatch detected. Retry download? Click Yes to retry, No to use previous available version, or Cancel to abort.", "Hash mismatch", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel, [System.Windows.Forms.MessageBoxIcon]::Warning, [System.Windows.Forms.MessageBoxDefaultButton]::Button1)
        if ($res -eq [System.Windows.Forms.DialogResult]::Yes) {
          try { if ($attemptVersion) { New-WtWingetPackage -PackageId $PackageId -PackageFolder $PackageFolder -Version $attemptVersion -ErrorAction Stop } else { New-WtWingetPackage -PackageId $PackageId -PackageFolder $PackageFolder -ErrorAction Stop }; return [pscustomobject]@{ Succeeded=$true; EffectiveVersion=$attemptVersion } } catch { return [pscustomobject]@{ Succeeded=$false; EffectiveVersion=$null; ErrorMessage=$_.Exception.Message } }
        } elseif ($res -eq [System.Windows.Forms.DialogResult]::No) {
          $latest = $attemptVersion; if (-not $latest) { $latest = $LatestVersion }
          $prev = Get-PreviousWingetVersion -PackageId $PackageId -LatestVersion $latest
          # Only allow previous if it's newer than current tenant version (if known)
          if ($prev -and ( -not $InstalledVersion -or (Test-IsNewerVersion $prev $InstalledVersion) )) {
            try { New-WtWingetPackage -PackageId $PackageId -PackageFolder $PackageFolder -Version $prev -ErrorAction Stop; return [pscustomobject]@{ Succeeded=$true; EffectiveVersion=$prev } } catch { return [pscustomobject]@{ Succeeded=$false; EffectiveVersion=$null; ErrorMessage=$_.Exception.Message } }
          } else { return [pscustomobject]@{ Succeeded=$false; EffectiveVersion=$null; ErrorMessage=$m } }
        } else { return [pscustomobject]@{ Succeeded=$false; EffectiveVersion=$null; ErrorMessage="Cancelled by user" } }
      } else { return [pscustomobject]@{ Succeeded=$false; EffectiveVersion=$null; ErrorMessage=$m } }
    } else { return [pscustomobject]@{ Succeeded=$false; EffectiveVersion=$null; ErrorMessage=$m } }
  }
}

# Build authoritative update candidate list using winget verification when available
function Get-UpdateCandidates {
  $all = @()
  try {
    $all = @(Get-WtWin32Apps -Superseded:$false -ErrorAction Stop)
    Write-Log ("Get-WtWin32Apps (all) returned {0} item(s)" -f ($all | Measure-Object).Count)
  } catch {
    Write-Log ("Get-WtWin32Apps (all) failed: {0}" -f $_.Exception.Message)
    $all = @()
  }
  $candidates = [System.Collections.Generic.List[object]]::new()
  foreach ($app in @($all)) {
    if (-not $app -or -not $app.CurrentVersion) { continue }
    $wingetId = Resolve-WingetIdForApp -App $app
    $usedLatest = $null
    if ($wingetId) {
      try { $wgVersions = @(Get-WingetVersions -PackageId $wingetId) } catch { $wgVersions = @() }
      if ($wgVersions -and $wgVersions.Count -gt 0) { $usedLatest = $wgVersions[0] }
    }
    if (-not $usedLatest) { $usedLatest = $app.LatestVersion }
    if ($usedLatest -and (Test-IsNewerVersion $usedLatest $app.CurrentVersion)) {
      try { 
        $app.LatestVersion = $usedLatest 
      } catch {
        Write-Log "Warning: Could not update LatestVersion property for $($app.Name): $($_.Exception.Message)"
      }
      $candidates.Add($app)
    }
  }
  return @($candidates)
}

# Dark mode theme colors
$script:darkTheme = @{
  BackColor       = [System.Drawing.Color]::FromArgb(32, 32, 32)
  ForeColor       = [System.Drawing.Color]::FromArgb(255, 255, 255)
  ButtonBackColor = [System.Drawing.Color]::FromArgb(64, 64, 64)
  ButtonForeColor = [System.Drawing.Color]::FromArgb(255, 255, 255)
  TextBoxBackColor= [System.Drawing.Color]::FromArgb(48, 48, 48)
  TextBoxForeColor= [System.Drawing.Color]::FromArgb(255, 255, 255)
  TabBackColor    = [System.Drawing.Color]::FromArgb(40, 40, 40)
  TabForeColor    = [System.Drawing.Color]::FromArgb(255, 255, 255)
}
# Light mode theme colors
$script:lightTheme = @{
  BackColor       = [System.Drawing.Color]::FromArgb(240, 240, 240)
  ForeColor       = [System.Drawing.Color]::FromArgb(0, 0, 0)
  ButtonBackColor = [System.Drawing.Color]::FromArgb(225, 225, 225)
  ButtonForeColor = [System.Drawing.Color]::FromArgb(0, 0, 0)
  TextBoxBackColor= [System.Drawing.Color]::FromArgb(255, 255, 255)
  TextBoxForeColor= [System.Drawing.Color]::FromArgb(0, 0, 0)
  TabBackColor    = [System.Drawing.Color]::FromArgb(240, 240, 240)
  TabForeColor    = [System.Drawing.Color]::FromArgb(0, 0, 0)
}

$script:isDarkMode   = $true
$script:currentTheme = $script:darkTheme

# Function to apply theme to all controls
function Set-GuiTheme {
  param([System.Windows.Forms.Control]$control, [hashtable]$theme)
  if ($control -is [System.Windows.Forms.Form]) {
    $control.BackColor = $theme.BackColor
    $control.ForeColor = $theme.ForeColor
  }
  elseif ($control -is [System.Windows.Forms.Button]) {
    $control.BackColor = $theme.ButtonBackColor
    $control.ForeColor = $theme.ButtonForeColor
    $control.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $control.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(100,100,100)
  }
  elseif ($control -is [System.Windows.Forms.TextBox]) {
    $control.BackColor = $theme.TextBoxBackColor
    $control.ForeColor = $theme.TextBoxForeColor
  }
  elseif ($control -is [System.Windows.Forms.ComboBox]) {
    $control.BackColor = $theme.TextBoxBackColor
    $control.ForeColor = $theme.TextBoxForeColor
    $control.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  }
  elseif ($control -is [System.Windows.Forms.Label]) {
    $control.BackColor = [System.Drawing.Color]::Transparent
    $control.ForeColor = $theme.ForeColor
  }
  elseif ($control -is [System.Windows.Forms.TabControl]) {
    $control.BackColor = $theme.TabBackColor
    $control.ForeColor = $theme.TabForeColor
  }
  elseif ($control -is [System.Windows.Forms.TabPage]) {
    $control.BackColor = $theme.BackColor
    $control.ForeColor = $theme.ForeColor
  }
  elseif ($control -is [System.Windows.Forms.ProgressBar]) {
    $control.BackColor = $theme.TextBoxBackColor
  }
  else {
    $control.BackColor = $theme.BackColor
    $control.ForeColor = $theme.ForeColor
  }
  foreach ($childControl in @($control.Controls)) {
    Set-GuiTheme -control $childControl -theme $theme
  }
}

# Function to toggle theme
function Switch-GuiTheme {
  $script:isDarkMode   = -not $script:isDarkMode
  $script:currentTheme = if ($script:isDarkMode) { $script:darkTheme } else { $script:lightTheme }
  Set-GuiTheme -control $form -theme $script:currentTheme
  # Button text indicates the action (target)
  $themeToggleButton.Text = if ($script:isDarkMode) { "Light Mode" } else { "Dark Mode" }
  $form.Refresh()
}

# Logging function (thread-safe for WinForms event handlers)


# Logging helper that never throws if Write-Log is unavailable in delegate scopes


# Runs an action on the UI thread if required
function Invoke-UiAction {
  param(
    [Parameter(Mandatory=$true)]
    [System.Windows.Forms.Control]$Control,
    [Parameter(Mandatory=$true)]
    [scriptblock]$Action
  )

  if (-not $Control) { return }
  if ($Control.IsDisposed) { return }

  if ($Control.InvokeRequired) {
    $Control.Invoke([Action]$Action)
  } else {
    & $Action
  }
}

# Status update function
function Update-Status {
  param([string]$status)
  $statusText = if ([string]::IsNullOrWhiteSpace($status)) { "" } else { $status }
  try {
    Invoke-UiAction -Control $script:statusLabel -Action {
      $script:statusLabel.Text = $statusText
    }
  } catch {
    # Keep status updates non-fatal even on cross-thread/disposed-control races
  }
  try {
    $safeLogger = Get-Command -Name Write-LogSafe -CommandType Function -ErrorAction SilentlyContinue
    if ($safeLogger) {
      & $safeLogger $status
    }
  } catch {}
}

# Async operation helper - runs long operations in background
function Invoke-AsyncOperation {
  <#
  .SYNOPSIS
    Executes a long-running operation in background without blocking UI
  .PARAMETER ScriptBlock
    The script block to execute
  .PARAMETER OnComplete
    Script block to execute when operation completes (runs on UI thread)
  .PARAMETER StatusText
    Status text to display
  .PARAMETER DisableControls
    Array of controls to disable during operation
  #>
  param(
    [Parameter(Mandatory=$true)]
    [scriptblock]$ScriptBlock,
    [scriptblock]$OnComplete,
    [string]$StatusText = "Processing...",
    [System.Windows.Forms.Control[]]$DisableControls = @()
  )

  # Capture a stable ProgressBar reference for async callbacks (avoid script-scope resolution drift)
  $progressControl = $null
  if ($script:progressBar -is [System.Windows.Forms.ProgressBar] -and -not $script:progressBar.IsDisposed) {
    $progressControl = $script:progressBar
  }
  
  # Update UI - show progress in marquee style (indefinite)
  Update-Status $StatusText
  if ($progressControl) {
    try {
      $progressControl.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
      $progressControl.MarqueeAnimationSpeed = 30
      $progressControl.Visible = $true
    } catch {
      Write-LogSafe "Async start progress warning: $($_.Exception.Message)"
    }
  }
  
  # Disable controls
  foreach ($ctrl in $DisableControls) {
    if ($ctrl) { $ctrl.Enabled = $false }
  }
  
  # Create background worker
  $bw = New-Object System.ComponentModel.BackgroundWorker
  $bw.WorkerReportsProgress = $false
  $bw.WorkerSupportsCancellation = $false
  
  # Store result
  $script:asyncResult = $null
  
  # Lokale Kopien der Parameter für Closure-Binding
  $_ScriptBlock     = $ScriptBlock
  $_OnComplete      = $OnComplete
  $_DisableControls = $DisableControls
  $_SafeLog         = {
    param([string]$Msg)
    if ([string]::IsNullOrWhiteSpace($Msg)) { return }
    try {
      $safeLogger = Get-Command -Name Write-LogSafe -CommandType Function -ErrorAction SilentlyContinue
      if ($safeLogger) {
        & $safeLogger $Msg
        return
      }
    } catch {}
    try {
      $base = if ($PSScriptRoot) { $PSScriptRoot } else { [Environment]::GetFolderPath('LocalApplicationData') }
      if (-not (Test-Path $base)) { New-Item -ItemType Directory -Path $base -Force | Out-Null }
      $logPath = Join-Path $base 'WinTuner_GUI.log'
      $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
      Add-Content -Path $logPath -Value "$timestamp - $Msg" -Encoding utf8 -ErrorAction SilentlyContinue
    } catch {}
  }.GetNewClosure()

  # Do work in background
  $doWork = {
    param($sender, $e)
    try {
      $rawResult = & $_ScriptBlock
      # Serialize the result to JSON so it survives the BackgroundWorker thread boundary intact.
      # PSCustomObject properties are silently dropped when crossing managed-thread boundaries
      # unless the object is serialized first. Primitives and $null pass through unchanged.
      if ($null -eq $rawResult) {
        $e.Result = $null
      } else {
        try { $e.Result = $rawResult | ConvertTo-Json -Depth 10 -Compress -ErrorAction Stop }
        catch { $e.Result = $rawResult }   # fallback: pass as-is for primitive types
      }
    } catch {
      $errMsg = $_.Exception.Message
      $e.Result = (@{ Error = $errMsg } | ConvertTo-Json -Compress)
      try {
        $base = if ($PSScriptRoot) { $PSScriptRoot } else { [Environment]::GetFolderPath('LocalApplicationData') }
        $logPath = Join-Path $base 'WinTuner_GUI.log'
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Add-Content -Path $logPath -Value "$timestamp - Async operation error: $errMsg" -Encoding utf8 -ErrorAction SilentlyContinue
      } catch {}
    }
  }.GetNewClosure()
  $bw.Add_DoWork($doWork)

  # On completion (runs on UI thread)
  # .GetNewClosure() wird genutzt, um $_OnComplete und $_DisableControls sowie
  # die stabile $progressControl-Referenz sicher in die Delegate-Scopes zu übernehmen.
  $runCompleted = {
    param($sender, $e)
    try {
      try {
        if ($progressControl -and -not $progressControl.IsDisposed) {
          $progressControl.Style   = [System.Windows.Forms.ProgressBarStyle]::Continuous
          $progressControl.Maximum = 100
          $progressControl.Value   = 100
        }
      } catch {
        & $_SafeLog "Async completion UI reset warning: $($_.Exception.Message)"
      }

      # Execute completion callback with result
      if ($_OnComplete) {
        try {
          # Deserialize the JSON-encoded result (serialized in DoWork to survive thread boundary)
          $callbackResult = $e.Result
          if ($callbackResult -is [string] -and $callbackResult -match '^\s*[\{\[]') {
            try { $callbackResult = $callbackResult | ConvertFrom-Json -ErrorAction Stop } catch {}
          }
          & $_OnComplete $callbackResult
        } catch {
          & $_SafeLog "Async completion callback error: $($_.Exception.Message)"
          # Update-Status cannot be called directly inside a BackgroundWorker closure
          # because PowerShell delegates do not inherit the parent session's function table.
          # Use Get-Command lookup (the same pattern used for Write-LogSafe/_SafeLog above).
          try {
            $statusCmd = Get-Command -Name 'Update-Status' -CommandType Function -ErrorAction SilentlyContinue
            if ($statusCmd) { & $statusCmd "Operation completed with errors" }
          } catch {}
        }
      }

      # Hide progress after short delay
      try {
        $hideTimer = New-Object System.Windows.Forms.Timer
        $hideTimer.Interval = 1000
        $hideTimer.Add_Tick({
          param($sender, $e)
          try {
            if ($progressControl -and -not $progressControl.IsDisposed) {
                $progressControl.Visible = $false
                $progressControl.Value = 0
              }
          } catch {
            & $_SafeLog "Async completion timer tick warning: $($_.Exception.Message)"
          } finally {
            try { $sender.Stop() } catch {}
            try { $sender.Dispose() } catch {}
          }
        })
        $hideTimer.Start()
      } catch {
        & $_SafeLog "Async completion timer warning: $($_.Exception.Message)"
      }
    } finally {
      # Re-enable controls even if callback/UI reset throws
      foreach ($ctrl in $_DisableControls) {
        if ($ctrl) { $ctrl.Enabled = $true }
      }
      # Dispose the BackgroundWorker to prevent memory leaks
      $sender.Dispose()
    }
  }.GetNewClosure()
  $bw.Add_RunWorkerCompleted($runCompleted)

  # Start async operation
  $bw.RunWorkerAsync()
}

function Request-WinTunerUpdateScanCancellation {
  if (-not $script:updateScanRunning -or -not $script:updateScanContext) { return }
  if ($script:cancelUpdateScan) { return }

  $script:cancelUpdateScan = $true
  try {
    [System.IO.File]::WriteAllText(
      $script:updateScanContext.CancelPath,
      (Get-Date).ToString('O'),
      [System.Text.UTF8Encoding]::new($false)
    )
  } catch {
    Write-Log "Update scan cancellation signal failed: $($_.Exception.Message)"
  }

  Update-Status 'Cancel requested - finishing current WinGet query...'
  Write-Log 'Update scan cancellation requested.'
  Update-UpdateActionState
}

function Complete-WinTunerUpdateScan {
  param([Parameter(Mandatory=$true)][object]$Context)

  $scanResult = $null
  $completionError = $null

  try {
    $output = @($Context.PowerShell.EndInvoke($Context.AsyncResult))
    if ($output.Count -eq 0) {
      if ($Context.PowerShell.Streams.Error.Count -gt 0) {
        throw $Context.PowerShell.Streams.Error[0].Exception
      }
      throw 'The update scan returned no result.'
    }

    $json = [string]$output[$output.Count - 1]
    $scanResult = $json | ConvertFrom-Json -ErrorAction Stop
  } catch {
    $completionError = $_.Exception.Message
  } finally {
    try { $Context.Timer.Stop() } catch {}
    try { $Context.Timer.Dispose() } catch {}
    try { $Context.PowerShell.Dispose() } catch {}
    Remove-Item -LiteralPath $Context.ProgressPath, $Context.CancelPath -Force -ErrorAction SilentlyContinue

    if ($script:updateScanContext -eq $Context) {
      $script:updateScanContext = $null
    }
    $script:updateScanRunning = $false
    $script:cancelUpdateScan = $false
    $script:isUpdateOperationActive = $false

    try {
      $script:progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
      $script:progressBar.Maximum = 100
      $script:progressBar.Value = 0
      $script:progressBar.Visible = $false
    } catch {}
  }

  if ($completionError) {
    Update-Status "Update scan failed: $completionError"
    Write-Log "Update scan failed: $completionError"
    Update-UpdateActionState
    return
  }

  if ($scanResult.ErrorMessage) {
    Update-Status "Update scan failed: $($scanResult.ErrorMessage)"
    Write-Log "Update scan failed: $($scanResult.ErrorMessage)"
    Update-UpdateActionState
    return
  }

  if ($scanResult.Canceled) {
    Update-Status "Update scan canceled | Checked: $($scanResult.CheckedApps)/$($scanResult.TotalApps)"
    Write-Log "Update scan canceled after $($scanResult.CheckedApps) of $($scanResult.TotalApps) apps; partial results discarded."
    Update-UpdateActionState
    return
  }

  $script:updateApps = [System.Collections.Generic.List[object]]::new()
  $script:updateVisibleApps = [System.Collections.Generic.List[object]]::new()
  $updateListBox.BeginUpdate()
  try {
    $updateListBox.Items.Clear()
    foreach ($app in @($scanResult.Candidates | Sort-Object Name)) {
      if (-not $app -or [string]::IsNullOrWhiteSpace([string]$app.Name)) { continue }
      $app.Checked = $false
      [void]$updateListBox.Items.Add((Get-UpdateCandidateDisplayText -App $app))
      [void]$script:updateApps.Add($app)
      [void]$script:updateVisibleApps.Add($app)
      Write-Log ("Update available: {0} ({1} -> {2})" -f $app.Name, $app.CurrentVersion, $app.LatestVersion)
    }
  } finally {
    $updateListBox.EndUpdate()
  }

  if ([int]$scanResult.LookupFailureCount -gt 0) {
    Write-Log "Update scan WinGet lookup failures: $($scanResult.LookupFailureCount); available tenant versions were used as fallback where possible."
  }

  $candidateCount = $script:updateApps.Count
  Update-Status "Update scan complete | Checked: $($scanResult.CheckedApps) | Candidates: $candidateCount"
  Write-Log "Update scan summary -> Checked: $($scanResult.CheckedApps), Candidates: $candidateCount"
  Update-UpdateActionState
}

function Start-WinTunerUpdateScan {
  if ($script:updateScanRunning -or $script:isUpdateOperationActive -or $script:isPackageSearchActive -or $script:isVersionLookupActive -or $script:isPackageBuildActive -or $script:isPackageUploadActive -or $script:isSupersededOperationActive -or $script:discoveryScanRunning -or $script:isDiscoveryDeploymentActive) { return }

  $scanModulePath = Join-Path $PSScriptRoot 'Modules\WinTuner.UpdateScan.psm1'
  if (-not (Test-Path -LiteralPath $scanModulePath -PathType Leaf)) {
    Update-Status 'Update scan failed: update scan module is missing.'
    Write-Log "Update scan module missing: $scanModulePath"
    return
  }

  $operationId = [guid]::NewGuid().ToString('N')
  $progressPath = Join-Path ([System.IO.Path]::GetTempPath()) "wintuner-update-scan-$operationId.progress.json"
  $cancelPath = Join-Path ([System.IO.Path]::GetTempPath()) "wintuner-update-scan-$operationId.cancel"

  $updateFilterBox.Text = ''
  $updateListBox.Items.Clear()
  $script:updateApps = [System.Collections.Generic.List[object]]::new()
  $script:updateVisibleApps = [System.Collections.Generic.List[object]]::new()
  $script:isUpdateOperationActive = $true
  $script:updateScanRunning = $true
  $script:cancelUpdateScan = $false

  $script:progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
  $script:progressBar.Minimum = 0
  $script:progressBar.Maximum = 1
  $script:progressBar.Value = 0
  $script:progressBar.Visible = $true
  Update-Status 'Loading apps from Intune...'
  Write-Log 'Starting asynchronous update scan.'
  Update-UpdateActionState

  $scanScript = @'
param($RepositoryRoot, $ProgressPath, $CancelPath)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

Import-Module WinTuner -ErrorAction Stop
Import-Module (Join-Path $RepositoryRoot 'Modules\WinTuner.Core.psm1') -Force -ErrorAction Stop
Import-Module (Join-Path $RepositoryRoot 'Modules\WinTuner.Winget.psm1') -Force -ErrorAction Stop
Import-Module (Join-Path $RepositoryRoot 'Modules\WinTuner.UpdateScan.psm1') -Force -ErrorAction Stop

function Write-UpdateScanProgressFile {
  param([object]$ProgressInfo)
  $tempPath = "$ProgressPath.$([guid]::NewGuid().ToString('N')).tmp"
  try {
    $json = $ProgressInfo | ConvertTo-Json -Depth 4 -Compress
    [System.IO.File]::WriteAllText($tempPath, $json, [System.Text.UTF8Encoding]::new($false))
    Move-Item -LiteralPath $tempPath -Destination $ProgressPath -Force
  } finally {
    Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
  }
}

$result = Invoke-WinTunerUpdateScan `
  -GetApps { @(Get-WtWin32Apps -Superseded:$false -ErrorAction Stop) } `
  -ResolvePackageId {
    param($app)
    foreach ($propertyName in 'PackageId','PackageID','WingetId','PackageIdentifier') {
      $property = $app.PSObject.Properties[$propertyName]
      if ($property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) {
        return [string]$property.Value
      }
    }

    $matches = @(Search-WtWinGetPackage -SearchQuery $app.Name -ErrorAction SilentlyContinue)
    $exact = @($matches | Where-Object { $_.Name -and $_.Name -eq $app.Name })
    if ($exact.Count -gt 0 -and $exact[0].PackageID) { return [string]$exact[0].PackageID }
    if ($matches.Count -gt 0 -and $matches[0].PackageID) { return [string]$matches[0].PackageID }
    return $null
  } `
  -GetVersions { param($packageId) @(Get-WingetVersions -PackageId $packageId -ErrorAction Stop) } `
  -IsNewerVersion { param($latest, $current) Test-IsNewerVersion -Latest $latest -Current $current } `
  -ShouldCancel { Test-Path -LiteralPath $CancelPath } `
  -ReportProgress { param($progressInfo) Write-UpdateScanProgressFile -ProgressInfo $progressInfo }

$result | ConvertTo-Json -Depth 6 -Compress
'@

  $powerShell = [System.Management.Automation.PowerShell]::Create()
  $null = $powerShell.AddScript($scanScript).AddArgument($PSScriptRoot).AddArgument($progressPath).AddArgument($cancelPath)

  $timer = New-Object System.Windows.Forms.Timer
  $timer.Interval = 150
  $context = [pscustomobject]@{
    PowerShell   = $powerShell
    AsyncResult = $null
    Timer        = $timer
    ProgressPath = $progressPath
    CancelPath   = $cancelPath
    LastProgress = $null
  }
  $script:updateScanContext = $context

  $timer.Add_Tick({
    $currentContext = $script:updateScanContext
    if (-not $currentContext) { return }

    if (Test-Path -LiteralPath $currentContext.ProgressPath -PathType Leaf) {
      try {
        $progressJson = Get-Content -LiteralPath $currentContext.ProgressPath -Raw -ErrorAction Stop
        if ($progressJson -and $progressJson -ne $currentContext.LastProgress) {
          $currentContext.LastProgress = $progressJson
          $progressInfo = $progressJson | ConvertFrom-Json -ErrorAction Stop
          if ($progressInfo.Stage -eq 'Checking') {
            $maximum = [Math]::Max(1, [int]$progressInfo.Total)
            $script:progressBar.Maximum = $maximum
            $script:progressBar.Value = [Math]::Min([int]$progressInfo.Processed, $maximum)
            Update-Status ("Checking ({0}/{1}): {2}" -f $progressInfo.Processed, $progressInfo.Total, $progressInfo.AppName)
          }
        }
      } catch {}
    }

    if ($currentContext.AsyncResult -and $currentContext.AsyncResult.IsCompleted) {
      Complete-WinTunerUpdateScan -Context $currentContext
    }
  })

  try {
    $context.AsyncResult = $powerShell.BeginInvoke()
    $timer.Start()
  } catch {
    try { $timer.Dispose() } catch {}
    try { $powerShell.Dispose() } catch {}
    Remove-Item -LiteralPath $progressPath, $cancelPath -Force -ErrorAction SilentlyContinue
    $script:updateScanContext = $null
    $script:updateScanRunning = $false
    $script:cancelUpdateScan = $false
    $script:isUpdateOperationActive = $false
    $script:progressBar.Visible = $false
    Update-Status "Update scan failed to start: $($_.Exception.Message)"
    Write-Log "Update scan failed to start: $($_.Exception.Message)"
    Update-UpdateActionState
  }
}

function Complete-WinTunerAppUpdateBatch {
  param([Parameter(Mandatory=$true)][object]$Context)
  $batchResult = $null
  $completionError = $null
  try {
    $output = @($Context.PowerShell.EndInvoke($Context.AsyncResult))
    if ($output.Count -eq 0) {
      if ($Context.PowerShell.Streams.Error.Count -gt 0) { throw $Context.PowerShell.Streams.Error[0].Exception }
      throw 'The background update returned no result.'
    }
    $batchResult = ([string]$output[$output.Count - 1]) | ConvertFrom-Json -ErrorAction Stop
  } catch { $completionError = $_.Exception.Message }
  finally {
    try { $Context.Timer.Stop() } catch {}
    try { $Context.Timer.Dispose() } catch {}
    try { $Context.PowerShell.Dispose() } catch {}
    Remove-Item -LiteralPath $Context.ProgressPath -Force -ErrorAction SilentlyContinue
    if ($script:updateDeploymentContext -eq $Context) { $script:updateDeploymentContext = $null }
    $script:isUpdateOperationActive = $false
    try {
      $script:progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
      $script:progressBar.Maximum = 100
      $script:progressBar.Value = 0
      $script:progressBar.Visible = $false
    } catch {}
  }
  if ($completionError) {
    Update-Status "App update failed: $completionError"
    Write-Log "Background app update failed: $completionError"
    Update-UpdateActionState
    Update-DiscoveryActionState
    return
  }
  $failedItems = [System.Collections.Generic.List[object]]::new()
  foreach ($itemResult in @($batchResult.Results)) {
    if ($itemResult.Succeeded) {
      Write-Log "Successfully updated in background: $($itemResult.Name) ($($itemResult.CurrentVersion) -> $($itemResult.EffectiveVersion))"
      $cachedApp = $null
      foreach ($candidate in @($script:updateApps)) {
        $graphMatches = (-not [string]::IsNullOrWhiteSpace([string]$itemResult.GraphId) -and [string]$candidate.GraphId -eq [string]$itemResult.GraphId)
        $packageMatches = ([string]::IsNullOrWhiteSpace([string]$itemResult.GraphId) -and [string]$candidate.PackageId -eq [string]$itemResult.PackageId -and [string]$candidate.Name -eq [string]$itemResult.Name)
        if ($graphMatches -or $packageMatches) { $cachedApp = $candidate; break }
      }
      if ($cachedApp) {
        $visibleIndex = $script:updateVisibleApps.IndexOf($cachedApp)
        if ($visibleIndex -ge 0) {
          $script:updateVisibleApps.RemoveAt($visibleIndex)
          $updateListBox.Items.RemoveAt($visibleIndex)
        }
        [void]$script:updateApps.Remove($cachedApp)
      }
    } else {
      Write-Log "Background update failed for $($itemResult.Name) ($($itemResult.ReasonCode)): $($itemResult.Message)"
      $failedItems.Add($itemResult)
    }
  }
  if ($failedItems.Count -gt 0) {
    $summary = "The following $($failedItems.Count) app(s) could not be updated:$([Environment]::NewLine)$([Environment]::NewLine)"
    $summary += ($failedItems | ForEach-Object { "• $($_.Name): $($_.Message)" }) -join [Environment]::NewLine
    [void][System.Windows.Forms.MessageBox]::Show($summary, "Update Summary – $($failedItems.Count) Failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
  }
  $successCount = [int]$batchResult.SuccessCount
  $failureCount = [int]$batchResult.FailureCount
  $statusPrefix = if ($Context.Mode -eq 'Checked') { 'Checked apps updated' } else { 'All updates completed' }
  Update-Status ("{0}: {1} successful, {2} failed" -f $statusPrefix, $successCount, $failureCount)
  Write-Log "Background update summary -> Mode: $($Context.Mode), Successful: $successCount, Failed: $failureCount"
  Update-UpdateActionState
  Update-DiscoveryActionState
}

function Start-WinTunerAppUpdateBatch {
  param(
    [Parameter(Mandatory=$true)][object[]]$Apps,
    [Parameter(Mandatory=$true)][string]$RootPackageFolder,
    [ValidateSet('Checked','All')][string]$Mode
  )
  if (-not $script:isConnected -or $script:isUpdateOperationActive -or $script:isPackageSearchActive -or $script:isVersionLookupActive -or $script:isPackageBuildActive -or $script:isPackageUploadActive -or $script:isSupersededOperationActive -or $script:discoveryScanRunning -or $script:isDiscoveryDeploymentActive) {
    Update-UpdateActionState
    return
  }
  $items = @($Apps)
  if ($items.Count -eq 0) {
    Update-Status 'No update candidates were provided.'
    Update-UpdateActionState
    return
  }
  $updateModulePath = Join-Path $PSScriptRoot 'Modules\WinTuner.AppUpdate.psm1'
  if (-not (Test-Path -LiteralPath $updateModulePath -PathType Leaf)) {
    Update-Status 'App update failed: update module is missing.'
    Write-Log "App update module missing: $updateModulePath"
    return
  }
  $workerApps = [System.Collections.Generic.List[object]]::new()
  foreach ($app in $items) {
    $workerApps.Add([pscustomobject]@{
      Name = [string]$app.Name
      CurrentVersion = [string]$app.CurrentVersion
      LatestVersion = [string]$app.LatestVersion
      GraphId = [string]$app.GraphId
      PackageId = [string]$app.PackageId
    })
  }
  $appsJson = @($workerApps) | ConvertTo-Json -Depth 5 -Compress
  $operationId = [guid]::NewGuid().ToString('N')
  $progressPath = Join-Path ([System.IO.Path]::GetTempPath()) "wintuner-app-update-$operationId.progress.json"
  $script:isUpdateOperationActive = $true
  $script:progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
  $script:progressBar.Minimum = 0
  $script:progressBar.Maximum = [Math]::Max(1, $workerApps.Count)
  $script:progressBar.Value = 0
  $script:progressBar.Visible = $true
  Update-Status "Starting background update for $($workerApps.Count) app(s)..."
  Write-Log "Starting background app update -> Mode: $Mode, Apps: $($workerApps.Count)"
  Update-UpdateActionState
  Update-DiscoveryActionState
  Update-SupersededActionState
  $updateScript = @'
param($RepositoryRoot, $AppsJson, $RootPackageFolder, $ProgressPath)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
Import-Module WinTuner -ErrorAction Stop
Import-Module (Join-Path $RepositoryRoot 'Modules\WinTuner.Core.psm1') -Force -ErrorAction Stop
Import-Module (Join-Path $RepositoryRoot 'Modules\WinTuner.Winget.psm1') -Force -ErrorAction Stop
Import-Module (Join-Path $RepositoryRoot 'Modules\WinTuner.PackageBuild.psm1') -Force -ErrorAction Stop
Import-Module (Join-Path $RepositoryRoot 'Modules\WinTuner.AppUpdate.psm1') -Force -ErrorAction Stop
function Write-AppUpdateProgressFile {
  param([object]$ProgressInfo)
  $tempPath = "$ProgressPath.$([guid]::NewGuid().ToString('N')).tmp"
  try {
    $json = $ProgressInfo | ConvertTo-Json -Depth 4 -Compress
    [System.IO.File]::WriteAllText($tempPath, $json, [System.Text.UTF8Encoding]::new($false))
    Move-Item -LiteralPath $tempPath -Destination $ProgressPath -Force
  } finally { Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue }
}
$apps = @($AppsJson | ConvertFrom-Json -ErrorAction Stop)
$result = Invoke-WinTunerAppUpdateBatch -Apps $apps -RootPackageFolder $RootPackageFolder -ReportProgress {
  param($progressInfo)
  Write-AppUpdateProgressFile -ProgressInfo $progressInfo
}
$result | ConvertTo-Json -Depth 7 -Compress
'@
  $powerShell = [System.Management.Automation.PowerShell]::Create()
  $null = $powerShell.AddScript($updateScript).AddArgument($PSScriptRoot).AddArgument($appsJson).AddArgument($RootPackageFolder).AddArgument($progressPath)
  $timer = New-Object System.Windows.Forms.Timer
  $timer.Interval = 150
  $context = [pscustomobject]@{ PowerShell = $powerShell; AsyncResult = $null; Timer = $timer; ProgressPath = $progressPath; LastProgress = $null; Mode = $Mode }
  $script:updateDeploymentContext = $context
  $timer.Add_Tick({
    $currentContext = $script:updateDeploymentContext
    if (-not $currentContext) { return }
    if (Test-Path -LiteralPath $currentContext.ProgressPath -PathType Leaf) {
      try {
        $progressJson = Get-Content -LiteralPath $currentContext.ProgressPath -Raw -ErrorAction Stop
        if ($progressJson -and $progressJson -ne $currentContext.LastProgress) {
          $currentContext.LastProgress = $progressJson
          $progressInfo = $progressJson | ConvertFrom-Json -ErrorAction Stop
          $maximum = [Math]::Max(1, [int]$progressInfo.Total)
          $script:progressBar.Maximum = $maximum
          $script:progressBar.Value = [Math]::Min([int]$progressInfo.Processed, $maximum)
          Update-Status ("Updating ({0}/{1}): {2}" -f $progressInfo.Processed, $progressInfo.Total, $progressInfo.AppName)
        }
      } catch {}
    }
    if ($currentContext.AsyncResult -and $currentContext.AsyncResult.IsCompleted) { Complete-WinTunerAppUpdateBatch -Context $currentContext }
  })
  try {
    $context.AsyncResult = $powerShell.BeginInvoke()
    $timer.Start()
  } catch {
    try { $timer.Dispose() } catch {}
    try { $powerShell.Dispose() } catch {}
    Remove-Item -LiteralPath $progressPath -Force -ErrorAction SilentlyContinue
    $script:updateDeploymentContext = $null
    $script:isUpdateOperationActive = $false
    $script:progressBar.Visible = $false
    Update-Status "App update failed to start: $($_.Exception.Message)"
    Write-Log "Background app update failed to start: $($_.Exception.Message)"
    Update-UpdateActionState
    Update-DiscoveryActionState
  }
}

function Complete-WinTunerSupersededSearch {
  param([Parameter(Mandatory=$true)][object]$Context)

  $searchResult = $null
  $completionError = $null

  try {
    $output = @($Context.PowerShell.EndInvoke($Context.AsyncResult))
    if ($output.Count -eq 0) {
      if ($Context.PowerShell.Streams.Error.Count -gt 0) {
        throw $Context.PowerShell.Streams.Error[0].Exception
      }
      throw 'The superseded-app search returned no result.'
    }

    $json = [string]$output[$output.Count - 1]
    $searchResult = $json | ConvertFrom-Json -ErrorAction Stop
  } catch {
    $completionError = $_.Exception.Message
  } finally {
    try { $Context.Timer.Stop() } catch {}
    try { $Context.Timer.Dispose() } catch {}
    try { $Context.PowerShell.Dispose() } catch {}
    if ($script:supersededSearchContext -eq $Context) {
      $script:supersededSearchContext = $null
    }
    $script:isSupersededOperationActive = $false
    try {
      $script:progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
      $script:progressBar.Value = 0
      $script:progressBar.Visible = $false
    } catch {}
  }

  if ($completionError) {
    Update-Status "Superseded-app search failed: $completionError"
    Write-Log "Superseded-app search failed: $completionError"
    Update-UpdateActionState
    Update-DiscoveryActionState
    Update-SupersededActionState
    return
  }

  $script:supersededApps = @()
  $supersededDropdown.BeginUpdate()
  try {
    $supersededDropdown.Items.Clear()
    foreach ($app in @($searchResult.Apps | Sort-Object Name, CurrentVersion)) {
      if (-not $app -or [string]::IsNullOrWhiteSpace([string]$app.Name) -or [string]::IsNullOrWhiteSpace([string]$app.GraphId)) { continue }
      $script:supersededApps += $app
      [void]$supersededDropdown.Items.Add("$($app.Name) — $($app.CurrentVersion)")
    }
    if ($supersededDropdown.Items.Count -gt 0) {
      $supersededDropdown.SelectedIndex = 0
    }
  } finally {
    $supersededDropdown.EndUpdate()
  }

  Update-Status ("Search completed: {0} superseded Apps found." -f $script:supersededApps.Count)
  Write-Log "Asynchronous superseded-app search completed: $($script:supersededApps.Count) result(s)."
  Update-UpdateActionState
  Update-DiscoveryActionState
  Update-SupersededActionState
}

function Start-WinTunerSupersededSearch {
  if (
    -not $script:isConnected -or
    $script:isSupersededOperationActive -or
    $script:isPackageSearchActive -or
    $script:isVersionLookupActive -or
    $script:isPackageBuildActive -or
    $script:isPackageUploadActive -or
    $script:isUpdateOperationActive -or
    $script:discoveryScanRunning -or
    $script:isDiscoveryDeploymentActive
  ) {
    Update-SupersededActionState
    return
  }

  $script:supersededApps = @()
  $supersededDropdown.Items.Clear()
  $script:isSupersededOperationActive = $true
  $script:progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
  $script:progressBar.MarqueeAnimationSpeed = 30
  $script:progressBar.Visible = $true
  Update-Status 'Search for superseded apps...'
  Write-Log 'Starting asynchronous superseded-app search.'
  Update-UpdateActionState
  Update-DiscoveryActionState
  Update-SupersededActionState

  $searchScript = @(
    '$ErrorActionPreference = ''Stop'''
    '$ProgressPreference = ''SilentlyContinue'''
    'Import-Module WinTuner -ErrorAction Stop'
    '$apps = @('
    '  Get-WtWin32Apps -Superseded:$true -ErrorAction Stop | ForEach-Object {'
    '    [pscustomobject]@{'
    '      Name = [string]$_.Name'
    '      CurrentVersion = [string]$_.CurrentVersion'
    '      GraphId = [string]$_.GraphId'
    '    }'
    '  }'
    ')'
    '[pscustomobject]@{ Apps = $apps } | ConvertTo-Json -Depth 4 -Compress'
  ) -join [Environment]::NewLine

  $powerShell = [System.Management.Automation.PowerShell]::Create()
  $null = $powerShell.AddScript($searchScript)
  $timer = New-Object System.Windows.Forms.Timer
  $timer.Interval = 150
  $context = [pscustomobject]@{
    PowerShell  = $powerShell
    AsyncResult = $null
    Timer       = $timer
  }
  $script:supersededSearchContext = $context

  $timer.Add_Tick({
    $currentContext = $script:supersededSearchContext
    if ($currentContext -and $currentContext.AsyncResult -and $currentContext.AsyncResult.IsCompleted) {
      Complete-WinTunerSupersededSearch -Context $currentContext
    }
  })

  try {
    $context.AsyncResult = $powerShell.BeginInvoke()
    $timer.Start()
  } catch {
    try { $timer.Dispose() } catch {}
    try { $powerShell.Dispose() } catch {}
    $script:supersededSearchContext = $null
    $script:isSupersededOperationActive = $false
    $script:progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $script:progressBar.Value = 0
    $script:progressBar.Visible = $false
    Update-Status "Superseded-app search failed to start: $($_.Exception.Message)"
    Write-Log "Superseded-app search failed to start: $($_.Exception.Message)"
    Update-UpdateActionState
    Update-DiscoveryActionState
    Update-SupersededActionState
  }
}
function Complete-WinTunerPackageSearch {
  param([Parameter(Mandatory=$true)][object]$Context)

  $searchResult = $null
  $completionError = $null
  try {
    $output = @($Context.PowerShell.EndInvoke($Context.AsyncResult))
    if ($output.Count -eq 0) {
      if ($Context.PowerShell.Streams.Error.Count -gt 0) {
        throw $Context.PowerShell.Streams.Error[0].Exception
      }
      throw 'The WinGet package search returned no result.'
    }

    $json = [string]$output[$output.Count - 1]
    $searchResult = $json | ConvertFrom-Json -ErrorAction Stop
  } catch {
    $completionError = $_.Exception.Message
  } finally {
    try { $Context.Timer.Stop() } catch {}
    try { $Context.Timer.Dispose() } catch {}
    try { $Context.PowerShell.Dispose() } catch {}
    if ($script:packageSearchContext -eq $Context) {
      $script:packageSearchContext = $null
    }
    $script:isPackageSearchActive = $false
    try {
      $script:progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
      $script:progressBar.Value = 0
      $script:progressBar.Visible = $false
    } catch {}
  }

  if ($completionError) {
    Update-Status "WinGet search failed: $completionError"
    Write-Log "WinGet package search failed for '$($Context.Query)': $completionError"
    Update-PackageActionState
    Update-UpdateActionState
    Update-DiscoveryActionState
    Update-SupersededActionState
    Update-PackageSearchActionState
    return
  }

  $dropdown.BeginUpdate()
  try {
    $dropdown.Items.Clear()
    $script:packageMap.Clear()
    foreach ($result in @($searchResult.Results)) {
      $name = [string]$result.Name
      $packageId = [string]$result.PackageID
      if ([string]::IsNullOrWhiteSpace($name) -or [string]::IsNullOrWhiteSpace($packageId)) { continue }
      $displayText = "$name — $packageId"
      [void]$dropdown.Items.Add($displayText)
      $script:packageMap[$displayText] = @{
        PackageID = $packageId
        Version   = [string]$result.Version
      }
    }
    if ($dropdown.Items.Count -gt 0) {
      $dropdown.SelectedIndex = 0
    }
  } finally {
    $dropdown.EndUpdate()
  }

  if ($dropdown.Items.Count -eq 0) {
    Update-Status "No results found for '$($Context.Query)'"
  } else {
    Update-Status "Search completed: $($dropdown.Items.Count) package(s) found."
  }
  Write-Log "Asynchronous WinGet package search completed for '$($Context.Query)': $($dropdown.Items.Count) result(s)."
  Update-PackageActionState
  Update-UpdateActionState
  Update-DiscoveryActionState
  Update-SupersededActionState
  Update-PackageSearchActionState
}

function Start-WinTunerPackageSearch {
  $query = [string]$appSearchBox.Text
  if ([string]::IsNullOrWhiteSpace($query)) {
    [void][System.Windows.Forms.MessageBox]::Show(
      "App search can't be empty.",
      'Validation',
      [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Information
    )
    return
  }
  $query = $query.Trim()

  if (
    $script:isPackageSearchActive -or
    $script:isVersionLookupActive -or
    $script:isPackageBuildActive -or
    $script:isPackageUploadActive -or
    $script:isUpdateOperationActive -or
    $script:discoveryScanRunning -or
    $script:isDiscoveryDeploymentActive -or
    $script:isSupersededOperationActive
  ) {
    Update-PackageSearchActionState
    return
  }

  $dropdown.Items.Clear()
  $script:packageMap.Clear()
  Update-PackageActionState
  $script:isPackageSearchActive = $true
  $script:progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
  $script:progressBar.MarqueeAnimationSpeed = 30
  $script:progressBar.Visible = $true
  Update-Status "Searching WinGet for '$query'..."
  Write-Log "Starting asynchronous WinGet package search for '$query'."
  Update-UpdateActionState
  Update-DiscoveryActionState
  Update-SupersededActionState
  Update-PackageSearchActionState

  $searchScript = @(
    'param($Query)'
    '$ErrorActionPreference = ''Stop'''
    '$ProgressPreference = ''SilentlyContinue'''
    'Import-Module WinTuner -ErrorAction Stop'
    '$results = @('
    '  Search-WtWinGetPackage -SearchQuery $Query -ErrorAction Stop | ForEach-Object {'
    '    [pscustomobject]@{'
    '      Name = [string]$_.Name'
    '      PackageID = [string]$_.PackageID'
    '      Version = [string]$_.Version'
    '    }'
    '  }'
    ')'
    '[pscustomobject]@{ Results = $results } | ConvertTo-Json -Depth 4 -Compress'
  ) -join [Environment]::NewLine

  $powerShell = [System.Management.Automation.PowerShell]::Create()
  $null = $powerShell.AddScript($searchScript).AddArgument($query)
  $timer = New-Object System.Windows.Forms.Timer
  $timer.Interval = 150
  $context = [pscustomobject]@{
    PowerShell  = $powerShell
    AsyncResult = $null
    Timer       = $timer
    Query       = $query
  }
  $script:packageSearchContext = $context

  $timer.Add_Tick({
    $currentContext = $script:packageSearchContext
    if ($currentContext -and $currentContext.AsyncResult -and $currentContext.AsyncResult.IsCompleted) {
      Complete-WinTunerPackageSearch -Context $currentContext
    }
  })

  try {
    $context.AsyncResult = $powerShell.BeginInvoke()
    $timer.Start()
  } catch {
    try { $timer.Dispose() } catch {}
    try { $powerShell.Dispose() } catch {}
    $script:packageSearchContext = $null
    $script:isPackageSearchActive = $false
    $script:progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $script:progressBar.Value = 0
    $script:progressBar.Visible = $false
    Update-Status "WinGet search failed to start: $($_.Exception.Message)"
    Write-Log "WinGet package search failed to start for '$query': $($_.Exception.Message)"
    Update-PackageActionState
    Update-UpdateActionState
    Update-DiscoveryActionState
    Update-SupersededActionState
    Update-PackageSearchActionState
  }
}
function Complete-WinTunerVersionLookup {
  param([Parameter(Mandatory=$true)][object]$Context)

  $lookupResult = $null
  $completionError = $null
  try {
    $output = @($Context.PowerShell.EndInvoke($Context.AsyncResult))
    if ($output.Count -eq 0) {
      if ($Context.PowerShell.Streams.Error.Count -gt 0) {
        throw $Context.PowerShell.Streams.Error[0].Exception
      }
      throw 'The WinGet version lookup returned no result.'
    }

    $json = [string]$output[$output.Count - 1]
    $lookupResult = $json | ConvertFrom-Json -ErrorAction Stop
  } catch {
    $completionError = $_.Exception.Message
  } finally {
    try { $Context.Timer.Stop() } catch {}
    try { $Context.Timer.Dispose() } catch {}
    try { $Context.PowerShell.Dispose() } catch {}
    if ($script:versionLookupContext -eq $Context) {
      $script:versionLookupContext = $null
    }
    try {
      $script:progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
      $script:progressBar.Value = 0
      $script:progressBar.Visible = $false
    } catch {}
  }

  if ($completionError) {
    $script:isVersionLookupActive = $false
    Update-Status "Version lookup failed: $completionError"
    Write-Log "WinGet version lookup failed for '$($Context.PackageId)': $completionError"
    Update-PackageActionState
    Update-UpdateActionState
    Update-DiscoveryActionState
    Update-SupersededActionState
    Update-PackageSearchActionState
    return
  }

  $versions = @($lookupResult.Versions | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
  if ($versions.Count -eq 0) {
    $script:isVersionLookupActive = $false
    Update-Status 'No versions found for the selected package.'
    Write-Log "WinGet version lookup returned no versions for '$($Context.PackageId)'."
    Update-PackageActionState
    Update-UpdateActionState
    Update-DiscoveryActionState
    Update-SupersededActionState
    Update-PackageSearchActionState
    return
  }

  Write-Log "WinGet version lookup completed for '$($Context.PackageId)': $($versions.Count) version(s)."
  try {
    $chosen = Show-VersionPickerDialog -Title ("Select version for {0}" -f $Context.PackageId) -Versions $versions
    if ($chosen) {
      $script:selectedPackageVersions[$Context.PackageId] = $chosen
      Update-Status ("Selected version for {0}: {1}" -f $Context.PackageId, $chosen)
    } else {
      Update-Status 'Version selection canceled.'
    }
  } catch {
    Update-Status "Version selection failed: $($_.Exception.Message)"
    Write-Log "Version selection dialog failed for '$($Context.PackageId)': $($_.Exception.Message)"
  } finally {
    $script:isVersionLookupActive = $false
    Update-PackageActionState
    Update-UpdateActionState
    Update-DiscoveryActionState
    Update-SupersededActionState
    Update-PackageSearchActionState
  }
}

function Start-WinTunerVersionLookup {
  $selectedDisplay = [string]$dropdown.SelectedItem
  if ([string]::IsNullOrWhiteSpace($selectedDisplay)) {
    Update-Status 'Please select a package first.'
    return
  }

  $package = $script:packageMap[$selectedDisplay]
  if (-not $package -or [string]::IsNullOrWhiteSpace([string]$package.PackageID)) {
    Update-Status 'Selected item is invalid.'
    Update-PackageSearchActionState
    return
  }

  if (
    $script:isPackageSearchActive -or
    $script:isVersionLookupActive -or
    $script:isPackageBuildActive -or
    $script:isPackageUploadActive -or
    $script:isUpdateOperationActive -or
    $script:discoveryScanRunning -or
    $script:isDiscoveryDeploymentActive -or
    $script:isSupersededOperationActive
  ) {
    Update-PackageSearchActionState
    return
  }

  $packageId = [string]$package.PackageID
  $script:isVersionLookupActive = $true
  $script:progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
  $script:progressBar.MarqueeAnimationSpeed = 30
  $script:progressBar.Visible = $true
  Update-Status "Loading versions for '$packageId'..."
  Write-Log "Starting asynchronous WinGet version lookup for '$packageId'."
  Update-UpdateActionState
  Update-DiscoveryActionState
  Update-SupersededActionState
  Update-PackageSearchActionState

  $lookupScript = @(
    'param($RepositoryRoot, $PackageId)'
    '$ErrorActionPreference = ''Stop'''
    '$ProgressPreference = ''SilentlyContinue'''
    'Import-Module (Join-Path $RepositoryRoot ''Modules\WinTuner.Winget.psm1'') -Force -ErrorAction Stop'
    '$versions = @(Get-WingetVersions -PackageId $PackageId -ErrorAction Stop | ForEach-Object { [string]$_ })'
    '[pscustomobject]@{ Versions = $versions } | ConvertTo-Json -Depth 3 -Compress'
  ) -join [Environment]::NewLine

  $powerShell = [System.Management.Automation.PowerShell]::Create()
  $null = $powerShell.AddScript($lookupScript).AddArgument($PSScriptRoot).AddArgument($packageId)
  $timer = New-Object System.Windows.Forms.Timer
  $timer.Interval = 150
  $context = [pscustomobject]@{
    PowerShell  = $powerShell
    AsyncResult = $null
    Timer       = $timer
    PackageId   = $packageId
  }
  $script:versionLookupContext = $context

  $timer.Add_Tick({
    $currentContext = $script:versionLookupContext
    if ($currentContext -and $currentContext.AsyncResult -and $currentContext.AsyncResult.IsCompleted) {
      Complete-WinTunerVersionLookup -Context $currentContext
    }
  })

  try {
    $context.AsyncResult = $powerShell.BeginInvoke()
    $timer.Start()
  } catch {
    try { $timer.Dispose() } catch {}
    try { $powerShell.Dispose() } catch {}
    $script:versionLookupContext = $null
    $script:isVersionLookupActive = $false
    $script:progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $script:progressBar.Value = 0
    $script:progressBar.Visible = $false
    Update-Status "Version lookup failed to start: $($_.Exception.Message)"
    Write-Log "WinGet version lookup failed to start for '$packageId': $($_.Exception.Message)"
    Update-PackageActionState
    Update-UpdateActionState
    Update-DiscoveryActionState
    Update-SupersededActionState
    Update-PackageSearchActionState
  }
}
function Complete-WinTunerPackageBuild {
  param([Parameter(Mandatory=$true)][object]$Context)

  $buildResult = $null
  $completionError = $null
  try {
    $output = @($Context.PowerShell.EndInvoke($Context.AsyncResult))
    if ($output.Count -eq 0) {
      if ($Context.PowerShell.Streams.Error.Count -gt 0) {
        throw $Context.PowerShell.Streams.Error[0].Exception
      }
      throw 'The package build returned no result.'
    }

    $json = [string]$output[$output.Count - 1]
    $buildResult = $json | ConvertFrom-Json -ErrorAction Stop
  } catch {
    $completionError = $_.Exception.Message
  } finally {
    try { $Context.Timer.Stop() } catch {}
    try { $Context.Timer.Dispose() } catch {}
    try { $Context.PowerShell.Dispose() } catch {}
    if ($script:packageBuildContext -eq $Context) {
      $script:packageBuildContext = $null
    }
    try {
      $script:progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
      $script:progressBar.Value = 0
      $script:progressBar.Visible = $false
    } catch {}
  }

  if ($completionError) {
    $script:isPackageBuildActive = $false
    Update-Status "Package creation failed: $completionError"
    Write-Log "Background package build failed for '$($Context.PackageId)': $completionError"
    Update-PackageActionState
    Update-UpdateActionState
    Update-DiscoveryActionState
    Update-SupersededActionState
    Update-PackageSearchActionState
    return
  }

  if ([string]$buildResult.ChoiceRequired -eq 'HashMismatch') {
    Write-Log "Package build for '$($Context.PackageId)' requires a hash-mismatch decision."
    $answer = [System.Windows.Forms.MessageBox]::Show(
      'Hash mismatch detected. Retry download? Click Yes to retry, No to use the previous available version, or Cancel to abort.',
      'Hash mismatch',
      [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
      [System.Windows.Forms.MessageBoxIcon]::Warning,
      [System.Windows.Forms.MessageBoxDefaultButton]::Button1
    )

    $script:isPackageBuildActive = $false
    if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) {
      Start-WinTunerPackageBuild -PackageId $Context.PackageId -PackageFolder $Context.PackageFolder -DesiredVersion $Context.DesiredVersion -LatestVersion $Context.LatestVersion -Mode RetrySame
      return
    }
    if ($answer -eq [System.Windows.Forms.DialogResult]::No) {
      Start-WinTunerPackageBuild -PackageId $Context.PackageId -PackageFolder $Context.PackageFolder -DesiredVersion $Context.DesiredVersion -LatestVersion $Context.LatestVersion -Mode Previous
      return
    }

    Update-Status 'Package creation canceled.'
    Write-Log "Package creation canceled after a hash mismatch for '$($Context.PackageId)'."
    Update-PackageActionState
    Update-UpdateActionState
    Update-DiscoveryActionState
    Update-SupersededActionState
    Update-PackageSearchActionState
    return
  }

  $script:isPackageBuildActive = $false
  if (-not [bool]$buildResult.Succeeded) {
    $message = if ([string]::IsNullOrWhiteSpace([string]$buildResult.ErrorMessage)) { 'Unknown package creation error.' } else { [string]$buildResult.ErrorMessage }
    Update-Status "Package creation failed: $message"
    Write-Log "Package creation failed for '$($Context.PackageId)': $message"
    Update-PackageActionState
    Update-UpdateActionState
    Update-DiscoveryActionState
    Update-SupersededActionState
    Update-PackageSearchActionState
    return
  }

  $effectiveVersion = [string]$buildResult.EffectiveVersion
  if ([string]::IsNullOrWhiteSpace($effectiveVersion)) {
    $effectiveVersion = [string]$Context.LatestVersion
  }
  $createdArtifact = Test-WinTunerPackageArtifact `
    -RootPackageFolder $Context.PackageFolder `
    -PackageId $Context.PackageId `
    -Version $effectiveVersion

  if ($createdArtifact.IsValid) {
    $script:builtVersions[$Context.PackageId] = $effectiveVersion
    Update-Status ("Package created successfully (version {0})" -f $effectiveVersion)
    Write-Log "Background package build validated for '$($Context.PackageId)' version $effectiveVersion -> $($createdArtifact.IntuneWinPath)"
  } else {
    $script:builtVersions.Remove($Context.PackageId)
    Update-Status 'Package creation completed, but artifact validation failed.'
    Write-Log "Package build validation failed for '$($Context.PackageId)' version $effectiveVersion ($($createdArtifact.ReasonCode)): $($createdArtifact.Reason)"
  }

  Update-PackageActionState
  Update-UpdateActionState
  Update-DiscoveryActionState
  Update-SupersededActionState
  Update-PackageSearchActionState
}

function Start-WinTunerPackageBuild {
  param(
    [Parameter(Mandatory=$true)][string]$PackageId,
    [Parameter(Mandatory=$true)][string]$PackageFolder,
    [AllowNull()][string]$DesiredVersion,
    [AllowNull()][string]$LatestVersion,
    [ValidateSet('Automatic', 'RetrySame', 'Previous')][string]$Mode = 'Automatic'
  )

  if (
    $script:isPackageSearchActive -or
    $script:isVersionLookupActive -or
    $script:isPackageBuildActive -or
    $script:isPackageUploadActive -or
    $script:isUpdateOperationActive -or
    $script:discoveryScanRunning -or
    $script:isDiscoveryDeploymentActive -or
    $script:isSupersededOperationActive
  ) {
    Update-PackageSearchActionState
    return
  }

  $buildModulePath = Join-Path $PSScriptRoot 'Modules\WinTuner.PackageBuild.psm1'
  if (-not (Test-Path -LiteralPath $buildModulePath -PathType Leaf)) {
    Update-Status 'Package creation failed: package build module is missing.'
    Write-Log "Package build module missing: $buildModulePath"
    return
  }

  $script:isPackageBuildActive = $true
  $script:progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
  $script:progressBar.MarqueeAnimationSpeed = 30
  $script:progressBar.Visible = $true
  $operationText = if ($Mode -eq 'RetrySame') { 'Retrying package creation' } elseif ($Mode -eq 'Previous') { 'Creating previous package version' } else { 'Creating package' }
  Update-Status "$operationText for $PackageId..."
  Write-Log "Starting background package build for '$PackageId' (mode: $Mode)."
  Update-UpdateActionState
  Update-DiscoveryActionState
  Update-SupersededActionState
  Update-PackageSearchActionState

  $buildScript = @(
    'param($RepositoryRoot, $PackageId, $PackageFolder, $DesiredVersion, $LatestVersion, $Mode)'
    '$ErrorActionPreference = ''Stop'''
    '$ProgressPreference = ''SilentlyContinue'''
    'Import-Module WinTuner -ErrorAction Stop'
    'Import-Module (Join-Path $RepositoryRoot ''Modules\WinTuner.Core.psm1'') -Force -ErrorAction Stop'
    'Import-Module (Join-Path $RepositoryRoot ''Modules\WinTuner.Winget.psm1'') -Force -ErrorAction Stop'
    'Import-Module (Join-Path $RepositoryRoot ''Modules\WinTuner.PackageBuild.psm1'') -Force -ErrorAction Stop'
    '$result = Invoke-WinTunerPackageBuild -PackageId $PackageId -PackageFolder $PackageFolder -DesiredVersion $DesiredVersion -LatestVersion $LatestVersion -Mode $Mode'
    '$result | ConvertTo-Json -Depth 4 -Compress'
  ) -join [Environment]::NewLine

  $powerShell = [System.Management.Automation.PowerShell]::Create()
  $null = $powerShell.AddScript($buildScript).AddArgument($PSScriptRoot).AddArgument($PackageId).AddArgument($PackageFolder).AddArgument($DesiredVersion).AddArgument($LatestVersion).AddArgument($Mode)
  $timer = New-Object System.Windows.Forms.Timer
  $timer.Interval = 150
  $context = [pscustomobject]@{
    PowerShell     = $powerShell
    AsyncResult    = $null
    Timer          = $timer
    PackageId      = $PackageId
    PackageFolder  = $PackageFolder
    DesiredVersion = $DesiredVersion
    LatestVersion  = $LatestVersion
    Mode           = $Mode
  }
  $script:packageBuildContext = $context

  $timer.Add_Tick({
    $currentContext = $script:packageBuildContext
    if ($currentContext -and $currentContext.AsyncResult -and $currentContext.AsyncResult.IsCompleted) {
      Complete-WinTunerPackageBuild -Context $currentContext
    }
  })

  try {
    $context.AsyncResult = $powerShell.BeginInvoke()
    $timer.Start()
  } catch {
    try { $timer.Dispose() } catch {}
    try { $powerShell.Dispose() } catch {}
    $script:packageBuildContext = $null
    $script:isPackageBuildActive = $false
    $script:progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $script:progressBar.Value = 0
    $script:progressBar.Visible = $false
    Update-Status "Package creation failed to start: $($_.Exception.Message)"
    Write-Log "Background package build failed to start for '$PackageId': $($_.Exception.Message)"
    Update-PackageActionState
    Update-UpdateActionState
    Update-DiscoveryActionState
    Update-SupersededActionState
    Update-PackageSearchActionState
  }
}
function Complete-WinTunerPackageUpload {
  param([Parameter(Mandatory=$true)][object]$Context)

  $uploadResult = $null
  $completionError = $null
  try {
    $output = @($Context.PowerShell.EndInvoke($Context.AsyncResult))
    if ($output.Count -eq 0) {
      if ($Context.PowerShell.Streams.Error.Count -gt 0) {
        throw $Context.PowerShell.Streams.Error[0].Exception
      }
      throw 'The package upload returned no result.'
    }

    $json = [string]$output[$output.Count - 1]
    $uploadResult = $json | ConvertFrom-Json -ErrorAction Stop
  } catch {
    $completionError = $_.Exception.Message
  } finally {
    try { $Context.Timer.Stop() } catch {}
    try { $Context.Timer.Dispose() } catch {}
    try { $Context.PowerShell.Dispose() } catch {}
    if ($script:packageUploadContext -eq $Context) {
      $script:packageUploadContext = $null
    }
    try {
      $script:progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
      $script:progressBar.Value = 0
      $script:progressBar.Visible = $false
    } catch {}
  }

  $script:isPackageUploadActive = $false
  if ($completionError -or -not [bool]$uploadResult.Succeeded) {
    $reasonCode = if ($completionError) { 'RunspaceError' } elseif ([string]::IsNullOrWhiteSpace([string]$uploadResult.ReasonCode)) { 'UploadFailed' } else { [string]$uploadResult.ReasonCode }
    $errorMessage = if ($completionError) { $completionError } elseif ([string]::IsNullOrWhiteSpace([string]$uploadResult.ErrorMessage)) { 'Unknown upload error.' } else { [string]$uploadResult.ErrorMessage }

    if ($reasonCode -ne 'DeploymentFailed' -and $script:builtVersions.ContainsKey($Context.PackageId)) {
      $script:builtVersions.Remove($Context.PackageId)
    }

    Update-Status 'Upload failed: See log for details'
    Write-Log "Background upload failed for '$($Context.PackageId)' version $($Context.Version) ($reasonCode): $errorMessage"
    Update-PackageActionState
    Update-UpdateActionState
    Update-DiscoveryActionState
    Update-SupersededActionState
    Update-PackageSearchActionState

    $errorDetails = "Upload of $($Context.PackageId) (v$($Context.Version)) failed.`n`n"
    $errorDetails += "Error: $errorMessage`n`n"
    $errorDetails += "Possible solutions:`n"
    $errorDetails += "1. Check if the app already exists in Intune (delete and retry)`n"
    $errorDetails += "2. Update WinTuner module: Update-Module WinTuner`n"
    $errorDetails += "3. Try a different app to test`n"
    $errorDetails += "4. Check Intune service health`n"
    $errorDetails += "`nActivity logged to WinTuner_GUI.log"
    [void][System.Windows.Forms.MessageBox]::Show(
      $errorDetails,
      'Upload Failed',
      [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Error
    )
    return
  }

  Update-Status 'Upload completed successfully'
  Write-Log "Background upload completed for '$($Context.PackageId)' version $($Context.Version) -> $($uploadResult.IntuneWinPath)"
  $appSearchBox.Text = ''
  $dropdown.Items.Clear()
  $script:packageMap.Clear()
  if ($script:selectedPackageVersions.ContainsKey($Context.PackageId)) {
    $script:selectedPackageVersions.Remove($Context.PackageId)
    Write-Log "Cleared cached version for $($Context.PackageId) after upload"
  }

  Update-PackageActionState
  Update-UpdateActionState
  Update-DiscoveryActionState
  Update-SupersededActionState
  Update-PackageSearchActionState
}

function Start-WinTunerPackageUpload {
  param(
    [Parameter(Mandatory=$true)][string]$PackageId,
    [Parameter(Mandatory=$true)][string]$Version,
    [Parameter(Mandatory=$true)][string]$RootPackageFolder
  )

  if (
    -not $script:isConnected -or
    $script:isPackageSearchActive -or
    $script:isVersionLookupActive -or
    $script:isPackageBuildActive -or
    $script:isPackageUploadActive -or
    $script:isUpdateOperationActive -or
    $script:discoveryScanRunning -or
    $script:isDiscoveryDeploymentActive -or
    $script:isSupersededOperationActive
  ) {
    Update-PackageActionState
    return
  }

  $artifactValidation = Test-WinTunerPackageArtifact `
    -RootPackageFolder $RootPackageFolder `
    -PackageId $PackageId `
    -Version $Version
  if (-not $artifactValidation.IsValid) {
    Update-Status "Cannot upload: $($artifactValidation.Reason)"
    Write-Log "Upload blocked for $PackageId version $Version ($($artifactValidation.ReasonCode)): $($artifactValidation.Reason)"
    if (
      $script:builtVersions.ContainsKey($PackageId) -and
      ([string]$script:builtVersions[$PackageId] -eq $Version)
    ) {
      $script:builtVersions.Remove($PackageId)
    }
    Update-PackageActionState
    return
  }

  $uploadModulePath = Join-Path $PSScriptRoot 'Modules\WinTuner.PackageUpload.psm1'
  if (-not (Test-Path -LiteralPath $uploadModulePath -PathType Leaf)) {
    Update-Status 'Upload failed: package upload module is missing.'
    Write-Log "Package upload module missing: $uploadModulePath"
    return
  }

  $script:isPackageUploadActive = $true
  $script:progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
  $script:progressBar.MarqueeAnimationSpeed = 30
  $script:progressBar.Visible = $true
  Update-Status "Uploading $PackageId (v$Version) to tenant..."
  Write-Log "Starting background upload for '$PackageId' version $Version after click-time validation -> $($artifactValidation.IntuneWinPath)"
  Update-PackageActionState
  Update-UpdateActionState
  Update-DiscoveryActionState
  Update-SupersededActionState
  Update-PackageSearchActionState

  $uploadScript = @(
    'param($RepositoryRoot, $PackageId, $Version, $RootPackageFolder)'
    '$ErrorActionPreference = ''Stop'''
    '$ProgressPreference = ''SilentlyContinue'''
    'Import-Module WinTuner -ErrorAction Stop'
    'Import-Module (Join-Path $RepositoryRoot ''Modules\WinTuner.Core.psm1'') -Force -ErrorAction Stop'
    'Import-Module (Join-Path $RepositoryRoot ''Modules\WinTuner.PackageUpload.psm1'') -Force -ErrorAction Stop'
    '$result = Invoke-WinTunerPackageUpload -PackageId $PackageId -Version $Version -RootPackageFolder $RootPackageFolder'
    '$result | ConvertTo-Json -Depth 4 -Compress'
  ) -join [Environment]::NewLine

  $powerShell = [System.Management.Automation.PowerShell]::Create()
  $null = $powerShell.AddScript($uploadScript).AddArgument($PSScriptRoot).AddArgument($PackageId).AddArgument($Version).AddArgument($RootPackageFolder)
  $timer = New-Object System.Windows.Forms.Timer
  $timer.Interval = 150
  $context = [pscustomobject]@{
    PowerShell       = $powerShell
    AsyncResult      = $null
    Timer            = $timer
    PackageId        = $PackageId
    Version          = $Version
    RootPackageFolder = $RootPackageFolder
  }
  $script:packageUploadContext = $context

  $timer.Add_Tick({
    $currentContext = $script:packageUploadContext
    if ($currentContext -and $currentContext.AsyncResult -and $currentContext.AsyncResult.IsCompleted) {
      Complete-WinTunerPackageUpload -Context $currentContext
    }
  })

  try {
    $context.AsyncResult = $powerShell.BeginInvoke()
    $timer.Start()
  } catch {
    try { $timer.Dispose() } catch {}
    try { $powerShell.Dispose() } catch {}
    $script:packageUploadContext = $null
    $script:isPackageUploadActive = $false
    $script:progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $script:progressBar.Value = 0
    $script:progressBar.Visible = $false
    Update-Status "Upload failed to start: $($_.Exception.Message)"
    Write-Log "Background upload failed to start for '$PackageId': $($_.Exception.Message)"
    Update-PackageActionState
    Update-UpdateActionState
    Update-DiscoveryActionState
    Update-SupersededActionState
    Update-PackageSearchActionState
  }
}
# Helper: validate M365 username (UPN-like)
function Test-ValidM365UserName {
  param([string]$UserName)
  if ([string]::IsNullOrWhiteSpace($UserName)) { return $false }
  # Balanced, pragmatic UPN check
  $upnRegex = '^(?=.{3,256}$)(?![.])(?!.*[.]{2})[A-Za-z0-9._%+\-]+@(?:[A-Za-z0-9-]+\.)+[A-Za-z]{2,}$'
  return ($UserName -match $upnRegex)
}

# Adds a UPN to the recent users list (only if RememberMe is active)
function Add-RecentUser {
  param([string]$Upn)
  if (-not $script:settings.RememberMe) { return }
  if ([string]::IsNullOrWhiteSpace($Upn)) { return }
  $max = if ($script:settings.MaxRecentUsers -gt 0) { $script:settings.MaxRecentUsers } else { 3 }
  $list = [System.Collections.Generic.List[string]]::new()
  foreach ($u in @($script:settings.RecentUsers)) {
    if ($u -and $u -ne $Upn) { $list.Add($u) }
  }
  $list.Insert(0, $Upn)
  while ($list.Count -gt $max) { $list.RemoveAt($list.Count - 1) }
  $script:settings.RecentUsers = $list.ToArray()
  $script:settings.LastUser = $Upn
  [void](Export-WinTunerSettings -Settings $script:settings -Path $script:settingsPath)
}

# Clears the recent users list and resets LastUser
function Clear-RecentUsers {
  $script:settings.RecentUsers = @()
  $script:settings.LastUser = ""
  [void](Export-WinTunerSettings -Settings $script:settings -Path $script:settingsPath)
}

# Helper: check if WinTuner is connected (simple smoke test)
function Test-WtConnected {
  param(
    [int]$MaxAttempts = 4,
    [int]$RetryDelayMs = 500
  )

  for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
    try {
      # Avoid Select-Object -First 1 to prevent WinForms pipeline crash during login
      $apps = Get-WtWin32Apps -Update:$false -Superseded:$false -ErrorAction Stop

      foreach ($app in $apps) {
        return $true # Exit safely on first found element
      }

      return $true # No apps found but no error either
    } catch {
      if ($attempt -ge $MaxAttempts) {
        Write-Log "Connection verification failed after $MaxAttempts attempts: $($_.Exception.Message)"
        return $false
      }

      Write-Log "Connection verification attempt $attempt failed; retrying in ${RetryDelayMs}ms..."
      Start-Sleep -Milliseconds $RetryDelayMs
      [System.Windows.Forms.Application]::DoEvents()
    }
  }

  return $false
}

# Heuristic filter to avoid very slow/low-value WinGet queries (mainly mobile/system artifacts)
function Test-WingetSearchCandidate {
  param(
    [string]$DisplayName
  )

  if ([string]::IsNullOrWhiteSpace($DisplayName)) { return $false }
  $name = $DisplayName.Trim()
  if ($name.Length -lt 3) { return $false }

  # Android-style package ids and similar technical identifiers are typically not useful for WinGet search
  if ($name -match '^[a-z0-9]+(\.[a-z0-9_]+){2,}$') { return $false }
  if ($name -match '(?i)^com\.') { return $false }

  # Skip common mobile/system terms that frequently stall searches and rarely map to WinGet packages
  if ($name -match '(?i)\b(apn|provisioner|sim toolkit|sim card|carrier services|system ui|one ui home|setup wizard)\b') { return $false }

  return $true
}

# Helper: toggle UI based on connection state
function Set-ConnectedUIState {
  param([bool]$Connected)
  if ($Connected) {
    $loginButton.Visible = $false
    $usernameBox.Visible = $false
    $usernameLabel.Visible = $false
    if ($usernameError) { $usernameError.Visible = $false }
    $tabControl.Visible = $true
    $logoutButton.Visible = $true
    if ($clearHistoryButton) { $clearHistoryButton.Visible = $false }
  } else {
    $loginButton.Visible = $true
    $usernameBox.Visible = $true
    $usernameLabel.Visible = $true
    if ($usernameError) { $usernameError.Visible = $true }
    $tabControl.Visible = $true
    $logoutButton.Visible = $false
    if ($clearHistoryButton) { $clearHistoryButton.Visible = $true }
  }
  if ($rememberCheckBox) { $rememberCheckBox.Visible = -not $Connected }
  if (-not $Connected) {
    $script:updateApps = [System.Collections.Generic.List[object]]::new()
    $script:updateVisibleApps = [System.Collections.Generic.List[object]]::new()
    if ($updateListBox) { $updateListBox.Items.Clear() }
    if ($updateFilterBox) { $updateFilterBox.Text = '' }
    Clear-DiscoveryCandidateState
    $script:supersededApps = @()
    if ($supersededDropdown) { $supersededDropdown.Items.Clear() }
  }
  Update-UpdateActionState
  Update-DiscoveryActionState
  Update-SupersededActionState
  
  if ($loginInfoLabel) {
    $loginInfoLabel.Visible = $Connected
    if ($Connected -and $script:currentUserUpn) { $loginInfoLabel.Text = "Logged in as: $($script:currentUserUpn)" }
  }
}

$script:isConnected = $false
$script:currentUserUpn = ""

# Cache effective builds and package versions validated from disk
$script:builtVersions = @{}
$script:updateApps = [System.Collections.Generic.List[object]]::new()
$script:updateVisibleApps = [System.Collections.Generic.List[object]]::new()
$script:isUpdateOperationActive = $false
$script:updateScanRunning = $false
$script:cancelUpdateScan = $false
$script:updateScanContext = $null
$script:updateDeploymentContext = $null
$script:supersededApps = @()
$script:isSupersededOperationActive = $false
$script:supersededSearchContext = $null
# Cache for winget version lookups (speeds up repeated searches)
# Disk cache loaded once at first use (Fix 1)
# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = "WinTuner GUI"
$form.Size = New-Object System.Drawing.Size(960, 850)
$form.MinimumSize = New-Object System.Drawing.Size(800, 700)
$form.Padding = '5,5,5,5'
$contentWidth = [Math]::Max(760, $form.ClientSize.Width - 20)

# Header panel – contains all login/top controls so they stay in one row
$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Dock = [System.Windows.Forms.DockStyle]::Top
$headerPanel.Height = 78
$form.Controls.Add($headerPanel)

# Theme toggle button (top right, anchored so it never clips)
$themeToggleButton = New-Object System.Windows.Forms.Button
$themeToggleButton.Text = "Light Mode"  # indicates action from dark -> light
$themeToggleButton.Location = New-Object System.Drawing.Point(835, 8)
$themeToggleButton.Size = New-Object System.Drawing.Size(100, 27)
$themeToggleButton.Add_Click({ Switch-GuiTheme })
$themeToggleButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$headerPanel.Controls.Add($themeToggleButton)

# Username label and textbox
$usernameLabel = New-Object System.Windows.Forms.Label
$usernameLabel.Text = "Username:"
$usernameLabel.Location = New-Object System.Drawing.Point(10, 10)
$usernameLabel.AutoSize = $true
$headerPanel.Controls.Add($usernameLabel)

$usernameBox = New-Object System.Windows.Forms.ComboBox
$usernameBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
# ENTER in Username -> "Login"
if ($usernameBox -ne $null) {
  $usernameBox.Add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
      if ($loginButton -and $loginButton.Enabled) {
        $loginButton.PerformClick()
      } else {
        [void][System.Windows.Forms.MessageBox]::Show(
          "Please enter a valid M365 UPN, e.g. name@firma.de",
          "Validation",
          [System.Windows.Forms.MessageBoxButtons]::OK,
          [System.Windows.Forms.MessageBoxIcon]::Information
        )
      }
      $e.SuppressKeyPress = $true
    }
  })
}
$usernameBox.Location = New-Object System.Drawing.Point(88, 10)
$usernameBox.Width = 365
$usernameBox.Height = 27
$headerPanel.Controls.Add($usernameBox)

# "Clear history" button next to username ComboBox
$clearHistoryButton = New-Object System.Windows.Forms.Button
$clearHistoryButton.Text = ([System.Char]::ConvertFromUtf32(0x1F5D1)) + " Clear History"
$clearHistoryButton.Width = 115
$clearHistoryButton.Height = 27
$clearHistoryButton.Location = New-Object System.Drawing.Point(461, 10)
$clearHistoryButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$headerPanel.Controls.Add($clearHistoryButton)

$clearHistoryButton.Add_Click({
  Clear-RecentUsers
  $usernameBox.Items.Clear()
  $usernameBox.Text = ""
  Update-Status "Username history cleared."
  Write-Log "Username history cleared by user."
})

# Validation hint label for username
$usernameError = New-Object System.Windows.Forms.Label
$usernameError.Text = ""
$usernameError.Location = New-Object System.Drawing.Point(88, 46)
$usernameError.AutoSize = $true
$usernameError.ForeColor = [System.Drawing.Color]::FromArgb(220,80,80)
$headerPanel.Controls.Add($usernameError)

# Live validation for username field
$usernameBox.add_TextChanged({
  if (Test-ValidM365UserName -UserName $usernameBox.Text) {
    $usernameError.Text = ""
    if ($loginButton) { $loginButton.Enabled = $true }
  } else {
    $usernameError.Text = "Please enter a valid M365 UPN, e.g. name@firma.de"
    if ($loginButton) { $loginButton.Enabled = $false }
  }
})

# Status label
$script:statusLabel = New-Object System.Windows.Forms.Label
$script:statusLabel.Text = ""
$script:statusLabel.Location = New-Object System.Drawing.Point(10, 745)
$script:statusLabel.Width = $contentWidth
$script:statusLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($script:statusLabel)

# Output textbox (Log area below tabs and progress bar)
$script:outputBox = New-Object System.Windows.Forms.TextBox
$script:outputBox.Location = New-Object System.Drawing.Point(10, 620)
$script:outputBox.Size = New-Object System.Drawing.Size($contentWidth, 120)
$script:outputBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$script:outputBox.Multiline = $true
$script:outputBox.ScrollBars = "Vertical"
$script:outputBox.ReadOnly = $true
$form.Controls.Add($script:outputBox)

# Attach GUI log output to the logging module
Initialize-WinTunerLogging -BasePath $PSScriptRoot -OutputBox $script:outputBox

# Progress bar (appears between tabs and log when active)
$script:progressBar = New-Object System.Windows.Forms.ProgressBar
$script:progressBar.Location = New-Object System.Drawing.Point(10, 595)
$script:progressBar.Width = $contentWidth
$script:progressBar.Height = 20
$script:progressBar.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$script:progressBar.Visible = $false
$form.Controls.Add($script:progressBar)

# Logout button
$logoutButton = New-Object System.Windows.Forms.Button
$logoutButton.Text = "Tenant Logout"
$logoutButton.Location = New-Object System.Drawing.Point(584, 10)
$logoutButton.Size = New-Object System.Drawing.Size(150, 27)
$logoutButton.Visible = $false
$headerPanel.Controls.Add($logoutButton)

$loginInfoLabel = New-Object System.Windows.Forms.Label
$loginInfoLabel.Text = ""
$loginInfoLabel.Location = New-Object System.Drawing.Point(88, 10)
$loginInfoLabel.AutoSize = $true
$loginInfoLabel.Visible = $false
$headerPanel.Controls.Add($loginInfoLabel)

# TabControl
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(10, 88)
$tabControl.Size = New-Object System.Drawing.Size($contentWidth, 500)
$tabControl.Visible = $true
$tabControl.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$form.Controls.Add($tabControl)

# Tab: WinGet Apps
$tabCreate = New-Object System.Windows.Forms.TabPage
$tabCreate.Text = "WinGet Apps"
$tabControl.TabPages.Add($tabCreate)

$appSearchLabel = New-Object System.Windows.Forms.Label
$appSearchLabel.Text = "App search:"
$appSearchLabel.Location = New-Object System.Drawing.Point(10,20)
$appSearchLabel.AutoSize = $true
$tabCreate.Controls.Add($appSearchLabel)

$appSearchBox = New-Object System.Windows.Forms.TextBox
# ENTER in App search -> "Search"
if ($appSearchBox -ne $null) {
  $appSearchBox.Add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
      if ($searchButton -and $searchButton.Enabled) {
        $searchButton.PerformClick()
      }
      $e.SuppressKeyPress = $true
    }
  })
}
$appSearchBox.Location = New-Object System.Drawing.Point(100,20)
$appSearchBox.Width = 450
$appSearchBox.BorderStyle = 'FixedSingle'
$tabCreate.Controls.Add($appSearchBox)

$searchButton = New-Object System.Windows.Forms.Button
$searchButton.Text = "Search"
$searchButton.Location = New-Object System.Drawing.Point(570,20)
$searchButton.Width = 180
$tabCreate.Controls.Add($searchButton)

$dropdown = New-Object System.Windows.Forms.ComboBox
$dropdown.Location = New-Object System.Drawing.Point(100,60)
$dropdown.Width = 450
$dropdown.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$dropdown.Enabled = $false
$tabCreate.Controls.Add($dropdown)

$versionsButton = New-Object System.Windows.Forms.Button
$versionsButton.Text = "Versions..."
$versionsButton.Location = New-Object System.Drawing.Point(570,60)
$versionsButton.Width = 180
$versionsButton.Enabled = $false
$tabCreate.Controls.Add($versionsButton)

$pathLabel = New-Object System.Windows.Forms.Label
$pathLabel.Text = "File path:"
$pathLabel.Location = New-Object System.Drawing.Point(10,100)
$pathLabel.AutoSize = $true
$tabCreate.Controls.Add($pathLabel)

$pathBox = New-Object System.Windows.Forms.TextBox
$pathBox.Location = New-Object System.Drawing.Point(100,100)
$pathBox.Width = 450
$pathBox.BorderStyle = 'FixedSingle'
$pathBox.Text = "C:\Temp"
$tabCreate.Controls.Add($pathBox)

$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = "Select..."
$browseButton.Location = New-Object System.Drawing.Point(570,100)
$browseButton.Width = 180
$tabCreate.Controls.Add($browseButton)

$browseButton.Add_Click({
  $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
  if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $pathBox.Text = $folderBrowser.SelectedPath
  }
})

$createButton = New-Object System.Windows.Forms.Button
$createButton.Text = "Create package"
$createButton.Location = New-Object System.Drawing.Point(100,140)
$createButton.Width = 180
$createButton.Enabled = $false
$tabCreate.Controls.Add($createButton)

$uploadButton = New-Object System.Windows.Forms.Button
$uploadButton.Text = "Upload to Tenant"
$uploadButton.Location = New-Object System.Drawing.Point(290,140)
$uploadButton.Width = 180
$uploadButton.Visible = $true
$uploadButton.Enabled = $false
$tabCreate.Controls.Add($uploadButton)

# Recalculate package readiness whenever the selected app or package root changes.
$dropdown.Add_SelectedIndexChanged({
  Update-PackageActionState
  Update-PackageSearchActionState
})
$pathBox.Add_TextChanged({
  Update-PackageActionState
})

# Tab: Updates
$tabUpdate = New-Object System.Windows.Forms.TabPage
$tabUpdate.Text = "Updates"
$tabControl.TabPages.Add($tabUpdate)

# Label over "Search Updates"
$updateHeaderLabel = New-Object System.Windows.Forms.Label
$updateHeaderLabel.Text = "Update existing apps"
$updateHeaderLabel.Location = New-Object System.Drawing.Point(100,20)
$updateHeaderLabel.AutoSize = $true
$updateHeaderLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$tabUpdate.Controls.Add($updateHeaderLabel)

$updateSearchButton = New-Object System.Windows.Forms.Button
$updateSearchButton.Text = "Search Updates"
$updateSearchButton.Location = New-Object System.Drawing.Point(100,50)
$updateSearchButton.Width = 180
$updateSearchButton.Enabled = $false
$tabUpdate.Controls.Add($updateSearchButton)

# Filter TextBox for Update List (on same row as Search button)
$updateFilterLabel = New-Object System.Windows.Forms.Label
$updateFilterLabel.Text = "Filter:"
$updateFilterLabel.Location = New-Object System.Drawing.Point(300,53)
$updateFilterLabel.AutoSize = $true
$tabUpdate.Controls.Add($updateFilterLabel)

$updateFilterBox = New-Object System.Windows.Forms.TextBox
$updateFilterBox.Location = New-Object System.Drawing.Point(355,50)
$updateFilterBox.Width = [Math]::Max(395, $tabUpdate.ClientSize.Width - 365)
$updateFilterBox.PlaceholderText = "Type to filter apps..."
$updateFilterBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$tabUpdate.Controls.Add($updateFilterBox)

# CheckedListBox for multi-select updates
$updateListBox = New-Object System.Windows.Forms.CheckedListBox
$updateListBox.Location = New-Object System.Drawing.Point(100,85)
$updateListBox.Width = [Math]::Max(650, $tabUpdate.ClientSize.Width - 120)
$updateListBox.Height = 130
$updateListBox.CheckOnClick = $true
$updateListBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$tabUpdate.Controls.Add($updateListBox)

$updateSummaryLabel = New-Object System.Windows.Forms.Label
$updateSummaryLabel.Text = "Candidates: 0 | Checked: 0"
$updateSummaryLabel.Location = New-Object System.Drawing.Point(100,218)
$updateSummaryLabel.AutoSize = $true
$tabUpdate.Controls.Add($updateSummaryLabel)

# Helper buttons for check/uncheck all
$checkAllButton = New-Object System.Windows.Forms.Button
$checkAllButton.Text = "✓ Check All"
$checkAllButton.Location = New-Object System.Drawing.Point(100,245)
$checkAllButton.Width = 120
$checkAllButton.Enabled = $false
$tabUpdate.Controls.Add($checkAllButton)

$uncheckAllButton = New-Object System.Windows.Forms.Button
$uncheckAllButton.Text = "☐ Uncheck All"
$uncheckAllButton.Location = New-Object System.Drawing.Point(230,245)
$uncheckAllButton.Width = 120
$uncheckAllButton.Enabled = $false
$tabUpdate.Controls.Add($uncheckAllButton)

$updateSelectedButton = New-Object System.Windows.Forms.Button
$updateSelectedButton.Text = "Update checked apps"
$updateSelectedButton.Location = New-Object System.Drawing.Point(370,245)
$updateSelectedButton.Width = 200
$updateSelectedButton.Enabled = $false
$tabUpdate.Controls.Add($updateSelectedButton)

$updateAllButton = New-Object System.Windows.Forms.Button
$updateAllButton.Text = "Update ALL (unchecked too)"
$updateAllButton.Location = New-Object System.Drawing.Point(580,245)
$updateAllButton.Width = 170
$updateAllButton.Enabled = $false
$tabUpdate.Controls.Add($updateAllButton)

# Label over "Search Superseded Apps"
$supersededHeaderLabel = New-Object System.Windows.Forms.Label
$supersededHeaderLabel.Text = "Search for superseded Apps"
$supersededHeaderLabel.Location = New-Object System.Drawing.Point(100,275)
$supersededHeaderLabel.AutoSize = $true
$supersededHeaderLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$tabUpdate.Controls.Add($supersededHeaderLabel)

$supersededSearchButton = New-Object System.Windows.Forms.Button
$supersededSearchButton.Text = "Search Superseded Apps"
$supersededSearchButton.Location = New-Object System.Drawing.Point(100,305)
$supersededSearchButton.Width = 250
$supersededSearchButton.Enabled = $false
$tabUpdate.Controls.Add($supersededSearchButton)

# Dropdown: Superseded Apps (Name + Version)
$supersededDropdown = New-Object System.Windows.Forms.ComboBox
$supersededDropdown.Location = New-Object System.Drawing.Point(100,345)
$supersededDropdown.Width = 650
$supersededDropdown.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$tabUpdate.Controls.Add($supersededDropdown)

# Button: Delete selected app
$deleteSelectedAppButton = New-Object System.Windows.Forms.Button
$deleteSelectedAppButton.Text = "Delete Selected App"
$deleteSelectedAppButton.Location = New-Object System.Drawing.Point(100,385)
$deleteSelectedAppButton.Width = 250
$deleteSelectedAppButton.Enabled = $false
$tabUpdate.Controls.Add($deleteSelectedAppButton)

$removeOldAppsButton = New-Object System.Windows.Forms.Button
$removeOldAppsButton.Text = "Delete all Superseded Apps"
$removeOldAppsButton.Location = New-Object System.Drawing.Point(360,385)
$removeOldAppsButton.Width = 250
$removeOldAppsButton.Enabled = $false
$tabUpdate.Controls.Add($removeOldAppsButton)

# ==================================================
# Tab: Discovered Apps
# ==================================================
$tabDiscovered = New-Object System.Windows.Forms.TabPage
$tabDiscovered.Text = "Discovered Apps"
$tabControl.TabPages.Add($tabDiscovered)

$discoveredHeaderLabel = New-Object System.Windows.Forms.Label
$discoveredHeaderLabel.Text = "Discovered Apps in Intune"
$discoveredHeaderLabel.Location = New-Object System.Drawing.Point(20,20)
$discoveredHeaderLabel.AutoSize = $true
$discoveredHeaderLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$tabDiscovered.Controls.Add($discoveredHeaderLabel)

$lastDiscoveryLabel = New-Object System.Windows.Forms.Label
$lastDiscoveryLabel.Text = "Last discovery: Never"
$lastDiscoveryLabel.Location = New-Object System.Drawing.Point(250,24)
$lastDiscoveryLabel.AutoSize = $true
$tabDiscovered.Controls.Add($lastDiscoveryLabel)

$scanDiscoveredButton = New-Object System.Windows.Forms.Button
$scanDiscoveredButton.Text = "1. Scan Discovered Apps"
$scanDiscoveredButton.Location = New-Object System.Drawing.Point(20,50)
$scanDiscoveredButton.Width = 200
$scanDiscoveredButton.Enabled = $false
$tabDiscovered.Controls.Add($scanDiscoveredButton)

$deployDiscoveredButton = New-Object System.Windows.Forms.Button
$deployDiscoveredButton.Text = "2. Deploy Checked Apps"
$deployDiscoveredButton.Location = New-Object System.Drawing.Point(230,50)
$deployDiscoveredButton.Width = 200
$deployDiscoveredButton.Enabled = $false
$tabDiscovered.Controls.Add($deployDiscoveredButton)

$exportDiscoveredCsvButton = New-Object System.Windows.Forms.Button
$exportDiscoveredCsvButton.Text = "3. Export Winget IDs CSV"
$exportDiscoveredCsvButton.Location = New-Object System.Drawing.Point(250,74)
$exportDiscoveredCsvButton.Width = 180
$exportDiscoveredCsvButton.Enabled = $false
$tabDiscovered.Controls.Add($exportDiscoveredCsvButton)

$checkAllDiscoveredButton = New-Object System.Windows.Forms.Button
$checkAllDiscoveredButton.Text = "☑ Check All"
$checkAllDiscoveredButton.Location = New-Object System.Drawing.Point(20,74)
$checkAllDiscoveredButton.Width = 100
$checkAllDiscoveredButton.Enabled = $false
$tabDiscovered.Controls.Add($checkAllDiscoveredButton)

$uncheckAllDiscoveredButton = New-Object System.Windows.Forms.Button
$uncheckAllDiscoveredButton.Text = "☐ Uncheck All"
$uncheckAllDiscoveredButton.Location = New-Object System.Drawing.Point(130,74)
$uncheckAllDiscoveredButton.Width = 110
$uncheckAllDiscoveredButton.Enabled = $false
$tabDiscovered.Controls.Add($uncheckAllDiscoveredButton)

# --- NEU: Filter & Sortierung ---
$discoveryFilterLabelX = [Math]::Max(440, $tabDiscovered.ClientSize.Width - 310)
$discoveryFilterControlX = [Math]::Max(540, $tabDiscovered.ClientSize.Width - 210)

$discoveredAppSearchLabel = New-Object System.Windows.Forms.Label
$discoveredAppSearchLabel.Text = "Search App:"
$discoveredAppSearchLabel.Location = New-Object System.Drawing.Point($discoveryFilterLabelX, 15)
$discoveredAppSearchLabel.AutoSize = $true
$discoveredAppSearchLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$tabDiscovered.Controls.Add($discoveredAppSearchLabel)

$discoveredAppSearchBox = New-Object System.Windows.Forms.TextBox
$discoveredAppSearchBox.Location = New-Object System.Drawing.Point($discoveryFilterControlX, 12)
$discoveredAppSearchBox.Width = 150
$discoveredAppSearchBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$tabDiscovered.Controls.Add($discoveredAppSearchBox)

$discoveredPublisherLabel = New-Object System.Windows.Forms.Label
$discoveredPublisherLabel.Text = "Publisher:"
$discoveredPublisherLabel.Location = New-Object System.Drawing.Point($discoveryFilterLabelX, 42)
$discoveredPublisherLabel.AutoSize = $true
$discoveredPublisherLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$tabDiscovered.Controls.Add($discoveredPublisherLabel)

$discoveredPublisherBox = New-Object System.Windows.Forms.ComboBox
$discoveredPublisherBox.Location = New-Object System.Drawing.Point($discoveryFilterControlX, 39)
$discoveredPublisherBox.Width = 150
$discoveredPublisherBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
[void]$discoveredPublisherBox.Items.Add("<All Publishers>")
$discoveredPublisherBox.SelectedIndex = 0
$discoveredPublisherBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$tabDiscovered.Controls.Add($discoveredPublisherBox)

$discoveredSortLabel = New-Object System.Windows.Forms.Label
$discoveredSortLabel.Text = "Sort by:"
$discoveredSortLabel.Location = New-Object System.Drawing.Point($discoveryFilterLabelX, 69)
$discoveredSortLabel.AutoSize = $true
$discoveredSortLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$tabDiscovered.Controls.Add($discoveredSortLabel)

$discoveredSortBox = New-Object System.Windows.Forms.ComboBox
$discoveredSortBox.Location = New-Object System.Drawing.Point($discoveryFilterControlX, 66)
$discoveredSortBox.Width = 150
$discoveredSortBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
[void]$discoveredSortBox.Items.Add("Device Count")
[void]$discoveredSortBox.Items.Add("Alphabetical")
$discoveredSortBox.SelectedIndex = 0
$discoveredSortBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$tabDiscovered.Controls.Add($discoveredSortBox)

$discoveredListBox = New-Object System.Windows.Forms.CheckedListBox
$discoveredListBox.Location = New-Object System.Drawing.Point(20,110)
$discoveredListBox.Width = [Math]::Max(710, $tabDiscovered.ClientSize.Width - 40)
$discoveredListBox.Height = 325
$discoveredListBox.CheckOnClick = $true
$discoveredListBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$tabDiscovered.Controls.Add($discoveredListBox)

$script:discoveredRaw = [System.Collections.Generic.List[object]]::new()
$script:discoveredVisible = [System.Collections.Generic.List[object]]::new()
$script:isDiscoveryDeploymentActive = $false

# ==================================================
# Tab: Settings
# ==================================================
$tabSettings = New-Object System.Windows.Forms.TabPage
$tabSettings.Text = "Settings"
$tabControl.TabPages.Add($tabSettings)

# Settings Header
$settingsHeaderLabel = New-Object System.Windows.Forms.Label
$settingsHeaderLabel.Text = "Application Settings"
$settingsHeaderLabel.Location = New-Object System.Drawing.Point(20,20)
$settingsHeaderLabel.AutoSize = $true
$settingsHeaderLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$tabSettings.Controls.Add($settingsHeaderLabel)

# Default Package Path
$defaultPathLabel = New-Object System.Windows.Forms.Label
$defaultPathLabel.Text = "Default Package Folder:"
$defaultPathLabel.Location = New-Object System.Drawing.Point(20,60)
$defaultPathLabel.AutoSize = $true
$tabSettings.Controls.Add($defaultPathLabel)

$defaultPathTextBox = New-Object System.Windows.Forms.TextBox
$defaultPathTextBox.Location = New-Object System.Drawing.Point(200,57)
$defaultPathTextBox.Width = 400
$defaultPathTextBox.Text = if ($script:settings.DefaultPackagePath) { $script:settings.DefaultPackagePath } else { "C:\Temp" }
$tabSettings.Controls.Add($defaultPathTextBox)

$browsePathButton = New-Object System.Windows.Forms.Button
$browsePathButton.Text = "Browse..."
$browsePathButton.Location = New-Object System.Drawing.Point(610,55)
$browsePathButton.Width = 100
$tabSettings.Controls.Add($browsePathButton)

# Auto-Check Updates on Login
$autoCheckUpdatesCheckbox = New-Object System.Windows.Forms.CheckBox
$autoCheckUpdatesCheckbox.Text = "Check for updates on login"
$autoCheckUpdatesCheckbox.Location = New-Object System.Drawing.Point(20,100)
$autoCheckUpdatesCheckbox.AutoSize = $true
$autoCheckUpdatesCheckbox.Checked = if ($script:settings.AutoCheckUpdates) { $script:settings.AutoCheckUpdates } else { $false }
$tabSettings.Controls.Add($autoCheckUpdatesCheckbox)

# RememberMe Checkbox (moved to settings)
$rememberMeCheckbox = New-Object System.Windows.Forms.CheckBox
$rememberMeCheckbox.Text = "Remember last username"
$rememberMeCheckbox.Location = New-Object System.Drawing.Point(20,130)
$rememberMeCheckbox.AutoSize = $true
$rememberMeCheckbox.Checked = if ($script:settings.RememberMe) { $script:settings.RememberMe } else { $false }
$tabSettings.Controls.Add($rememberMeCheckbox)

# Save Settings Button
$saveSettingsButton = New-Object System.Windows.Forms.Button
$saveSettingsButton.Text = "Save Settings"
$saveSettingsButton.Location = New-Object System.Drawing.Point(20,180)
$saveSettingsButton.Width = 150
$saveSettingsButton.Height = 35
$tabSettings.Controls.Add($saveSettingsButton)

# Clear Version Cache Button
$clearCacheButton = New-Object System.Windows.Forms.Button
$clearCacheButton.Text = "Clear All Caches"
$clearCacheButton.Location = New-Object System.Drawing.Point(20,225)
$clearCacheButton.Width = 180
$clearCacheButton.Height = 35
$tabSettings.Controls.Add($clearCacheButton)

# Browse Path Button Handler
$browsePathButton.Add_Click({
  $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
  $folderBrowser.Description = "Select default package folder"
  $folderBrowser.SelectedPath = $defaultPathTextBox.Text
  
  if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $defaultPathTextBox.Text = $folderBrowser.SelectedPath
    Update-Status "Package folder path updated (not saved yet)"
  }
})

# Save Settings Button Handler
$saveSettingsButton.Add_Click({
  try {
    $script:settings.DefaultPackagePath = $defaultPathTextBox.Text
    $script:settings.AutoCheckUpdates = $autoCheckUpdatesCheckbox.Checked
    $script:settings.RememberMe = $rememberMeCheckbox.Checked
    $rememberCheckBox.Checked = $rememberMeCheckbox.Checked
    if (-not $script:settings.RememberMe) {
      $script:settings.LastUser    = ""
      $script:settings.RecentUsers = @()
    }
    
    # Update pathBox on WinGet Apps tab with new default
    if ($pathBox) {
      $pathBox.Text = $script:settings.DefaultPackagePath
    }
    
    [void](Export-WinTunerSettings -Settings $script:settings -Path $script:settingsPath)
    Update-Status "Settings saved successfully!"
    
    [System.Windows.Forms.MessageBox]::Show(
      "Settings have been saved successfully!",
      "Settings Saved",
      [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Information
    )
  } catch {
    Update-Status "Failed to save settings: $($_.Exception.Message)"
    Write-Log "Settings save error: $($_.Exception.Message)"
    
    [System.Windows.Forms.MessageBox]::Show(
      "Failed to save settings: $($_.Exception.Message)",
      "Error",
      [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Error
    )
  }
})

# Clear Version Cache Button Handler
$clearCacheButton.Add_Click({
  try {
    Clear-WingetVersionCache
    Clear-WinTunerDiscoveryCache
    Clear-WinTunerDetectedAppsCache

    Write-Log "All local caches cleared: WinGet versions, Discovery WinGet searches, Graph detected apps."
    Update-Status "All local caches cleared."
  } catch {
    Write-Log "Cache cleanup failed: $($_.Exception.Message)"
    Update-Status "Cache cleanup failed: $($_.Exception.Message)"
  }
})

# --- Self-Update Section in Settings Tab ---
$updateSectionLabel = New-Object System.Windows.Forms.Label
$updateSectionLabel.Text = "Application Updates"
$updateSectionLabel.Location = New-Object System.Drawing.Point(20, 270)
$updateSectionLabel.AutoSize = $true
$updateSectionLabel.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$tabSettings.Controls.Add($updateSectionLabel)

$currentVersionLabel = New-Object System.Windows.Forms.Label
$currentVersionLabel.Text = "Current version: v$($script:appVersion)"
$currentVersionLabel.Location = New-Object System.Drawing.Point(20, 300)
$currentVersionLabel.AutoSize = $true
$tabSettings.Controls.Add($currentVersionLabel)

$checkUpdateButton = New-Object System.Windows.Forms.Button
$checkUpdateButton.Text = "Check for Updates"
$checkUpdateButton.Location = New-Object System.Drawing.Point(20, 330)
$checkUpdateButton.Width = 180
$checkUpdateButton.Height = 35
$tabSettings.Controls.Add($checkUpdateButton)

$checkUpdateButton.Add_Click({

  $checkUpdateButton.Enabled = $false
  Invoke-AsyncUpdateCheck -StatusText "Checking for updates..." -OnComplete {
    param($updateResult)
    $checkUpdateButton.Enabled = $true
    Invoke-UpdateCheckFeedback -UpdateResult $updateResult -Context 'Manual'
  }
})

# Hashtable: AppName -> {PackageID, Version}
$script:packageMap = @{}
$script:isPackageSearchActive = $false
$script:packageSearchContext = $null
$script:isVersionLookupActive = $false
$script:versionLookupContext = $null
$script:isPackageBuildActive = $false
$script:packageBuildContext = $null
$script:isPackageUploadActive = $false
$script:packageUploadContext = $null

# Optional: user-chosen versions per PackageID
$script:selectedPackageVersions = @{}

function Get-UpdateCandidateDisplayText {
    param([Parameter(Mandatory)][object]$App)

    $current = if ([string]::IsNullOrWhiteSpace([string]$App.CurrentVersion)) { '?' } else { [string]$App.CurrentVersion }
    $latest = if ([string]::IsNullOrWhiteSpace([string]$App.LatestVersion)) { '?' } else { [string]$App.LatestVersion }
    return "$($App.Name) [$current -> $latest]"
}

function Resolve-WinTunerPackageRootForOperation {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Path,
        [switch]$CreateIfMissing
    )

    $validation = Test-WinTunerPackageRoot -RootPackageFolder $Path
    if (-not $validation.IsValid) {
        Update-Status "Invalid package folder: $($validation.Reason)"
        Write-Log "Package operation blocked ($($validation.ReasonCode)): $($validation.Reason)"
        [void][System.Windows.Forms.MessageBox]::Show(
            $validation.Reason,
            'Invalid Folder',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return $null
    }

    if ($CreateIfMissing -and -not (Test-Path -LiteralPath $validation.FullPath -PathType Container)) {
        try {
            New-Item -ItemType Directory -Path $validation.FullPath -Force -ErrorAction Stop | Out-Null
        }
        catch {
            Update-Status "Cannot create package folder: $($_.Exception.Message)"
            Write-Log "Package folder creation failed: $($_.Exception.Message)"
            return $null
        }
    }

    return [string]$validation.FullPath
}
function Update-PackageActionState {
    $uploadButton.Enabled = $false
    $uploadButton.Text = if ($script:isPackageUploadActive) { 'Uploading...' } else { 'Upload to Tenant' }

    if ($script:isPackageBuildActive -or $script:isPackageUploadActive) {
        return
    }

    try {
        if (-not $dropdown.SelectedItem) {
            return
        }

        $appName = [string]$dropdown.SelectedItem
        $package = $script:packageMap[$appName]

        if (-not $package -or [string]::IsNullOrWhiteSpace([string]$package.PackageID)) {
            return
        }

        $packageID = [string]$package.PackageID
        $desiredVersion = if ($script:selectedPackageVersions.ContainsKey($packageID)) {
            [string]$script:selectedPackageVersions[$packageID]
        } else {
            [string]$package.Version
        }

        if ([string]::IsNullOrWhiteSpace($desiredVersion)) {
            return
        }

        $validation = Test-WinTunerPackageArtifact `
            -RootPackageFolder ([string]$pathBox.Text) `
            -PackageId $packageID `
            -Version $desiredVersion

        if (-not $validation.IsValid) {
            if (
                $script:builtVersions.ContainsKey($packageID) -and
                ([string]$script:builtVersions[$packageID] -eq $desiredVersion)
            ) {
                $script:builtVersions.Remove($packageID)
                Write-Log "Invalidated package state for $packageID version $desiredVersion ($($validation.ReasonCode)): $($validation.Reason)"
            }

            return
        }

        $knownBuildMatches = (
            $script:builtVersions.ContainsKey($packageID) -and
            ([string]$script:builtVersions[$packageID] -eq $desiredVersion)
        )

        if (-not $knownBuildMatches) {
            $script:builtVersions[$packageID] = $desiredVersion
            Write-Log "Validated existing package for cross-session reuse: $packageID version $desiredVersion -> $($validation.IntuneWinPath)"
        }

        if (-not $script:isConnected) {
            return
        }

        $uploadButton.Enabled = $true
    } catch {
        $uploadButton.Enabled = $false
        Write-LogSafe "Package action state validation warning: $($_.Exception.Message)"
    }
}

function Update-PackageSearchActionState {
    try {
        $resultCount = if ($dropdown) { [int]$dropdown.Items.Count } else { 0 }
        $selectedIndex = if ($dropdown) { [int]$dropdown.SelectedIndex } else { -1 }
        $isBusy = (
            [bool]$script:isPackageSearchActive -or
            [bool]$script:isVersionLookupActive -or
            [bool]$script:isPackageBuildActive -or
            [bool]$script:isPackageUploadActive -or
            [bool]$script:isUpdateOperationActive -or
            [bool]$script:discoveryScanRunning -or
            [bool]$script:isDiscoveryDeploymentActive -or
            [bool]$script:isSupersededOperationActive
        )
        $state = Get-WinTunerPackageSearchActionState `
            -IsBusy $isBusy `
            -ResultCount $resultCount `
            -SelectedIndex $selectedIndex

        $searchButton.Enabled = $state.CanSearch
        $appSearchBox.Enabled = $state.CanEditQuery
        $dropdown.Enabled = $state.CanSelectResult
        $versionsButton.Text = if ($script:isVersionLookupActive) { 'Loading...' } else { 'Versions...' }
        $versionsButton.Enabled = $state.CanSelectVersion
        $createButton.Text = if ($script:isPackageBuildActive) { 'Creating...' } else { 'Create package' }
        $createButton.Enabled = $state.CanCreatePackage
        $pathBox.Enabled = -not [bool]($script:isPackageBuildActive -or $script:isPackageUploadActive)
        $browseButton.Enabled = -not [bool]($script:isPackageBuildActive -or $script:isPackageUploadActive)
    } catch {
        if ($searchButton) { $searchButton.Enabled = $false }
        if ($appSearchBox) { $appSearchBox.Enabled = $false }
        if ($dropdown) { $dropdown.Enabled = $false }
        if ($versionsButton) { $versionsButton.Enabled = $false }
        if ($createButton) { $createButton.Enabled = $false }
        Write-LogSafe "Package search action state warning: $($_.Exception.Message)"
    }
}
function Update-LogoutActionState {
    try {
        $logoutButton.Enabled = (
            [bool]$script:isConnected -and
            -not [bool]$script:isUpdateOperationActive -and
            -not [bool]$script:discoveryScanRunning -and
            -not [bool]$script:isDiscoveryDeploymentActive -and
            -not [bool]$script:isSupersededOperationActive -and
            -not [bool]$script:isPackageBuildActive -and
            -not [bool]$script:isPackageUploadActive
        )
    } catch {
        $logoutButton.Enabled = $false
        Write-LogSafe "Logout action state warning: $($_.Exception.Message)"
    }
}

function Update-UpdateActionState {
    try {
        $isDeployingUpdates = [bool]$script:updateDeploymentContext
        $updateSelectedButton.Text = if ($isDeployingUpdates) { 'Updating...' } else { 'Update checked apps' }
        $updateAllButton.Text = if ($isDeployingUpdates) { 'Updating...' } else { 'Update ALL (unchecked too)' }
        $candidates = @($script:updateApps)
        $checkedCount = @($candidates | Where-Object { $_ -and $_.Checked }).Count
        $state = Get-WinTunerUpdateActionState `
            -Connected ([bool]$script:isConnected) `
            -IsBusy ([bool]($script:isUpdateOperationActive -or $script:isSupersededOperationActive -or $script:isPackageSearchActive -or $script:isVersionLookupActive -or $script:isPackageBuildActive -or $script:isPackageUploadActive)) `
            -CandidateCount $candidates.Count `
            -CheckedCount $checkedCount `
            -IsScanRunning ([bool]$script:updateScanRunning) `
            -CancelRequested ([bool]$script:cancelUpdateScan)

        $updateSearchButton.Text = $state.SearchButtonText
        $updateSearchButton.Enabled = $state.CanSearch
        $checkAllButton.Enabled = $state.CanCheckAll
        $uncheckAllButton.Enabled = $state.CanUncheckAll
        $updateSelectedButton.Enabled = $state.CanUpdateSelected
        $updateAllButton.Enabled = $state.CanUpdateAll
        if ($updateSummaryLabel) {
            $updateSummaryLabel.Text = "Candidates: $($candidates.Count) | Checked: $checkedCount"
        }
    } catch {
        $updateSearchButton.Enabled = $false
        $checkAllButton.Enabled = $false
        $uncheckAllButton.Enabled = $false
        $updateSelectedButton.Enabled = $false
        $updateAllButton.Enabled = $false
        Write-LogSafe "Update action state warning: $($_.Exception.Message)"
    }

    Update-LogoutActionState
    Update-SupersededActionState
    Update-PackageSearchActionState
}

function Clear-DiscoveryCandidateState {
    $script:discoveredRaw = [System.Collections.Generic.List[object]]::new()
    $script:discoveredVisible = [System.Collections.Generic.List[object]]::new()
    $discoveredListBox.Items.Clear()
    $discoveredAppSearchBox.Text = ''
    $discoveredPublisherBox.BeginUpdate()
    $discoveredPublisherBox.Items.Clear()
    [void]$discoveredPublisherBox.Items.Add('<All Publishers>')
    $discoveredPublisherBox.SelectedIndex = 0
    $discoveredPublisherBox.EndUpdate()
    $lastDiscoveryLabel.Text = 'Last discovery: Never'
}

function Update-DiscoveryActionState {
    try {
        $results = @($script:discoveredRaw)
        $checkedCount = @($results | Where-Object { $_ -and $_.Checked }).Count
        $state = Get-WinTunerDiscoveryActionState `
            -Connected ([bool]$script:isConnected) `
            -IsScanning ([bool]$script:discoveryScanRunning) `
            -CancelRequested ([bool]$script:cancelDiscoveryScan) `
            -IsDeploying ([bool]$script:isDiscoveryDeploymentActive) `
            -IsOtherOperationActive ([bool]($script:isUpdateOperationActive -or $script:isSupersededOperationActive -or $script:isPackageSearchActive -or $script:isVersionLookupActive -or $script:isPackageBuildActive -or $script:isPackageUploadActive)) `
            -ResultCount $results.Count `
            -CheckedCount $checkedCount

        $scanDiscoveredButton.Text = if ($script:discoveryScanRunning) {
            if ($script:cancelDiscoveryScan) { 'Cancelling...' } else { 'Cancel Scan' }
        } else {
            '1. Scan Discovered Apps'
        }
        $scanDiscoveredButton.Enabled = $state.CanScan
        $deployDiscoveredButton.Enabled = $state.CanDeploy
        $exportDiscoveredCsvButton.Enabled = $state.CanExport
        $checkAllDiscoveredButton.Enabled = $state.CanCheckAll
        $uncheckAllDiscoveredButton.Enabled = $state.CanUncheckAll
    } catch {
        $scanDiscoveredButton.Enabled = $false
        $deployDiscoveredButton.Enabled = $false
        $exportDiscoveredCsvButton.Enabled = $false
        $checkAllDiscoveredButton.Enabled = $false
        $uncheckAllDiscoveredButton.Enabled = $false
        Write-LogSafe "Discovery action state warning: $($_.Exception.Message)"
    }

    Update-LogoutActionState
    Update-SupersededActionState
    Update-PackageSearchActionState
}

function Update-SupersededActionState {
    try {
        $results = @($script:supersededApps)
        $selectedIndex = if ($supersededDropdown) { [int]$supersededDropdown.SelectedIndex } else { -1 }
        $isBusy = (
            [bool]$script:isSupersededOperationActive -or
            [bool]$script:isUpdateOperationActive -or
            [bool]$script:discoveryScanRunning -or
            [bool]$script:isDiscoveryDeploymentActive -or
            [bool]$script:isPackageSearchActive -or
            [bool]$script:isVersionLookupActive -or
            [bool]$script:isPackageBuildActive -or
            [bool]$script:isPackageUploadActive
        )
        $state = Get-WinTunerSupersededActionState `
            -Connected ([bool]$script:isConnected) `
            -IsBusy $isBusy `
            -ResultCount $results.Count `
            -SelectedIndex $selectedIndex

        $supersededSearchButton.Enabled = $state.CanSearch
        $deleteSelectedAppButton.Enabled = $state.CanDeleteSelected
        $removeOldAppsButton.Enabled = $state.CanDeleteAll
    } catch {
        if ($supersededSearchButton) { $supersededSearchButton.Enabled = $false }
        if ($deleteSelectedAppButton) { $deleteSelectedAppButton.Enabled = $false }
        if ($removeOldAppsButton) { $removeOldAppsButton.Enabled = $false }
        Write-LogSafe "Superseded action state warning: $($_.Exception.Message)"
    }

    Update-LogoutActionState
    Update-PackageSearchActionState
}
# Cache for winget searches to speed up repeated searches
# (initialized at script scope; see earlier declaration)

# Module check
Update-Status "Checking WinTuner Module..."
try {
  if (Get-Module -ListAvailable -Name WinTuner) {
    Update-Status "Module found."
    
    # Check if update is available (optional - don't force update every time)
    # Uncomment the following block if you want automatic updates:
    <#
    try {
      $installedVersion = (Get-Module -ListAvailable -Name WinTuner | Sort-Object Version -Descending | Select-Object -First 1).Version
      $onlineVersion = (Find-Module -Name WinTuner -ErrorAction SilentlyContinue).Version
      
      if ($onlineVersion -and $onlineVersion -gt $installedVersion) {
        Update-Status "Module update available ($installedVersion → $onlineVersion). Updating..."
        
        # Temporarily disable PSDefaultParameterValues for Update-Module
        $savedDefaults = $PSDefaultParameterValues.Clone()
        $PSDefaultParameterValues.Clear()
        
        Update-Module -Name WinTuner -ErrorAction Stop
        
        # Restore defaults
        foreach ($key in $savedDefaults.Keys) {
          $PSDefaultParameterValues[$key] = $savedDefaults[$key]
        }
        
        Update-Status "Module updated to $onlineVersion"
      } else {
        Update-Status "Module is up to date (v$installedVersion)"
      }
    } catch { 
      Write-Log "Module update check failed: $($_.Exception.Message)"
      Update-Status "Module update skipped (using existing version)"
    }
    #>
  } else {
    Update-Status "Module not found, installing..."
    try { Install-Module -Name WinTuner -Scope CurrentUser -Repository PSGallery -Force -ErrorAction Stop } catch { Update-Status ("Module install failed: {0}" -f $_.Exception.Message) }
  }
} catch {
  Update-Status ("Module install/update error: {0}" -f $_.Exception.Message)
}
try { 
  Import-Module WinTuner -ErrorAction Stop 
} catch {
  $errMsg = $_.Exception.Message
  Write-Log "Failed to import WinTuner module: $errMsg"
  [System.Windows.Forms.MessageBox]::Show(
    "Failed to import WinTuner module.`n`nError: $errMsg`n`nPlease install it:`nInstall-Module WinTuner -Scope CurrentUser",
    "Module Import Failed",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error
  )
  # Disable all functional tabs except Settings
  foreach ($tab in $tabControl.TabPages) {
    if ($tab.Text -ne "Settings") { $tab.Enabled = $false }
  }
  if ($loginButton) { $loginButton.Enabled = $false }
}
Update-Status "Module imported."

# Login button
$loginButton = New-Object System.Windows.Forms.Button
$loginButton.Text = "Login to Tenant"
$loginButton.Location = New-Object System.Drawing.Point(584, 10)
$loginButton.Size = New-Object System.Drawing.Size(150, 27)
$headerPanel.Controls.Add($loginButton)

# initialize login button enabled state based on username validation
$loginButton.Enabled = (Test-ValidM365UserName -UserName $usernameBox.Text)

$rememberCheckBox = New-Object System.Windows.Forms.CheckBox
$rememberCheckBox.Text = "Remember me"
$rememberCheckBox.Location = New-Object System.Drawing.Point(461, 46)
$rememberCheckBox.AutoSize = $true
$rememberCheckBox.Checked = $false
$headerPanel.Controls.Add($rememberCheckBox)

function Update-WinTunerHeaderLayout {
    $rightMargin = 10
    $themeX = [Math]::Max(
        $rightMargin,
        $headerPanel.ClientSize.Width - $themeToggleButton.Width - $rightMargin
    )
    $fixedActionRight = [Math]::Max($loginButton.Right, $logoutButton.Right)
    $themeY = if ($themeX -lt ($fixedActionRight + 10)) { 44 } else { 8 }

    $themeToggleButton.Location = New-Object System.Drawing.Point($themeX, $themeY)
}

$headerPanel.Add_Resize({ Update-WinTunerHeaderLayout })
Update-WinTunerHeaderLayout

$script:settingsPath = Join-Path ([Environment]::GetFolderPath('ApplicationData')) 'WinTunerGUI\settings.json'
$script:settings = Import-WinTunerSettings -Path $script:settingsPath
$rememberCheckBox.Checked = [bool]$script:settings.RememberMe
$rememberMeCheckbox.Checked = [bool]$script:settings.RememberMe
if ($script:settings.RememberMe -and $script:settings.LastUser) { $usernameBox.Text = $script:settings.LastUser } else { $usernameBox.Text = "" }

# Populate username ComboBox with recent users (only if RememberMe is on)
if ($script:settings.RememberMe -and $script:settings.RecentUsers) {
  $usernameBox.Items.Clear()
  foreach ($u in @($script:settings.RecentUsers)) {
    if ($u) { [void]$usernameBox.Items.Add($u) }
  }
}

# Initialize pathBox with saved default package path
if ($pathBox) {
  if ($script:settings.DefaultPackagePath) {
    $pathBox.Text = $script:settings.DefaultPackagePath
  } else {
    $pathBox.Text = "C:\Temp"
  }
}

$rememberCheckBox.Add_CheckedChanged({
  try {
    $script:settings.RememberMe = [bool]$rememberCheckBox.Checked
    $rememberMeCheckbox.Checked = $rememberCheckBox.Checked
    if ($script:settings.RememberMe) { $script:settings.LastUser = $usernameBox.Text } else {
      $script:settings.LastUser = ""
      $script:settings.RecentUsers = @()
      $usernameBox.Items.Clear()
    }
    [void](Export-WinTunerSettings -Settings $script:settings -Path $script:settingsPath)
  } catch {
    Write-Log "Error in RememberMe checkbox handler: $($_.Exception.Message)"
  }
})

$loginButton.Add_Click({
  if (-not (Test-ValidM365UserName -UserName $usernameBox.Text)) {
    [void][System.Windows.Forms.MessageBox]::Show(
      "Please enter a valid M365 UPN.",
      "Invalid Username",
      [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    return
  }
  $loginButton.Enabled = $false
  $loginButton.Text = "Connecting..."
  [System.Windows.Forms.Application]::DoEvents()
  try {
    Update-Status "Connecting to tenant..."
    $script:isConnected = $false
    $null = Connect-WtWinTuner -Username $usernameBox.Text -ErrorAction Stop
    if (-not (Test-WtConnected)) { throw "Authentication error or failed." }
    $script:isConnected = $true
    Update-Status "Login success."
    $script:currentUserUpn = $usernameBox.Text
    Update-PackageActionState
    if ($loginInfoLabel) { $loginInfoLabel.Text = "Logged in as: $($script:currentUserUpn)" }
    if ($rememberCheckBox) { $script:settings.RememberMe = [bool]$rememberCheckBox.Checked }
    if ($script:settings.RememberMe) { $script:settings.LastUser = $usernameBox.Text } else { $script:settings.LastUser = "" }
    Add-RecentUser -Upn $usernameBox.Text
    # Update dropdown list
    $usernameBox.Items.Clear()
    foreach ($u in @($script:settings.RecentUsers)) {
      if ($u) { [void]$usernameBox.Items.Add($u) }
    }
    [void](Export-WinTunerSettings -Settings $script:settings -Path $script:settingsPath)
    Set-ConnectedUIState -Connected $true
    
    # Auto-check for updates if enabled
    if ($script:settings.AutoCheckUpdates) {
      Write-Log "Auto-check for updates enabled - triggering update search"
      Update-Status "Auto-checking for updates..."
      try {
        # Switch to Updates tab first so PerformClick works
        $tabControl.SelectedTab = $tabUpdate
        Start-Sleep -Milliseconds 100
        $updateSearchButton.PerformClick()
      } catch {
        Write-Log "Auto-check for updates failed: $($_.Exception.Message)"
      }
    }
  } catch {
    $msg = $_.Exception.Message
    if ($msg -imatch 'network|connection|timeout|unreachable') {
      [void][System.Windows.Forms.MessageBox]::Show(
        "Network error: Please check your internet connection.`n`nDetails: $msg",
        "Network Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
      )
    } elseif ($msg -imatch 'unauthorized|authentication|credential|access') {
      [void][System.Windows.Forms.MessageBox]::Show(
        "Authentication failed: Please check your credentials.`n`nDetails: $msg",
        "Authentication Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
      )
    } else {
      [void][System.Windows.Forms.MessageBox]::Show(
        "Login failed: $msg",
        "Login Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
      )
    }
    Update-Status ("Login canceled/failed: {0}" -f $msg)
    Set-ConnectedUIState -Connected $false
  } finally {
    $loginButton.Text = "Login to Tenant"
    $loginButton.Enabled = (Test-ValidM365UserName -UserName $usernameBox.Text)
  }
})

$searchButton.Add_Click({
  Start-WinTunerPackageSearch
})

$versionsButton.Add_Click({
  Start-WinTunerVersionLookup
})

$createButton.Add_Click({
  if (
    $script:isPackageSearchActive -or
    $script:isVersionLookupActive -or
    $script:isPackageBuildActive -or
    $script:isPackageUploadActive -or
    $script:isUpdateOperationActive -or
    $script:discoveryScanRunning -or
    $script:isDiscoveryDeploymentActive -or
    $script:isSupersededOperationActive
  ) {
    Update-PackageSearchActionState
    return
  }

  if (-not $dropdown.SelectedItem) { Update-Status 'Please select a package.'; return }
  $appName = [string]$dropdown.SelectedItem
  $package = $script:packageMap[$appName]
  if (-not $package -or [string]::IsNullOrWhiteSpace([string]$package.PackageID)) { Update-Status 'Selected item is invalid.'; return }
  $packageID = [string]$package.PackageID
  $folder = Resolve-WinTunerPackageRootForOperation -Path $pathBox.Text -CreateIfMissing
  if ([string]::IsNullOrWhiteSpace($folder)) { return }
  $desired = $null
  if ($script:selectedPackageVersions.ContainsKey($packageID)) {
    $desired = [string]$script:selectedPackageVersions[$packageID]
  }

  $targetVersion = if ($desired) { $desired } else { [string]$package.Version }
  $existingArtifact = Test-WinTunerPackageArtifact `
    -RootPackageFolder $folder `
    -PackageId $packageID `
    -Version $targetVersion

  if ($targetVersion -and $existingArtifact.IsValid) {
    $script:builtVersions[$packageID] = $targetVersion
    Update-Status ("Package already built (version {0}). Reusing validated package files." -f $targetVersion)
    Write-Log "Reusing validated package files for $packageID version $targetVersion -> $($existingArtifact.IntuneWinPath)"
    Update-PackageActionState
    return
  }

  if (
    $targetVersion -and
    $script:builtVersions.ContainsKey($packageID) -and
    ([string]$script:builtVersions[$packageID] -eq $targetVersion)
  ) {
    $script:builtVersions.Remove($packageID)
    Write-Log "Existing package for $packageID version $targetVersion failed validation ($($existingArtifact.ReasonCode)); rebuilding."
  }

  Start-WinTunerPackageBuild `
    -PackageId $packageID `
    -PackageFolder $folder `
    -DesiredVersion $desired `
    -LatestVersion ([string]$package.Version)
})
$uploadButton.Add_Click({
    if (
        $script:isPackageSearchActive -or
        $script:isVersionLookupActive -or
        $script:isPackageBuildActive -or
        $script:isPackageUploadActive -or
        $script:isUpdateOperationActive -or
        $script:discoveryScanRunning -or
        $script:isDiscoveryDeploymentActive -or
        $script:isSupersededOperationActive
    ) {
        Update-PackageActionState
        return
    }
    if (-not $script:isConnected) {
        [void][System.Windows.Forms.MessageBox]::Show(
            'Please login to your tenant first.',
            'Information',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return
    }
    if (-not $dropdown.SelectedItem) { Update-Status 'Please select a package.'; return }
    $appName = [string]$dropdown.SelectedItem
    $package = $script:packageMap[$appName]
    if (-not $package) { Update-Status 'Selected item is invalid.'; return }
    $packageID = [string]$package.PackageID
    $version = if ($script:selectedPackageVersions.ContainsKey($packageID)) {
        [string]$script:selectedPackageVersions[$packageID]
    } else {
        [string]$package.Version
    }
    if ([string]::IsNullOrWhiteSpace($version)) { Update-Status 'Version could not be determined.'; return }
    if ([string]::IsNullOrWhiteSpace($packageID)) { Update-Status 'Cannot upload: failed to resolve PackageId.'; return }
    $folder = Resolve-WinTunerPackageRootForOperation -Path $pathBox.Text
    if ([string]::IsNullOrWhiteSpace($folder)) { return }

    Start-WinTunerPackageUpload `
        -PackageId $packageID `
        -Version $version `
        -RootPackageFolder $folder
})
# ----------------------------------------------
# Check All / Uncheck All Buttons
# ----------------------------------------------
$checkAllButton.Add_Click({
  foreach ($app in $script:updateApps) { $app.Checked = $true }
  for ($i = 0; $i -lt $updateListBox.Items.Count; $i++) {
    $updateListBox.SetItemChecked($i, $true)
  }
  Update-UpdateActionState
  Update-Status "All update candidates checked ($($script:updateApps.Count) items)"
})

$uncheckAllButton.Add_Click({
  foreach ($app in $script:updateApps) { $app.Checked = $false }
  for ($i = 0; $i -lt $updateListBox.Items.Count; $i++) {
    $updateListBox.SetItemChecked($i, $false)
  }
  Update-UpdateActionState
  Update-Status "All update candidates unchecked"
})

# Save checked state when user checks/unchecks an item in the update list
$updateListBox.Add_ItemCheck({
  param($sender, $e)
  if ($e.Index -ge 0 -and $e.Index -lt $script:updateVisibleApps.Count) {
    $appObj = $script:updateVisibleApps[$e.Index]
    $appObj.Checked = ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked)
    Update-UpdateActionState
  }
})

# ----------------------------------------------
# Update List Filter - filters as you type (debounced 200ms)
# ----------------------------------------------
$updateFilterDebounceTimer = New-Object System.Windows.Forms.Timer
$updateFilterDebounceTimer.Interval = 200
$updateFilterDebounceTimer.Add_Tick({
  $updateFilterDebounceTimer.Stop()
  $filterText = $updateFilterBox.Text.Trim()

  # Clear and repopulate list with filtered items
  $updateListBox.BeginUpdate()
  $updateListBox.Items.Clear()
  $script:updateVisibleApps = [System.Collections.Generic.List[object]]::new()

  if ([string]::IsNullOrWhiteSpace($filterText)) {
    # No filter - show all apps
    foreach ($app in @($script:updateApps)) {
      if ($app -and $app.Name) {
        $idx = $updateListBox.Items.Add((Get-UpdateCandidateDisplayText -App $app))
        [void]$script:updateVisibleApps.Add($app)
        if ($app.Checked) { $updateListBox.SetItemChecked($idx, $true) }
      }
    }
  } else {
    # Filter apps by name (case-insensitive)
    $filtered = $script:updateApps | Where-Object {
      $_.Name -like "*$filterText*"
    }
    foreach ($app in @($filtered)) {
      if ($app -and $app.Name) {
        $idx = $updateListBox.Items.Add((Get-UpdateCandidateDisplayText -App $app))
        [void]$script:updateVisibleApps.Add($app)
        if ($app.Checked) { $updateListBox.SetItemChecked($idx, $true) }
      }
    }
  }
  $updateListBox.EndUpdate()
  Update-UpdateActionState
  # Update status with filter info
  if (-not [string]::IsNullOrWhiteSpace($filterText)) {
    Update-Status "Filter: $($updateListBox.Items.Count) apps match '$filterText'"
  }
})

$updateFilterBox.Add_TextChanged({
  $updateFilterDebounceTimer.Stop()
  $updateFilterDebounceTimer.Start()
})


# ----------------------------------------------
# Asynchronous and cancelable update scan
# ----------------------------------------------
$updateSearchButton.Add_Click({
  if ($script:updateScanRunning) {
    Request-WinTunerUpdateScanCancellation
    return
  }

  if (-not $script:isConnected) {
    Update-Status 'Please login to your tenant first.'
    return
  }

  Start-WinTunerUpdateScan
})

# -----------------------------
# UPDATED: Update Checked Apps flow
# -----------------------------
$updateSelectedButton.Add_Click({
    $checkedApps = [System.Collections.Generic.List[object]]::new()
    foreach ($app in $script:updateApps) {
        if ($app -and $app.Checked) {
            [void]$checkedApps.Add($app)
        }
    }

    if ($checkedApps.Count -eq 0) {
        Update-Status "No update candidates are checked."
        Write-Log "Update checked apps blocked: no candidates are checked."
        Update-UpdateActionState
        return
    }

    Write-Log "Starting update for $($checkedApps.Count) checked candidate(s), including filtered items."

    $confirmationLines = @($checkedApps | ForEach-Object { Get-UpdateCandidateDisplayText -App $_ })
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "The following checked apps will be updated:$([Environment]::NewLine)$([Environment]::NewLine)$($confirmationLines -join [Environment]::NewLine)",
        'Confirm checked updates',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
        Update-Status 'Checked app update canceled.'
        Write-Log 'Update checked apps canceled by user.'
        return
    }

    $rootPackageFolder = Resolve-WinTunerPackageRootForOperation -Path $pathBox.Text -CreateIfMissing
    if ([string]::IsNullOrWhiteSpace($rootPackageFolder)) { return }
    try {
        Update-Status "Starting update for $($checkedApps.Count) checked apps..."
        Start-WinTunerAppUpdateBatch -Apps @($checkedApps) -RootPackageFolder $rootPackageFolder -Mode Checked
    } catch {
        Update-Status "Update error: $($_.Exception.Message)"
        Write-Log "updateSelectedButton error: $($_.Exception.Message)"
    }
})


# -------------------------
# UPDATED: Update All flow
# -------------------------
$updateAllButton.Add_Click({
    $rootPackageFolder = Resolve-WinTunerPackageRootForOperation -Path $pathBox.Text -CreateIfMissing
    if ([string]::IsNullOrWhiteSpace($rootPackageFolder)) { return }

    $updatedApps = @($script:updateApps | Where-Object {
        $_ -and $_.LatestVersion -and $_.CurrentVersion -and
        (Test-IsNewerVersion $_.LatestVersion $_.CurrentVersion)
    } | Sort-Object Name)
    if (-not $updatedApps -or $updatedApps.Count -eq 0) {
        Update-Status "No update candidates found."
        return
    }

    # Show confirmation dialog
    $appNames = ($updatedApps | ForEach-Object { Get-UpdateCandidateDisplayText -App $_ }) -join [Environment]::NewLine
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "The following apps will be updated:$([Environment]::NewLine)$([Environment]::NewLine)$appNames",
        "Confirm",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
        Update-Status "Mass update canceled."
        return
    }

    try {
        Update-Status "Starting mass update for $($updatedApps.Count) apps..."
        Start-WinTunerAppUpdateBatch -Apps $updatedApps -RootPackageFolder $rootPackageFolder -Mode All
    } catch {
        Update-Status "Mass update error: $($_.Exception.Message)"
        Write-Log "updateAllButton error: $($_.Exception.Message)"
    }
})


$removeOldAppsButton.Add_Click({
  $supersededApps = @($script:supersededApps)
  if (-not $script:isConnected -or $script:isSupersededOperationActive) {
    Update-SupersededActionState
    return
  }
  if ($supersededApps.Count -eq 0) {
    Update-Status 'No Superseded Apps Found'
    Update-SupersededActionState
    return
  }

  $appNames = ($supersededApps | Select-Object -ExpandProperty Name) -join [Environment]::NewLine
  $result = [System.Windows.Forms.MessageBox]::Show(
    "The following outdated apps will be removed:$([Environment]::NewLine)$appNames",
    'Confirmation',
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Question
  )
  if ($result -ne [System.Windows.Forms.DialogResult]::Yes) {
    Update-Status 'Removal aborted.'
    return
  }

  $script:isSupersededOperationActive = $true
  $script:progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
  $script:progressBar.Minimum = 0
  $script:progressBar.Maximum = [Math]::Max(1, $supersededApps.Count)
  $script:progressBar.Value = 0
  $script:progressBar.Visible = $true
  Update-UpdateActionState
  Update-DiscoveryActionState
  Update-SupersededActionState

  $processedCount = 0
  try {
    foreach ($app in $supersededApps) {
      try {
        Remove-WtWin32App -GraphId $app.GraphId -ErrorAction Stop
        Update-Status ("Removed: {0}" -f $app.Name)
      } catch {
        if ($_.Exception.Message -match 'not found') {
          Write-Log "App already removed or not found in Intune: $($app.Name)"
          Update-Status "Already removed: $($app.Name)"
        } else {
          Update-Status ("Error removing {0}: {1}" -f $app.Name, $_.Exception.Message)
          Write-Log "Error while removal: $($_.Exception.Message)"
        }
      }
      $processedCount++
      $script:progressBar.Value = [Math]::Min($processedCount, $script:progressBar.Maximum)
    }
    Update-Status 'Deleted all superseded Apps.'
  } catch {
    Write-Log "Error removing superseded apps: $($_.Exception.Message)"
    Update-Status "Error: $($_.Exception.Message)"
  } finally {
    $script:isSupersededOperationActive = $false
    $script:progressBar.Value = 0
    $script:progressBar.Visible = $false
    Update-UpdateActionState
    Update-DiscoveryActionState
    Update-SupersededActionState
  }

  Start-WinTunerSupersededSearch
})

# Handler: Search superseded apps
$supersededSearchButton.Add_Click({
  Start-WinTunerSupersededSearch
})

$supersededDropdown.Add_SelectedIndexChanged({
  Update-SupersededActionState
})

# Handler: Delete selected superseded app
$deleteSelectedAppButton.Add_Click({
  $selectedIndex = $supersededDropdown.SelectedIndex
  if (
    -not $script:isConnected -or
    $script:isSupersededOperationActive -or
    $selectedIndex -lt 0 -or
    $selectedIndex -ge @($script:supersededApps).Count
  ) {
    Update-Status 'Please first select a superseded app from the dropdown.'
    Update-SupersededActionState
    return
  }

  $app = $script:supersededApps[$selectedIndex]
  $result = [System.Windows.Forms.MessageBox]::Show(
    "Delete App '$($app.Name)'?",
    'Confirmation',
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Question
  )
  if ($result -ne [System.Windows.Forms.DialogResult]::Yes) {
    Update-Status 'Removal aborted.'
    return
  }

  $script:isSupersededOperationActive = $true
  Update-UpdateActionState
  Update-DiscoveryActionState
  Update-SupersededActionState
  $removed = $false
  try {
    Remove-WtWin32App -GraphId $app.GraphId -ErrorAction Stop
    Update-Status ("Deleted: {0}" -f $app.Name)
    $removed = $true
  } catch {
    Update-Status ("Error while removal: {0}" -f $_.Exception.Message)
    Write-Log "Error while removing superseded app '$($app.Name)': $($_.Exception.Message)"
  } finally {
    $script:isSupersededOperationActive = $false
    Update-UpdateActionState
    Update-DiscoveryActionState
    Update-SupersededActionState
  }

  if ($removed) {
    Start-WinTunerSupersededSearch
  }
})
$logoutButton.Add_Click({
  try {
    Disconnect-WtWinTuner -ErrorAction Stop
  } catch {
    Write-Log "Logout warning: $($_.Exception.Message)"
  }
  Disconnect-WinTunerGraph
  $script:isConnected = $false
  $script:currentUserUpn = ""
  Update-PackageActionState
  if ($loginInfoLabel) { $loginInfoLabel.Text = "" }
  if (-not $script:settings.RememberMe) { $usernameBox.Text = "" }
  Update-Status "Logout success."
  Set-ConnectedUIState -Connected $false
})
# ==================================================
# ==================================================
# Discovered Apps Handlers
# ==================================================

# Handler für das Filtern / Sortieren
function Update-DiscoveredListUI {
    $discoveredListBox.BeginUpdate()
    $discoveredListBox.Items.Clear()
    $script:discoveredVisible = [System.Collections.Generic.List[object]]::new()
    
    $searchText = $discoveredAppSearchBox.Text
    $pubText = $discoveredPublisherBox.Text
    $sortType = $discoveredSortBox.Text

    # Use a List for efficient collection building
    $newFiltered = [System.Collections.Generic.List[object]]::new()

    foreach ($item in $script:discoveredRaw) {
        $match = $true

        # 1. Filtern nach Textfeld (DisplayName oder Winget-Name)
        if (-not [string]::IsNullOrWhiteSpace($searchText)) {
            $escapedSearch = [regex]::Escape($searchText)
            # Wenn der Text weder im Anzeigenamen noch im Winget-Namen vorkommt, ist es kein Match
            if (($item.DisplayName -notmatch "(?i)$escapedSearch") -and ($item.WingetApp.Name -notmatch "(?i)$escapedSearch") -and ($item.WingetApp.PackageID -notmatch "(?i)$escapedSearch")) {
                $match = $false
            }
        }

        # 2. Filtern nach Publisher (Dropdown)
        if ($match -and -not [string]::IsNullOrWhiteSpace($pubText) -and $pubText -ne "<All Publishers>") {
            $escapedPub = [regex]::Escape($pubText)
            if ($item.Publisher -notmatch "(?i)$escapedPub") {
                $match = $false
            }
        }

        # Wenn die App beide Filter übersteht, zum neuen Array hinzufügen
        if ($match) {
            $newFiltered.Add($item)
        }
    }
    
    # 3. Sortieren
    if ($sortType -eq "Alphabetical") {
        $newFiltered = $newFiltered | Sort-Object DisplayName
    } else {
        $newFiltered = $newFiltered | Sort-Object DeviceCount -Descending
    }
    
    # 4. In die sichtbare ListBox einfügen
    if ($newFiltered) {
        foreach ($obj in $newFiltered) {
            $idx = $discoveredListBox.Items.Add($obj.DisplayText)
            [void]$script:discoveredVisible.Add($obj)
            # Stellt den Haken (Checked-Status) wieder her, falls er vorher gesetzt war
            $discoveredListBox.SetItemChecked($idx, $obj.Checked)
        }
    }
    $discoveredListBox.EndUpdate()
    Update-DiscoveryActionState
}

# Listener für das Suchfeld (Text-Eingabe) – debounced 200ms
$discoveredSearchDebounceTimer = New-Object System.Windows.Forms.Timer
$discoveredSearchDebounceTimer.Interval = 200
$discoveredSearchDebounceTimer.Add_Tick({
  $discoveredSearchDebounceTimer.Stop()
  Update-DiscoveredListUI
})

$discoveredAppSearchBox.Add_TextChanged({
  $discoveredSearchDebounceTimer.Stop()
  $discoveredSearchDebounceTimer.Start()
})

# Listener für das Publisher-Dropdown
$discoveredPublisherBox.Add_SelectedIndexChanged({ Update-DiscoveredListUI })

# Listener für das Sortierungs-Dropdown
$discoveredSortBox.Add_SelectedIndexChanged({ Update-DiscoveredListUI })

# Wenn ein Haken gesetzt/entfernt wird, Zustand im Array speichern (überlebt Filterung!)
$discoveredListBox.Add_ItemCheck({
    param($sender, $e)
    if ($e.Index -ge 0 -and $e.Index -lt $script:discoveredVisible.Count) {
        $obj = $script:discoveredVisible[$e.Index]
        $obj.Checked = ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked)
        Update-DiscoveryActionState
    }
})

$checkAllDiscoveredButton.Add_Click({
    foreach ($obj in $script:discoveredRaw) { $obj.Checked = $true }
    Update-DiscoveredListUI
    Update-Status "All discovered apps checked ($($script:discoveredRaw.Count) items)"
})

$uncheckAllDiscoveredButton.Add_Click({
    foreach ($obj in $script:discoveredRaw) { $obj.Checked = $false }
    Update-DiscoveredListUI
    Update-Status "All discovered apps unchecked"
})

$script:discoveryScanRunning = $false
$script:cancelDiscoveryScan = $false

$scanDiscoveredButton.Add_Click({
  if ($script:discoveryScanRunning) {
    $script:cancelDiscoveryScan = $true
    Update-DiscoveryActionState
    Update-Status "Cancel requested - finishing current WinGet query..."
    Write-Log "Discovery scan cancellation requested by user."
    return
  }

  if (-not $script:isConnected) {
    Update-Status "Please login first."
    Update-DiscoveryActionState
    return
  }

  $script:discoveryScanRunning = $true
  $script:cancelDiscoveryScan = $false
  Update-DiscoveryActionState

  # Speichere die originalen Streams und schalte sie stumm, um Threading-Crashes zu vermeiden
  $oldProgress = $ProgressPreference
  $oldInfo = $InformationPreference
  $ProgressPreference = 'SilentlyContinue'
  $InformationPreference = 'SilentlyContinue'

  try {
    Clear-DiscoveryCandidateState
    Update-DiscoveryActionState
    
    $script:progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
    $script:progressBar.Visible = $true
    [System.Windows.Forms.Application]::DoEvents()  # TODO: refactor to use Invoke-AsyncOperation

# --- GRAPH-AUTH BLOCK (FIXED) ---
    # Ensure Microsoft Graph session has the required account and scopes
    Update-Status "Checking Microsoft Graph session..."
    [System.Windows.Forms.Application]::DoEvents()

    $null = Connect-WinTunerGraph -UserPrincipalName $script:currentUserUpn

    # 1. Vorhandene Apps checken (EXTREM SCHNELL DURCH "Resolve" STATT "Try-Resolve")
    Update-Status "Loading existing managed apps to filter them out..."
    [System.Windows.Forms.Application]::DoEvents()  # TODO: refactor to use Invoke-AsyncOperation
    $existingApps = @(Get-WtWin32Apps -Superseded:$false -ErrorAction SilentlyContinue 2>$null 3>$null 4>$null 5>$null 6>$null)
    $existingPackageIds = [System.Collections.Generic.List[object]]::new()
    foreach ($eApp in $existingApps) {
		[System.Windows.Forms.Application]::DoEvents()  # TODO: refactor to use Invoke-AsyncOperation
        $id = Resolve-WtWingetId -AppOrResult $eApp 2>$null 3>$null 4>$null 5>$null 6>$null
        if ($id) { $existingPackageIds.Add($id) }
    }

    # 2. Hole ALLE Discovered Apps aus Intune (inklusive Paginierung)
    Update-Status "Fetching ALL detected apps from Intune API (this might take a moment)..."
    [System.Windows.Forms.Application]::DoEvents()  # TODO: refactor to use Invoke-AsyncOperation
    
    $detectedResult = Get-WinTunerDetectedApps -PageSize 500 -MaxPages 1000

    if ($detectedResult.FromCache) {
        Write-Log "Detected apps loaded from Graph cache: $(@($detectedResult.Apps).Count) apps."
    } else {
        Write-Log "Detected apps loaded from Microsoft Graph: $(@($detectedResult.Apps).Count) apps across $($detectedResult.PageCount) pages."
    }
    $detectedApps = $detectedResult.Apps

    if ($detectedResult.LimitReached) {
        Write-Log "Warning: Graph API pagination limit (100 pages) reached. Some apps may not be shown."
    }

    if (-not $detectedApps -or $detectedApps.Count -eq 0) {
        Update-Status "No discovered apps found in Intune."
        return
    }

    $filteredApps = @($detectedApps | Where-Object { 
        $_.publisher -notmatch "(?i)Intel|HP|Dell|Lenovo|AMD|NVIDIA|Realtek|Synaptics|VMware" 
    })

    $total = $filteredApps.Count
    $matchCount = 0              # unique PackageIDs shown in UI
    $matchedRawCount = 0         # total matched detected apps (before dedupe)

    # Prepare normalized list first (phase 1) so matching can run with cached query results (phase 2)
    $normalizedApps = [System.Collections.Generic.List[object]]::new()
    $skippedNonCandidateCount = 0
    foreach ($app in $filteredApps) {
        # 1. Entfernt restlos alles, was in Klammern steht (z.B. "(x64 de)", "(x86 en-US)")
        $searchName = $app.displayName -replace '\s*\([^)]*\)', ''
        # 2. Entfernt typische Versionsnummern, die aus Zahlen und Punkten bestehen
        $searchName = $searchName -replace '\s+[\d\.]+', ''
        $searchName = $searchName.Trim()
        if ([string]::IsNullOrWhiteSpace($searchName)) { continue }
        if (-not (Test-WingetSearchCandidate -DisplayName $searchName)) {
            if ($script:skipLowValueWingetCandidates) {
                $skippedNonCandidateCount++
                continue
            }
        }
        $normalizedApps.Add([pscustomobject]@{
            App        = $app
            SearchName = $searchName
        })
    }

    # Cache Search-WtWinGetPackage results by normalized search term
    # to reduce expensive/repetitive module calls in large environments
    $searchResultCache = @{}
    # Fast lookup for already created discovered entries by PackageID
    $discoveredByPackageId = @{}

    $uniqueSearchNameSet = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    $uniqueSearchNames = @(
        foreach ($normalizedApp in $normalizedApps) {
            $candidateSearchName = [string]$normalizedApp.SearchName

            if (
                -not [string]::IsNullOrWhiteSpace($candidateSearchName) -and
                $uniqueSearchNameSet.Add($candidateSearchName)
            ) {
                $candidateSearchName
            }
        }
    )
    $queryTotal = $uniqueSearchNames.Count
    $queryCurrent = 0
    Update-Status "Prepared $($normalizedApps.Count) apps for matching ($queryTotal unique search terms, skipped: $skippedNonCandidateCount, skip-mode: $($script:skipLowValueWingetCandidates))."
    Write-Log "Discovery prep -> Filtered apps: $total, Normalized apps: $($normalizedApps.Count), Unique search terms: $queryTotal, Skipped non-candidates: $skippedNonCandidateCount, Skip-mode: $($script:skipLowValueWingetCandidates)"

    # Phase 1: fetch/search all unique terms using isolated worker processes
    $script:progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
    $script:progressBar.MarqueeAnimationSpeed = 25

    Update-Status "WinGet discovery: processing $queryTotal unique search terms in isolated workers..."
    Write-Log "Discovery WinGet batch search starting -> Queries: $queryTotal, Batch size: 25, Cache TTL: 24h"
    [System.Windows.Forms.Application]::DoEvents()

    try {
        $batchResult = Search-WinTunerDiscoveryPackagesBatchCached `
            -SearchQueries $uniqueSearchNames `
            -BatchSize 25 `
            -QueryTimeoutSeconds 12 `
            -CacheTtlHours 24 `
            -OnWait {
                [System.Windows.Forms.Application]::DoEvents()
            } `
            -ShouldCancel {
                return [bool]$script:cancelDiscoveryScan
            }

        if ($batchResult.Canceled -or $script:cancelDiscoveryScan) {
            Clear-DiscoveryCandidateState
            Update-Status "Discovery scan canceled during WinGet search phase."
            Write-Log "Discovery scan canceled during isolated WinGet worker phase; partial results discarded."
            return
        }

        $queryCurrent = 0

        foreach ($searchResult in @($batchResult.Results)) {
            $queryCurrent++
            $searchName = [string]$searchResult.Query

            if ([string]::IsNullOrWhiteSpace($searchName)) {
                continue
            }

            if ([bool]$searchResult.Success) {
                $searchResultCache[$searchName] = @($searchResult.Results)
            } else {
                $searchResultCache[$searchName] = @()
                Write-Log "Search failed for '$searchName': $($searchResult.Error)"
            }
        }

        Write-Log "Discovery WinGet batch search complete -> Queries: $($batchResult.TotalQueries), Cache hits: $($batchResult.CacheHits), Worker queries: $($batchResult.WorkerQueries), Workers: $($batchResult.WorkerCount)"
        Update-Status "WinGet searches complete: $($batchResult.TotalQueries) queries, $($batchResult.CacheHits) cache hits, $($batchResult.WorkerCount) workers."
    }
    catch {
        Write-Log "Discovery WinGet batch search failed: $($_.Exception.Message)"
        Update-Status "Discovery WinGet search failed: $($_.Exception.Message)"
        throw
    }
    finally {
        $script:progressBar.MarqueeAnimationSpeed = 0
        $script:progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
    }

    # Phase 2: match normalized discovered apps against cached results
    $processTotal = $normalizedApps.Count
    $processCurrent = 0
    $script:progressBar.Maximum = if ($processTotal -gt 0) { $processTotal } else { 1 }
    $script:progressBar.Value = 0

    foreach ($entry in $normalizedApps) {
        [System.Windows.Forms.Application]::DoEvents()

        if ($script:cancelDiscoveryScan) {
            Clear-DiscoveryCandidateState
            Update-Status "Discovery scan canceled during matching."
            Write-Log "Discovery scan canceled during matching phase; partial results discarded."
            return
        }

        $processCurrent++
        $script:progressBar.Value = $processCurrent
        if (($processCurrent -eq 1) -or ($processCurrent % 25 -eq 0) -or ($processCurrent -eq $processTotal)) {
            Update-Status "Matching apps ($processCurrent/$processTotal): $($entry.App.displayName)..."
            [System.Windows.Forms.Application]::DoEvents()  # TODO: refactor to use Invoke-AsyncOperation
        }

        try {
            $app = $entry.App
            $searchName = $entry.SearchName
            $wingetResults = if ($searchResultCache.ContainsKey($searchName)) { @($searchResultCache[$searchName]) } else { @() }

            $bestMatch = $null
            $highestScore = 0
            foreach ($wgApp in $wingetResults) {
                $score = Get-StringSimilarity -str1 $app.displayName -str2 $wgApp.Name
                if ($score -gt $highestScore) {
                    $highestScore = $score
                    $bestMatch = $wgApp
                }
            }

            if ($bestMatch -and $highestScore -ge 50) {
                $matchedRawCount++
                if ($existingPackageIds -contains $bestMatch.PackageID) { continue }

                # Prüfen, ob diese Winget-App (PackageID) bereits vorhanden ist
                $existingEntry = $null
                if ($discoveredByPackageId.ContainsKey($bestMatch.PackageID)) {
                    $existingEntry = $discoveredByPackageId[$bestMatch.PackageID]
                }

                if ($existingEntry) {
                    # App existiert bereits in der Liste: Wir addieren die Geräteanzahl (DeviceCount)
                    $existingEntry.DeviceCount += $app.deviceCount
                    $existingEntry.MatchScore = [Math]::Max([double]$existingEntry.MatchScore, [double]$highestScore)
                    # Den Anzeigetext mit der neuen, kombinierten Anzahl aktualisieren
                    $existingEntry.DisplayText = "[$($existingEntry.DeviceCount) PCs] $($existingEntry.DisplayName) ($($existingEntry.Publisher))  -->  Winget: $($existingEntry.WingetApp.Name) [$($existingEntry.WingetApp.PackageID)] | Match: $([Math]::Round($existingEntry.MatchScore))%"
                } else {
                    # App ist neu: Wir nutzen den sauberen Winget-Namen (ohne Versionsnummern aus Intune)
                    $cleanName = $bestMatch.Name
                    $itemObj = [pscustomobject]@{
                        DisplayName = $cleanName
                        Publisher   = $app.publisher
                        DeviceCount = $app.deviceCount
                        WingetApp   = $bestMatch
                        MatchScore  = [double]$highestScore
                        Checked     = $false
                        DisplayText = "[$($app.deviceCount) PCs] $cleanName ($($app.publisher))  -->  Winget: $($bestMatch.Name) [$($bestMatch.PackageID)] | Match: $([Math]::Round($highestScore))%"
                    }
                    [void]$script:discoveredRaw.Add($itemObj)
                    $discoveredByPackageId[$bestMatch.PackageID] = $itemObj
                    $matchCount++
                }
            }
        } catch {
            Write-Log "Failed to process '$($entry.App.displayName)': $($_.Exception.Message)"
        }
    }
    
# --- NEU: Befülle das Publisher-Dropdown mit eindeutigen Werten ---
    $uniquePublishers = $script:discoveredRaw | Select-Object -ExpandProperty Publisher -Unique | Sort-Object
    
    $discoveredPublisherBox.BeginUpdate()
    $discoveredPublisherBox.Items.Clear()
    [void]$discoveredPublisherBox.Items.Add("<All Publishers>")
    foreach ($pub in $uniquePublishers) {
        if (-not [string]::IsNullOrWhiteSpace($pub)) {
            [void]$discoveredPublisherBox.Items.Add($pub)
        }
    }
    $discoveredPublisherBox.SelectedIndex = 0
    $discoveredPublisherBox.EndUpdate()

    # Befüllt die Liste initial mit Sortierung
    Update-DiscoveredListUI

    $lastDiscoveryTime = Get-Date
    $lastDiscoveryLabel.Text = "Last discovery: $($lastDiscoveryTime.ToString('HH:mm:ss'))"

    $graphSource = if ([bool]$detectedResult.FromCache) { "Cached" } else { "Fresh" }
    $wingetCacheSummary = "$($batchResult.CacheHits)/$($batchResult.TotalQueries) cached"

    if ($matchCount -gt 0) {
        Update-Status "Scanned: $($detectedApps.Count) | Filtered: $total | Matched apps: $matchedRawCount | Unique packages: $matchCount | Graph: $graphSource | WinGet: $wingetCacheSummary"
        Write-Log "Discovery summary -> Scanned: $($detectedApps.Count), Filtered: $total, Matched apps: $matchedRawCount, Unique packages: $matchCount, Graph: $graphSource, WinGet cache: $wingetCacheSummary"
    } else {
        Update-Status "No Winget matches found (or all are already managed). | Graph: $graphSource | WinGet: $wingetCacheSummary"
        Write-Log "Discovery summary -> No Winget matches, Graph: $graphSource, WinGet cache: $wingetCacheSummary"
    }

  } catch {
    Clear-DiscoveryCandidateState
    Update-Status "Error fetching discovered apps: $($_.Exception.Message)"
    Write-Log "Scan Discovered Error: $($_.Exception.Message); partial results discarded."
  } finally {
    try {
        Save-WinTunerDiscoveryCache
    } catch {
        Write-Log "Could not save Discovery WinGet cache: $($_.Exception.Message)"
    }

    $ProgressPreference = $oldProgress
    $InformationPreference = $oldInfo
    $script:discoveryScanRunning = $false
    $script:cancelDiscoveryScan = $false
    Update-DiscoveryActionState
    $script:progressBar.Maximum = 100
    $script:progressBar.Value = 0
    $script:progressBar.Visible = $false
  }
})

$deployDiscoveredButton.Add_Click({
    $checkedItems = @($script:discoveredRaw | Where-Object { $_.Checked })
    if ($checkedItems.Count -eq 0) { 
        Update-Status "No apps checked."
        return 
    }

    $rootFolder = $script:settings.DefaultPackagePath
    if ([string]::IsNullOrWhiteSpace([string]$rootFolder)) { $rootFolder = 'C:\Temp' }
    $rootFolder = Resolve-WinTunerPackageRootForOperation -Path $rootFolder -CreateIfMissing
    if ([string]::IsNullOrWhiteSpace($rootFolder)) { return }
    $oldProgress = $ProgressPreference
    $oldInfo = $InformationPreference
    $ProgressPreference = 'SilentlyContinue'
    $InformationPreference = 'SilentlyContinue'

    $script:isDiscoveryDeploymentActive = $true
    Update-DiscoveryActionState

    try {
        $script:progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
        $script:progressBar.Maximum = $checkedItems.Count
        $script:progressBar.Value = 0
        $script:progressBar.Visible = $true

        $successCount = 0
        $failedCount = 0
        $i = 0
        $successfulItems = [System.Collections.Generic.List[object]]::new()

        foreach ($item in $checkedItems) {
            $i++
            $script:progressBar.Value = $i
            $wingetApp = $item.WingetApp
            
            Update-Status "Packaging & Deploying ($i/$($checkedItems.Count)): $($wingetApp.Name)..."
            [System.Windows.Forms.Application]::DoEvents()  # TODO: refactor to use Invoke-AsyncOperation

            try {
                $packageId = $wingetApp.PackageID
                $version = $wingetApp.Version
                
                Write-Log "Creating package for discovered app: $packageId v$version"
                $pkgRes = New-WingetPackageWithFallback `
                    -PackageId $packageId `
                    -PackageFolder $rootFolder `
                    -LatestVersion $version `
                    -ErrorAction Stop
                
                if (-not $pkgRes -or -not $pkgRes.Succeeded) {
                    $packageError = if ($pkgRes -and $pkgRes.ErrorMessage) { $pkgRes.ErrorMessage } else { 'Unknown package creation error.' }
                    throw "Package creation failed: $packageError"
                }

                $effVersion = if ($pkgRes.EffectiveVersion) { $pkgRes.EffectiveVersion } else { $version }
                $artifactValidation = Test-WinTunerPackageArtifact `
                    -RootPackageFolder $rootFolder `
                    -PackageId $packageId `
                    -Version ([string]$effVersion)

                if (-not $artifactValidation.IsValid) {
                    throw "Package validation failed ($($artifactValidation.ReasonCode)): $($artifactValidation.Reason)"
                }

                Write-Log "Validated discovered-app package before upload: $packageId v$effVersion -> $($artifactValidation.IntuneWinPath)"
                Write-Log "Uploading new app to tenant: $packageId v$effVersion"
                Deploy-WtWin32App `
                    -PackageId $packageId `
                    -Version $effVersion `
                    -RootPackageFolder $rootFolder `
                    -ErrorAction Stop
                
                $successCount++
                [void]$successfulItems.Add($item)
                Write-Log "Successfully deployed new app: $packageId"
            } catch {
                $failedCount++
                Write-Log "Failed to deploy $($wingetApp.Name): $($_.Exception.Message)"
            }
        }

        foreach ($successfulItem in $successfulItems) {
            [void]$script:discoveredRaw.Remove($successfulItem)
        }
        Update-DiscoveredListUI

        Update-Status "Deployment complete: $successCount successful, $failedCount failed."
        [System.Windows.Forms.MessageBox]::Show(
            "Deployment finished!`n`nSuccessful: $successCount`nFailed: $failedCount`n`nNewly deployed apps will now appear in your Intune tenant.",
            "Deploy Complete",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )

    } catch {
        Update-Status "Deployment error: $($_.Exception.Message)"
        Write-Log "Deploy Discovered Apps Error: $($_.Exception.Message)"
    } finally {
        $ProgressPreference = $oldProgress
        $InformationPreference = $oldInfo
        $script:isDiscoveryDeploymentActive = $false
        Update-DiscoveryActionState
        $script:progressBar.Maximum = 100
        $script:progressBar.Value = 0
        $script:progressBar.Visible = $false
    }
})

$exportDiscoveredCsvButton.Add_Click({
    if (-not $script:discoveredRaw -or $script:discoveredRaw.Count -eq 0) {
        Update-Status "No discovered Winget matches to export."
        return
    }

    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Title = "Export Discovered Winget IDs"
    $sfd.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $sfd.FileName = ("Discovered_WingetIDs_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

    if ($sfd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Update-Status "CSV export canceled."
        return
    }

    try {
        $rows = @($script:discoveredRaw | Sort-Object DisplayName | ForEach-Object {
            [pscustomobject]@{
                DisplayName   = $_.DisplayName
                Publisher     = $_.Publisher
                DeviceCount   = $_.DeviceCount
                MatchScore    = [Math]::Round([double]$_.MatchScore)
                WingetName    = $_.WingetApp.Name
                WingetId      = $_.WingetApp.PackageID
                WingetVersion = $_.WingetApp.Version
            }
        })

        $rows | Export-Csv -Path $sfd.FileName -NoTypeInformation -Encoding utf8
        Write-Log "Exported discovered Winget IDs: $($rows.Count) row(s) -> $($sfd.FileName)"
        Update-Status "Export completed: $($rows.Count) row(s) saved to $($sfd.FileName)"
    } catch {
        Write-Log "Export discovered Winget IDs failed: $($_.Exception.Message)"
        Update-Status "CSV export failed: $($_.Exception.Message)"
    }
})
# Apply initial theme (Dark by default)
Set-GuiTheme -control $form -theme $script:currentTheme

# Safe logger for closing context
function Write-FileLog {
  param([string]$message)
  try { Write-Log $message } catch {}
}

# Re-entrancy protection for closing
$script:_closingInProgress = $false
$form.Add_FormClosing({
    param($sender, [System.Windows.Forms.FormClosingEventArgs]$e)
    
    # 1. Einstellungen speichern
    try { 
        if ($script:settings) { 
            if ($script:settings.RememberMe) { $script:settings.LastUser = $usernameBox.Text } 
            else { $script:settings.LastUser = "" }
            [void](Export-WinTunerSettings -Settings $script:settings -Path $script:settingsPath)
        } 
    } catch {}

    # 2. Wenn bereits geschlossen wird, ignorieren
    if ($script:_closingInProgress) { return }
    $script:_closingInProgress = $true

    # 3. Falls verbunden, regulär abmelden
    if ($script:isConnected) {
        try {
            $form.Enabled = $false
            if ($script:statusLabel) { 
                Update-Status "Closing... signing out from tenant"
                # Zwingt die UI, sich noch einmal schnell zu aktualisieren, bevor sie blockiert wird
                [System.Windows.Forms.Application]::DoEvents()  # TODO: refactor to use Invoke-AsyncOperation
            }
        } catch {}

        Write-FileLog 'Shutdown: starting tenant disconnect.'

        try {
            # Use Start-ThreadJob (PS 7+) so the WinTuner module is available in the same process
            $job = Start-ThreadJob { Disconnect-WtWinTuner }
            $null = Wait-Job $job -Timeout 5
            if ($job.State -ne 'Completed') {
                Write-FileLog 'Shutdown: disconnect timed out after 5s, closing anyway.'
            }
            Remove-Job $job -Force -ErrorAction SilentlyContinue
        } catch {
            Write-FileLog "FormClosing disconnect warning: $($_.Exception.Message)"
        }
        
        Write-FileLog 'Shutdown: disconnect finished. Closing form.'
        $script:isConnected = $false
    }
})

# ==================================================
# Globale Fehlererfassung (Crashes & unhandled Exceptions)
# ==================================================
try {
    # Fängt Abstürze ab, die direkt durch die Benutzeroberfläche (Klicks etc.) passieren
    [System.Windows.Forms.Application]::add_ThreadException({
        param($sender, $e)
        $ex = $e.Exception
        $errMsg = "FATAL UI ERROR: $($ex.Message)`n$($ex.StackTrace)"
        Write-FileLog $errMsg
    })
    
    # Fängt tieferliegende System- und PowerShell-Abstürze ab
    [System.AppDomain]::CurrentDomain.add_UnhandledException({
        param($sender, $e)
        $ex = $e.ExceptionObject
        $errMsg = "FATAL APP ERROR: $($ex.Message)`n$($ex.StackTrace)"
        Write-FileLog $errMsg
    })
} catch {
    # Ignoriere Fehler, falls die Event-Registrierung in älteren PS-Versionen zickt
}

# Async update check on startup so it doesn't block the UI
$form.Add_Shown({
  Invoke-AsyncUpdateCheck -StatusText "Checking for updates..." -OnComplete {
    param($updateResult)
    Invoke-UpdateCheckFeedback -UpdateResult $updateResult -Context 'Startup'
  }
})

# Tooltips for main buttons
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 5000
$toolTip.InitialDelay = 500
$toolTip.ReshowDelay  = 500
$toolTip.ShowAlways   = $true

$toolTip.SetToolTip($searchButton,          "Search the WinGet repository without blocking the window")
$toolTip.SetToolTip($versionsButton,        "Load and select a specific version without blocking the window")
$toolTip.SetToolTip($browseButton,          "Choose the local folder to store package files")
if ($createButton)          { $toolTip.SetToolTip($createButton,          "Create the .wtpackage file locally") }
if ($uploadButton)          { $toolTip.SetToolTip($uploadButton,          "Upload and deploy the package to Microsoft Intune without blocking the window") }
if ($updateSearchButton)    { $toolTip.SetToolTip($updateSearchButton,    "Scan Intune Win32 apps for WinGet updates; click again to cancel an active scan") }
if ($updateAllButton)       { $toolTip.SetToolTip($updateAllButton,       "Update all available apps in the background") }
if ($updateSelectedButton)  { $toolTip.SetToolTip($updateSelectedButton,  "Update only the checked apps in the background") }
if ($scanDiscoveredButton)  { $toolTip.SetToolTip($scanDiscoveredButton,  "Scan Intune Discovered Apps and match them to WinGet packages") }
if ($logoutButton)          { $toolTip.SetToolTip($logoutButton,          "Disconnect from the current Microsoft 365 tenant") }
if ($themeToggleButton)     { $toolTip.SetToolTip($themeToggleButton,     "Switch between Dark Mode and Light Mode") }
if ($clearHistoryButton)    { $toolTip.SetToolTip($clearHistoryButton,    "Clears the list of saved M365 login names") }

# Header / Login area
if ($loginButton)           { $toolTip.SetToolTip($loginButton,           "Sign in to your Microsoft 365 tenant") }
if ($rememberCheckBox)      { $toolTip.SetToolTip($rememberCheckBox,      "Save your username so it is pre-filled on the next launch") }

# tabUpdate
if ($checkAllButton)        { $toolTip.SetToolTip($checkAllButton,        "Check all apps in the update list") }
if ($uncheckAllButton)      { $toolTip.SetToolTip($uncheckAllButton,      "Uncheck all apps in the update list") }
if ($supersededSearchButton){ $toolTip.SetToolTip($supersededSearchButton,"Search for outdated (superseded) app versions in Intune without blocking the window") }
if ($deleteSelectedAppButton){ $toolTip.SetToolTip($deleteSelectedAppButton, "Delete the app currently selected in the dropdown from Intune") }
if ($removeOldAppsButton)   { $toolTip.SetToolTip($removeOldAppsButton,   "Delete all superseded app versions from Intune at once") }

# tabDiscovered
if ($deployDiscoveredButton){ $toolTip.SetToolTip($deployDiscoveredButton,"Deploy the checked discovered apps to Microsoft Intune") }
if ($exportDiscoveredCsvButton){ $toolTip.SetToolTip($exportDiscoveredCsvButton,"Export discovered apps with Winget IDs to a CSV file") }
if ($checkAllDiscoveredButton)  { $toolTip.SetToolTip($checkAllDiscoveredButton,   "Check all apps in the discovered apps list") }
if ($uncheckAllDiscoveredButton){ $toolTip.SetToolTip($uncheckAllDiscoveredButton, "Uncheck all apps in the discovered apps list") }

# tabSettings
if ($browsePathButton)         { $toolTip.SetToolTip($browsePathButton,         "Open a folder browser to choose the default package folder") }
if ($autoCheckUpdatesCheckbox) { $toolTip.SetToolTip($autoCheckUpdatesCheckbox, "Automatically scan for app updates each time you log in") }
if ($rememberMeCheckbox)       { $toolTip.SetToolTip($rememberMeCheckbox,       "Save your username so it is pre-filled on the next launch") }
if ($saveSettingsButton)       { $toolTip.SetToolTip($saveSettingsButton,       "Save all settings to disk") }
if ($clearCacheButton)         { $toolTip.SetToolTip($clearCacheButton,         "Clear WinGet version cache, Discovery search cache, and Graph detected-apps cache") }
if ($checkUpdateButton)        { $toolTip.SetToolTip($checkUpdateButton,        "Check GitHub for a newer version of WinTuner GUI") }

# Run the form mit finalem Sicherheitsnetz
try {
    [System.Windows.Forms.Application]::Run($form)
} catch {
    # Fängt ab, falls das Skript als Ganzes unerwartet beendet wird
    Write-FileLog "FATAL SCRIPT CRASH: $($_.Exception.Message)`n$($_.ScriptStackTrace)"
} finally {
    $activeScan = $script:updateScanContext
    if ($activeScan) {
        try {
            [System.IO.File]::WriteAllText(
                $activeScan.CancelPath,
                (Get-Date).ToString('O'),
                [System.Text.UTF8Encoding]::new($false)
            )
        } catch {}
        try { $activeScan.Timer.Stop() } catch {}
        try { $activeScan.PowerShell.Stop() } catch {}
        try { $activeScan.PowerShell.Dispose() } catch {}
        try { $activeScan.Timer.Dispose() } catch {}
        Remove-Item -LiteralPath $activeScan.ProgressPath, $activeScan.CancelPath -Force -ErrorAction SilentlyContinue
        $script:updateScanContext = $null
    }

    $activeUpdateDeployment = $script:updateDeploymentContext
    if ($activeUpdateDeployment) {
        try { $activeUpdateDeployment.Timer.Stop() } catch {}
        try { $activeUpdateDeployment.PowerShell.Stop() } catch {}
        try { $activeUpdateDeployment.PowerShell.Dispose() } catch {}
        try { $activeUpdateDeployment.Timer.Dispose() } catch {}
        Remove-Item -LiteralPath $activeUpdateDeployment.ProgressPath -Force -ErrorAction SilentlyContinue
        $script:updateDeploymentContext = $null
    }
    $activeSupersededSearch = $script:supersededSearchContext
    if ($activeSupersededSearch) {
        try { $activeSupersededSearch.Timer.Stop() } catch {}
        try { $activeSupersededSearch.PowerShell.Stop() } catch {}
        try { $activeSupersededSearch.PowerShell.Dispose() } catch {}
        try { $activeSupersededSearch.Timer.Dispose() } catch {}
        $script:supersededSearchContext = $null
    }

    $activePackageSearch = $script:packageSearchContext
    if ($activePackageSearch) {
        try { $activePackageSearch.Timer.Stop() } catch {}
        try { $activePackageSearch.PowerShell.Stop() } catch {}
        try { $activePackageSearch.PowerShell.Dispose() } catch {}
        try { $activePackageSearch.Timer.Dispose() } catch {}
        $script:packageSearchContext = $null
    }

    $activeVersionLookup = $script:versionLookupContext
    if ($activeVersionLookup) {
        try { $activeVersionLookup.Timer.Stop() } catch {}
        try { $activeVersionLookup.PowerShell.Stop() } catch {}
        try { $activeVersionLookup.PowerShell.Dispose() } catch {}
        try { $activeVersionLookup.Timer.Dispose() } catch {}
        $script:versionLookupContext = $null
    }

    $activePackageBuild = $script:packageBuildContext
    if ($activePackageBuild) {
        try { $activePackageBuild.Timer.Stop() } catch {}
        try { $activePackageBuild.PowerShell.Stop() } catch {}
        try { $activePackageBuild.PowerShell.Dispose() } catch {}
        try { $activePackageBuild.Timer.Dispose() } catch {}
        $script:packageBuildContext = $null
    }

    $activePackageUpload = $script:packageUploadContext
    if ($activePackageUpload) {
        try { $activePackageUpload.Timer.Stop() } catch {}
        try { $activePackageUpload.PowerShell.Stop() } catch {}
        try { $activePackageUpload.PowerShell.Dispose() } catch {}
        try { $activePackageUpload.Timer.Dispose() } catch {}
        $script:packageUploadContext = $null
    }
}
