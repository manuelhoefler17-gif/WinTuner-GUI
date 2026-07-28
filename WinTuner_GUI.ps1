# WinTuner GUI by Manuel Höfler
# v0.10.12 – Hotfix: Update-Check – Robuste GitHub API Fehlerbehandlung + Debug-Logging
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

# --- Application metadata ---
$script:appVersion  = "0.10.12"
$script:githubRepo  = "manuelhoefler17-gif/WinTuner-GUI"
$script:githubApiUrl = "https://api.github.com/repos/manuelhoefler17-gif/WinTuner-GUI/releases/latest"

# Version comparison helper: returns $true if Latest > Current
function Test-IsNewerVersion {
    param([string]$Latest, [string]$Current)
    if (-not $Latest -or -not $Current) { return $false }
    try {
        return ([version]$Latest -gt [version]$Current)
    } catch {
        $mL = [regex]::Match($Latest, '^\s*(\d+(?:\.\d+){0,3})')
        $mC = [regex]::Match($Current, '^\s*(\d+(?:\.\d+){0,3})')
        if (-not $mL.Success -or -not $mC.Success) { return $false }
        $vL = $mL.Groups[1].Value
        $vC = $mC.Groups[1].Value
        try { return ([version]$vL -gt [version]$vC) } catch {
            $numsL = $vL.Split('.') | ForEach-Object {[int]$_}
            $numsC = $vC.Split('.') | ForEach-Object {[int]$_}
            $len = [Math]::Max($numsL.Count, $numsC.Count)
            for ($i=0; $i -lt $len; $i++) {
                $a = if ($i -lt $numsL.Count) { $numsL[$i] } else { 0 }
                $b = if ($i -lt $numsC.Count) { $numsC[$i] } else { 0 }
                if     ($a -gt $b) { return $true }
                elseif ($a -lt $b) { return $false }
            }
            return $false
        }
    }
}


function Test-AppUpdateAvailable {
  <#
  .SYNOPSIS
    Checks GitHub for a newer release of WinTuner GUI
  .OUTPUTS
    PSCustomObject with properties: UpdateAvailable, LatestVersion, DownloadUrl, ReleaseUrl, ReleaseNotes, ErrorMessage
  #>
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
    Write-Log "Checking for app updates from GitHub..."

    $headers = @{
      'Accept'     = 'application/vnd.github.v3+json'
      'User-Agent' = 'WinTuner-GUI-UpdateCheck'
    }

    $savedDefaults = $PSDefaultParameterValues.Clone()
    try {
      $PSDefaultParameterValues = @{}
      $response = Invoke-RestMethod -Uri $script:githubApiUrl -Headers $headers -TimeoutSec 10 -ErrorAction Stop
    } finally {
      $PSDefaultParameterValues = $savedDefaults
    }

    # === ROBUSTE VERSION-EXTRAKTION ===
    # 1. Prüfe ob $response existiert und Properties hat
    if (-not $response) {
      throw "GitHub API returned empty response"
    }

    Write-Log "DEBUG: GitHub API response type: $($response.GetType().Name)"
    
    # 2. Extrahiere tag_name mit mehreren Fallbacks
    $remoteTag = $null
    
    # Versuch 1: Direkter Zugriff auf tag_name Property
    if ($response.PSObject.Properties['tag_name']) {
      $remoteTag = $response.tag_name
    }
    # Versuch 2: Wenn response ist ein Hashtable
    elseif ($response -is [hashtable] -and $response['tag_name']) {
      $remoteTag = $response['tag_name']
    }
    # Versuch 3: Konvertiere zu JSON und zurück (Recovery-Fallback)
    else {
      $responseJson = $response | ConvertTo-Json
      $parsedResponse = $responseJson | ConvertFrom-Json
      if ($parsedResponse.tag_name) {
        $remoteTag = $parsedResponse.tag_name
      }
    }

    if (-not $remoteTag) {
      throw "Could not extract tag_name from GitHub response. Response keys: $($response.PSObject.Properties.Name -join ', ')"
    }

    Write-Log "DEBUG: Extracted tag_name from GitHub: '$remoteTag'"

    # 3. Bereinige Version: entferne "v" Präfix und Suffix (z.B. "-Beta", "-RC1")
    $cleanVersion = $remoteTag
    $cleanVersion = $cleanVersion -replace '^v', ''  # Remove leading "v"
    $cleanVersion = $cleanVersion -replace '-.*$', ''  # Remove "-Beta", "-RC1" etc.
    $cleanVersion = $cleanVersion.Trim()

    if (-not $cleanVersion) {
      throw "Version string is empty after cleaning. Original tag: '$remoteTag'"
    }

    Write-Log "DEBUG: Cleaned version: '$cleanVersion'"

    $result.LatestVersion = $cleanVersion
    
    # 4. Extrahiere weitere Metadaten
    if ($response.PSObject.Properties['html_url']) {
      $result.ReleaseUrl = $response.html_url
    }
    if ($response.PSObject.Properties['body']) {
      $result.ReleaseNotes = $response.body
    }

    # 5. Finde .ps1 Download-Asset
    if ($response.PSObject.Properties['assets'] -and $response.assets) {
      $ps1Asset = $response.assets | Where-Object { $_.name -like '*.ps1' } | Select-Object -First 1
      if ($ps1Asset) {
        $result.DownloadUrl = $ps1Asset.browser_download_url
        Write-Log "DEBUG: Found .ps1 asset: $($ps1Asset.name)"
      }

      # 6. Finde optionales SHA256 Checksum-Asset
      $shaAsset = $response.assets | Where-Object { $_.name -like '*.sha256' } | Select-Object -First 1
      if ($shaAsset) {
        $result.HashUrl = $shaAsset.browser_download_url
        Write-Log "DEBUG: Found .sha256 asset: $($shaAsset.name)"
      }
    }

    # 7. Vergleiche Versionen
    if (Test-IsNewerVersion -Latest $cleanVersion -Current $script:appVersion) {
      $result.UpdateAvailable = $true
      Write-Log "Update available: $($script:appVersion) -> $cleanVersion"
    } else {
      Write-Log "App is up to date (v$($script:appVersion), latest: v$cleanVersion)"
    }

  } catch {
    $result.ErrorMessage = $_.Exception.Message
    Write-Log "Update check failed: $($_.Exception.Message)"
    Write-Log "DEBUG: Stack trace: $($_.ScriptStackTrace)"
  }

  return $result
}

function Invoke-AppSelfUpdate {
  param(
    [Parameter(Mandatory=$true)]
    [string]$DownloadUrl,
    [string]$HashUrl = $null
  )

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

    Write-Log "Downloading update from: $DownloadUrl"
    Update-Status "Downloading update..."

    $tempFile = [System.IO.Path]::GetTempFileName() + ".ps1"

    # Temporarily clear PSDefaultParameterValues to prevent parameter binding conflicts
    # (wildcard entries like '*:ProgressAction' can corrupt URI resolution in some PS7 builds)
    $savedDefaults = $PSDefaultParameterValues.Clone()
    try {
      $PSDefaultParameterValues = @{}
      $headers = @{ 'User-Agent' = 'WinTuner-GUI-UpdateCheck' }
      Invoke-WebRequest -Uri $DownloadUrl -OutFile $tempFile -Headers $headers -TimeoutSec 60 -UseBasicParsing -ErrorAction Stop
    } finally {
      $PSDefaultParameterValues = $savedDefaults
    }

    # Validate download
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

    # SHA256 integrity check (optional – skipped if no hash URL provided)
    if ($HashUrl) {
      $hashMismatch = $false
      try {
        Write-Log "Verifying SHA256 integrity..."
        $savedDefaults2 = $PSDefaultParameterValues.Clone()
        try {
          $PSDefaultParameterValues = @{}
          $expectedHash = (Invoke-RestMethod -Uri $HashUrl -TimeoutSec 15 -ErrorAction Stop).Trim().ToUpper()
        } finally {
          $PSDefaultParameterValues = $savedDefaults2
        }
        # Hash file may contain "HASH filename" or just "HASH"
        $expectedHash = ($expectedHash -split '\s+')[0].ToUpper()
        $actualHash = (Get-FileHash $tempFile -Algorithm SHA256).Hash.ToUpper()
        if ($actualHash -ne $expectedHash) {
          $hashMismatch = $true
          throw "SHA256 mismatch: download may be corrupt or tampered! Expected: $expectedHash, Got: $actualHash"
        }
        Write-Log "SHA256 verified OK: $actualHash"
      } catch {
        # Re-throw only real hash mismatches, not network errors
        if ($hashMismatch) { throw }
        Write-Log "Warning: SHA256 check skipped (could not fetch hash): $($_.Exception.Message)"
      }
    }

    Write-Log "Download complete ($fileSize bytes). Replacing script..."

    # Create backup
    $backupPath = "$currentPath.backup"
    try {
      Copy-Item -Path $currentPath -Destination $backupPath -Force -ErrorAction Stop
      Write-Log "Backup created: $backupPath"
    } catch {
      Write-Log "Warning: Could not create backup: $($_.Exception.Message)"
    }

    # Replace current script
    Move-Item -Path $tempFile -Destination $currentPath -Force -ErrorAction Stop

    Write-Log "Script replaced successfully. Restart required."
    return $true

  } catch {
    Write-Log "Self-update failed: $($_.Exception.Message)"
    if ($tempFile -and (Test-Path $tempFile)) {
      Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
    }
    [System.Windows.Forms.MessageBox]::Show(
      "Update failed: $($_.Exception.Message)`n`nYou can update manually from:`nhttps://github.com/$($script:githubRepo)/releases/latest",
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
    Update-Status "Update check failed (v$($script:appVersion)) – check internet connection"
    return
  }

  if ($UpdateResult -and $UpdateResult.UpdateAvailable) {
    Update-Status "Update available: v$($UpdateResult.LatestVersion)"
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
          Update-Status "Downloading update..."
          [System.Windows.Forms.Application]::DoEvents()

          $success = Invoke-AppSelfUpdate -DownloadUrl $UpdateResult.DownloadUrl -HashUrl $UpdateResult.HashUrl

          if ($success) {
            Update-Status "Update installed successfully. Please restart WinTuner GUI."
            $restartMsg  = "Update installed successfully!`n`n"
            $restartMsg += "WinTuner GUI needs to restart to apply the update.`n"
            $restartMsg += "Click OK to close. Please start the script again manually."

            [System.Windows.Forms.MessageBox]::Show(
              $restartMsg,
              "Update Complete",
              [System.Windows.Forms.MessageBoxButtons]::OK,
              [System.Windows.Forms.MessageBoxIcon]::Information
            )

            $form.Close()
          } else {
            Update-Status "Update download/install failed. See log for details."
          }
        } else {
          if ($isManual) {
            Update-Status "Update postponed by user"
          } else {
            Update-Status "Update available: v$($UpdateResult.LatestVersion) - Go to Settings to update later."
          }
        }
      } else {
        Update-Status "Update available: v$($UpdateResult.LatestVersion) (manual download required)"
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
      Update-Status "Update check completed (dialog error). See log for details."
    }
    return
  }

  $latestVer = if ($UpdateResult -and $UpdateResult.LatestVersion) { $UpdateResult.LatestVersion } else { "unknown" }
  $statusMsg = "Up to date – Local: v$($script:appVersion) | GitHub: v$latestVer"
  Update-Status $statusMsg

  if ($isManual) {
    [System.Windows.Forms.MessageBox]::Show(
      "WinTuner GUI is up to date.`n`nLocal version:  v$($script:appVersion)`nGitHub version: v$latestVer",
      "No Update Available",
      [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Information
    )
  }
}
