# WinTuner GUI by Manuel Höfler
# v0.10.10 – Hotfix: Invoke-AsyncOperation Closure-Bug (OnComplete nie ausgeführt) + Update-Check Feedback
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
$script:appVersion  = "0.10.10"
$script:githubRepo  = "manuelhoefler17-gif/WinTuner-GUI"
$script:githubApiUrl = "https://api.github.com/repos/manuelhoefler17-gif/WinTuner-GUI/releases/latest"

# --- Runtime state (set during execution) ---
# $script:isConnected      – whether the user is logged in to a tenant
# $script:currentUserUpn   – UPN of the currently logged-in user
# $script:builtVersions    – tracks effective built package versions per PackageId
# $script:wingetVersionCache – in-memory cache for winget version lookups
# $script:versionCachePath – path to the on-disk version cache JSON file
# $script:isDarkMode       – current theme state (true = dark)
# $script:currentTheme     – active theme hashtable (darkTheme or lightTheme)
# $script:asyncResult      – last result from Invoke-AsyncOperation
# $script:diskCache        – in-memory copy of the on-disk version cache (loaded once)
# $script:diskCacheLoaded  – whether $script:diskCache has been populated from disk

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

    # Extract version from tag_name (strip leading "v" and any suffix like "-Beta")
    $remoteTag = $response.tag_name
    $remoteVersionStr = $remoteTag -replace '^v', ''
    $cleanVersion = $remoteVersionStr -replace '-.*$', ''  # Remove "-Beta", "-RC1" etc.

    $result.LatestVersion = $cleanVersion
    $result.ReleaseUrl    = $response.html_url
    $result.ReleaseNotes  = $response.body

    # Find the .ps1 download asset
    $ps1Asset = $response.assets | Where-Object { $_.name -like '*.ps1' } | Select-Object -First 1
    if ($ps1Asset) {
      $result.DownloadUrl = $ps1Asset.browser_download_url
    }

    # Find optional SHA256 checksum asset
    $shaAsset = $response.assets | Where-Object { $_.name -like '*.sha256' } | Select-Object -First 1
    if ($shaAsset) {
      $result.HashUrl = $shaAsset.browser_download_url
    }

    # Compare versions using existing function
    if (Test-IsNewerVersion -Latest $cleanVersion -Current $script:appVersion) {
      $result.UpdateAvailable = $true
      Write-Log "Update available: $($script:appVersion) -> $cleanVersion"
    } else {
      Write-Log "App is up to date (v$($script:appVersion), latest: v$cleanVersion)"
    }

  } catch {
    $result.ErrorMessage = $_.Exception.Message
    Write-Log "Update check failed: $($_.Exception.Message)"
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

function Get-VersionDiskCache {
  if (-not $script:versionCachePath) { return @{} }
  try {
    if (Test-Path $script:versionCachePath) {
      $raw = Get-Content $script:versionCachePath -Raw -Encoding utf8 -ErrorAction Stop
      $parsed = $raw | ConvertFrom-Json -ErrorAction Stop
      $ht = @{}
      foreach ($prop in $parsed.PSObject.Properties) {
        $ht[$prop.Name] = @{
          versions  = @($prop.Value.versions)
          timestamp = [datetime]::Parse($prop.Value.timestamp, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind)
        }
      }
      return $ht
    }
  } catch {
    Write-Log "Warning: Could not read version cache: $($_.Exception.Message)"
  }
  return @{}
}

function Save-VersionDiskCache {
  param([hashtable]$Cache)
  if (-not $script:versionCachePath) { return }
  try {
    $obj = @{}
    foreach ($key in $Cache.Keys) {
      $obj[$key] = @{
        versions  = $Cache[$key].versions
        timestamp = $Cache[$key].timestamp.ToString('o')
      }
    }
    $obj | ConvertTo-Json -Depth 4 | Set-Content -Path $script:versionCachePath -Encoding utf8 -ErrorAction SilentlyContinue
  } catch {
    Write-Log "Warning: Could not save version cache: $($_.Exception.Message)"
  }
}

function Get-WingetVersions {
  param([string]$PackageId)

  # 1) RAM cache
  if ($script:wingetVersionCache.ContainsKey($PackageId)) {
    return $script:wingetVersionCache[$PackageId]
  }

  # 2) Disk cache (TTL 6h) – loaded once per session
  if (-not $script:diskCacheLoaded) {
    $script:diskCache = Get-VersionDiskCache
    $script:diskCacheLoaded = $true
  }
  if ($script:diskCache.ContainsKey($PackageId)) {
    $entry = $script:diskCache[$PackageId]
    $ageHours = ([datetime]::UtcNow - $entry.timestamp.ToUniversalTime()).TotalHours
    if ($ageHours -lt 6 -and $entry.versions -and $entry.versions.Count -gt 0) {
      $script:wingetVersionCache[$PackageId] = $entry.versions
      Write-Log "Version cache hit (disk) for $PackageId (age: $([math]::Round($ageHours,1))h)"
      return $entry.versions
    }
  }

  # 3) Query winget
  try { $output = & winget show --id $PackageId --versions 2>$null } catch { return @() }
  if (-not $output) { return @() }

  $cand = @()
  foreach ($line in @($output)) {
    $t = ($line -replace '^[\s\-•]+','').Trim()
    if (-not $t) { continue }
    if ($t -match '^(\d+)(\.[0-9A-Za-z]+)*([\-+._][0-9A-Za-z]+)*$') { $cand += $t }
  }

  $unique = @($cand | Select-Object -Unique)
  $parsed = foreach ($v in $unique) {
    $ok = $false; $vo = $null
    try { $vo = [version]$v; $ok = $true } catch {}
    [pscustomobject]@{ Text = $v; Parsed = $vo; Numeric = $ok }
  }

  $result = @()
  if ($parsed | Where-Object Numeric) {
    $result = @($parsed | Where-Object Numeric | Sort-Object Parsed -Descending | Select-Object -ExpandProperty Text)
  } else {
    $result = @($parsed | Sort-Object Text -Descending | Select-Object -ExpandProperty Text)
  }

  # 4) Store in RAM cache
  $script:wingetVersionCache[$PackageId] = $result

  # 5) Store in disk cache (update script-level cache variable and persist to disk)
  $script:diskCache[$PackageId] = @{
    versions  = $result
    timestamp = [datetime]::UtcNow
  }
  Save-VersionDiskCache -Cache $script:diskCache

  return $result
}

function Get-PreviousWingetVersion {
  param([string]$PackageId, [string]$LatestVersion)

  $allVersions = @(Get-WingetVersions -PackageId $PackageId)
  if (-not $allVersions -or $allVersions.Count -eq 0) { return $null }

  $candidates = @($allVersions | Where-Object { $_ -ne $LatestVersion })
  if ($candidates.Count -gt 0) { return $candidates[0] }
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

# Performs update workflow for a single app (create package + deploy)
function Update-SingleApp {
  param(
    [Parameter(Mandatory=$true)]
    [string]$AppName,
    [Parameter(Mandatory=$false)]
    [string]$CurrentVersion,
    [Parameter(Mandatory=$false)]
    [string]$LatestVersion,
    [Parameter(Mandatory=$false)]
    [string]$GraphId,
    [Parameter(Mandatory=$false)]
    [string]$PackageIdentifier,
    [Parameter(Mandatory=$true)]
    [string]$RootPackageFolder,
    [switch]$AllowUserRetry
  )
  
  $result = @{
    Success = $false
    Message = ""
    EffectiveVersion = $null
  }
  
  try {
    Write-Log "Starting update for: $AppName (Current: $CurrentVersion, Latest: $LatestVersion)"
    
    # 1) Resolve Winget ID - use PackageIdentifier if available, otherwise fail
    $wingetId = $PackageIdentifier
    
    Write-Log ("Resolved winget id for {0}: {1}" -f $AppName, ($wingetId ? $wingetId : '<none>'))
    
    if ([string]::IsNullOrWhiteSpace($wingetId)) {
      $result.Message = "Cannot determine PackageId for '$AppName'"
      Write-Log $result.Message
      return $result
    }
    
    # 2) Create/refresh package using fallback logic
    Write-Log "Creating package for $AppName..."
    # For updates, ALWAYS use LatestVersion (ignore cached selectedPackageVersions)
    $desired = $LatestVersion
    Write-Log "Update workflow: forcing LatestVersion $desired (ignoring any cached selection)"
    
    $resPkg = New-WingetPackageWithFallback `
      -PackageId $wingetId `
      -PackageFolder $RootPackageFolder `
      -DesiredVersion $desired `
      -LatestVersion $LatestVersion `
      -InstalledVersion $CurrentVersion `
      -AllowUserRetry:$AllowUserRetry `
      -ErrorAction SilentlyContinue
    
    if (-not $resPkg -or -not $resPkg.Succeeded) {
      $errDetail = if ($resPkg -and $resPkg.ErrorMessage) { ": $($resPkg.ErrorMessage)" } else { "" }
      $result.Message = "Package creation failed for $AppName$errDetail"
      Write-Log $result.Message
      return $result
    }
    
    $effectiveVersion = if ($resPkg.EffectiveVersion) { $resPkg.EffectiveVersion } else { $LatestVersion }
    $result.EffectiveVersion = $effectiveVersion
    
    # 3) Deploy with best available identifier
    Write-Log "Deploying $AppName version $effectiveVersion..."
    $deploySplat = @{ 
      RootPackageFolder = $RootPackageFolder
      ErrorAction = 'Stop' 
    }
    
    if ($GraphId) {
      $deploySplat.GraphId = $GraphId
      $deploySplat.KeepAssignments = $true
      $deploySplat.PackageId = $wingetId
      $deploySplat.Version = $effectiveVersion
      Write-Log "Deploying by GraphId ($GraphId) + PackageId/Version"
    } else {
      $deploySplat.PackageId = $wingetId
      $deploySplat.Version = $effectiveVersion
      Write-Log "Deploying by PackageId ($wingetId) version $effectiveVersion"
    }
    
    Deploy-WtWin32App @deploySplat
    
    $result.Success = $true
    $result.Message = "Update completed successfully for $AppName"
    Write-Log $result.Message
    
  } catch {
    $result.Message = "Update failed for ${AppName}: $($_.Exception.Message)"
    Write-Log $result.Message
  }
  
  return $result
}

function Invoke-AppUpdateBatch {
  param(
    [Parameter(Mandatory=$true)]
    [object[]]$Apps,
    [Parameter(Mandatory=$true)]
    [string]$RootPackageFolder
  )

  $updateSelectedButton.Enabled = $false
  $updateAllButton.Enabled = $false
  $updateSearchButton.Enabled = $false
  $checkAllButton.Enabled = $false
  $uncheckAllButton.Enabled = $false

  $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
  $progressBar.MarqueeAnimationSpeed = 30
  $progressBar.Visible = $true

  $successCount = 0
  $failedCount = 0
  $totalCount = $Apps.Count
  $currentIndex = 0
  $failedList = [System.Collections.Generic.List[object]]::new()

  try {
    foreach ($app in $Apps) {
      $currentIndex++
      Update-Status ("Updating ({0}/{1}): {2}" -f $currentIndex, $totalCount, $app.Name)
      [System.Windows.Forms.Application]::DoEvents()  # TODO: refactor to use Invoke-AsyncOperation

      # Extract properties from WtWin32App object
      $appName = $app.Name
      $appCurrentVersion = $app.CurrentVersion
      $appLatestVersion = $app.LatestVersion
      $appGraphId = $app.GraphId

      # Get PackageIdentifier - use Resolve-WingetIdForApp
      $appPackageId = Resolve-WingetIdForApp -App $app

      Write-Log "Calling Update-SingleApp with: Name='$appName', Current='$appCurrentVersion', Latest='$appLatestVersion', GraphId='$appGraphId', PackageId='$appPackageId'"

      $result = Update-SingleApp `
        -AppName $appName `
        -CurrentVersion $appCurrentVersion `
        -LatestVersion $appLatestVersion `
        -GraphId $appGraphId `
        -PackageIdentifier $appPackageId `
        -RootPackageFolder $RootPackageFolder

      if ($result.Success) {
        $successCount++
        Write-Log "Successfully updated: $appName"

        # Immediately remove the updated app from the UI list
        $idxToRemove = $updateListBox.Items.IndexOf($appName)
        if ($idxToRemove -ge 0) { $updateListBox.Items.RemoveAt($idxToRemove) }
        $toRemove = $script:updateApps | Where-Object { $_.Name -eq $appName }
        foreach ($item in @($toRemove)) { [void]$script:updateApps.Remove($item) }
        [System.Windows.Forms.Application]::DoEvents()
      } else {
        $failedCount++
        Write-Log "Failed to update: $appName - $($result.Message)"
        $failedList.Add([pscustomobject]@{ Name = $appName; Reason = $result.Message })
      }
    }

    if ($failedList.Count -gt 0) {
      $summary = "The following $($failedList.Count) app(s) could not be updated:`n`n"
      $summary += ($failedList | ForEach-Object { "• $($_.Name): $($_.Reason)" }) -join "`n"
      [System.Windows.Forms.MessageBox]::Show(
        $summary,
        "Update Summary – $($failedList.Count) Failed",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
      )
    }

    # Refresh update list
    try { $updateSearchButton.PerformClick() } catch {}

    return @{ SuccessCount = $successCount; FailedList = $failedList }
  } finally {
    $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $progressBar.Visible = $false
    $progressBar.Value = 0
    $updateSelectedButton.Enabled = $true
    $updateAllButton.Enabled = $true
    $updateSearchButton.Enabled = $true
    $checkAllButton.Enabled = $true
    $uncheckAllButton.Enabled = $true
  }
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
function Write-Log {
  param([string]$message)
  if ([string]::IsNullOrWhiteSpace($message)) { return }
  
  try {
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logLine = "$timestamp - $message"
    
    # Write to file
    try {
      $base = if ($PSScriptRoot) { $PSScriptRoot } elseif ($MyInvocation -and $MyInvocation.MyCommand -and $MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path } else { (Get-Location).Path }
      if ([string]::IsNullOrWhiteSpace($base)) { $base = [Environment]::GetFolderPath('LocalApplicationData') }
      if (-not (Test-Path $base)) { 
          try { New-Item -ItemType Directory -Path $base -Force | Out-Null } catch { return }
      }
      $logPath = Join-Path $base 'WinTuner_GUI.log'
      
      # --- Log rotation: limit log file size ---
      $maxLogSize = 2MB # Maximum log file size before rotation
      if (Test-Path $logPath) {
          $logFile = Get-Item $logPath
          if ($logFile.Length -gt $maxLogSize) {
              $oldLogPath = Join-Path $base 'WinTuner_GUI_old.log'
              # Move current log to backup (overwrites existing backup)
              Move-Item -Path $logPath -Destination $oldLogPath -Force -ErrorAction SilentlyContinue
          }
      }
      # --- End log rotation ---

      Add-Content -Path $logPath -Value $logLine -Encoding utf8 -ErrorAction SilentlyContinue
    } catch {
      # Silently ignore file write errors
    }
    
    # Update UI - always try to append (suppress any errors)
    if ($outputBox) {
      try {
        if ($outputBox.InvokeRequired) {
          # Cross-thread call - use Invoke
          $outputBox.Invoke([Action]{
            $outputBox.AppendText("$logLine`r`n")
          })
        } else {
          # Same thread - direct call
          $outputBox.AppendText("$logLine`r`n")
        }
      } catch {
        # Silently ignore UI update errors (threading issues)
      }
    }
  } catch {
    # Completely suppress all logging errors to prevent crashes
  }
}

# Status update function
function Update-Status {
  param([string]$status)
  $statusLabel.Text = $status
  Write-Log $status
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
  
  # Update UI - show progress in marquee style (indefinite)
  Update-Status $StatusText
  $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
  $progressBar.MarqueeAnimationSpeed = 30
  $progressBar.Visible = $true
  
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
  
  # Do work in background
  $doWork = {
    param($sender, $e)
    try {
      $e.Result = & $ScriptBlock
    } catch {
      $errMsg = $_.Exception.Message
      $e.Result = @{ Error = $errMsg }
      # Write directly to log file (thread-safe, no UI access)
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
  $runCompleted = {
    param($sender, $e)
    
    # Restore progress bar to normal
	$progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
	$progressBar.Maximum = 100   # ← Zeile hinzufügen
	$progressBar.Value = 100
    
    # Re-enable controls
    foreach ($ctrl in $DisableControls) {
      if ($ctrl) { $ctrl.Enabled = $true }
    }
    
    # Execute completion callback with result
    if ($OnComplete) {
      try {
        & $OnComplete $e.Result
      } catch {
        Write-Log "Async completion callback error: $($_.Exception.Message)"
        Update-Status "Operation completed with errors"
      }
    }
    
    # Hide progress after short delay
    $hideTimer = New-Object System.Windows.Forms.Timer
    $hideTimer.Interval = 1000
    $hideTimer.Add_Tick({
      param($sender, $e)
      $progressBar.Visible = $false
      $progressBar.Value = 0
      $sender.Stop()
      $sender.Dispose()
    })
    $hideTimer.Start()
    
    # Dispose the BackgroundWorker to prevent memory leaks
    $sender.Dispose()
  }.GetNewClosure()
  $bw.Add_RunWorkerCompleted($runCompleted)
  
  # Start async operation
  $bw.RunWorkerAsync()
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
  Save-Settings
}

# Clears the recent users list and resets LastUser
function Clear-RecentUsers {
  $script:settings.RecentUsers = @()
  $script:settings.LastUser = ""
  Save-Settings
}

# Helper: check if WinTuner is connected (simple smoke test)
function Test-WtConnected {
  try {
    # Avoid Select-Object -First 1 to prevent WinForms pipeline crash during login
    $apps = Get-WtWin32Apps -Update:$false -Superseded:$false -ErrorAction Stop
    foreach ($app in $apps) {
        return $true # Exit safely on first found element
    }
    return $true # No apps found but no error either
  } catch { return $false }
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
  if ($updateSearchButton) { $updateSearchButton.Enabled = $Connected }
  if ($scanDiscoveredButton) { $scanDiscoveredButton.Enabled = $Connected }
  if ($updateSelectedButton) { $updateSelectedButton.Enabled = $Connected }
  if ($updateAllButton) { $updateAllButton.Enabled = $Connected }
  if ($supersededSearchButton) { $supersededSearchButton.Enabled = $Connected }
  if ($deleteSelectedAppButton) { $deleteSelectedAppButton.Enabled = $Connected }
  if ($removeOldAppsButton) { $removeOldAppsButton.Enabled = $Connected }
  
  if ($loginInfoLabel) {
    $loginInfoLabel.Visible = $Connected
    if ($Connected -and $script:currentUserUpn) { $loginInfoLabel.Text = "Logged in as: $($script:currentUserUpn)" }
  }
}

$script:isConnected = $false
$script:currentUserUpn = ""

# Track effective built versions per PackageId
$script:builtVersions = @{}
# Cache for winget version lookups (speeds up repeated searches)
$script:wingetVersionCache = @{}
$script:versionCachePath = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'WinTuner_VersionCache.json'
# Disk cache loaded once at first use (Fix 1)
$script:diskCache = @{}
$script:diskCacheLoaded = $false

# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = "WinTuner GUI"
$form.Size = New-Object System.Drawing.Size(960, 850)
$form.Padding = '5,5,5,5'

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
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = ""
$statusLabel.Location = New-Object System.Drawing.Point(10, 745)
$statusLabel.Width = 750
$statusLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($statusLabel)

# Output textbox (Log area below tabs and progress bar)
$outputBox = New-Object System.Windows.Forms.TextBox
$outputBox.Location = New-Object System.Drawing.Point(10, 620)
$outputBox.Size = New-Object System.Drawing.Size(760, 120)
$outputBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$outputBox.Multiline = $true
$outputBox.ScrollBars = "Vertical"
$outputBox.ReadOnly = $true
$form.Controls.Add($outputBox)

# Progress bar (appears between tabs and log when active)
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 595)
$progressBar.Width = 760
$progressBar.Height = 20
$progressBar.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$progressBar.Visible = $false
$form.Controls.Add($progressBar)

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
$tabControl.Size = New-Object System.Drawing.Size(760, 500)
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
$tabCreate.Controls.Add($dropdown)

$versionsButton = New-Object System.Windows.Forms.Button
$versionsButton.Text = "Versions..."
$versionsButton.Location = New-Object System.Drawing.Point(570,60)
$versionsButton.Width = 180
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
$tabCreate.Controls.Add($createButton)

$uploadButton = New-Object System.Windows.Forms.Button
$uploadButton.Text = "Upload to Tenant"
$uploadButton.Location = New-Object System.Drawing.Point(290,140)
$uploadButton.Width = 180
$uploadButton.Visible = $true
$uploadButton.Enabled = $false
$tabCreate.Controls.Add($uploadButton)

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
$updateFilterBox.Width = 395
$updateFilterBox.PlaceholderText = "Type to filter apps..."
$tabUpdate.Controls.Add($updateFilterBox)

# CheckedListBox for multi-select updates
$updateListBox = New-Object System.Windows.Forms.CheckedListBox
$updateListBox.Location = New-Object System.Drawing.Point(100,85)
$updateListBox.Width = 650
$updateListBox.Height = 150
$updateListBox.CheckOnClick = $true
$tabUpdate.Controls.Add($updateListBox)

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
$discoveredAppSearchLabel = New-Object System.Windows.Forms.Label
$discoveredAppSearchLabel.Text = "Search App:"
$discoveredAppSearchLabel.Location = New-Object System.Drawing.Point(440, 15)
$discoveredAppSearchLabel.AutoSize = $true
$tabDiscovered.Controls.Add($discoveredAppSearchLabel)

$discoveredAppSearchBox = New-Object System.Windows.Forms.TextBox
$discoveredAppSearchBox.Location = New-Object System.Drawing.Point(540, 12)
$discoveredAppSearchBox.Width = 150
$tabDiscovered.Controls.Add($discoveredAppSearchBox)

$discoveredPublisherLabel = New-Object System.Windows.Forms.Label
$discoveredPublisherLabel.Text = "Publisher:"
$discoveredPublisherLabel.Location = New-Object System.Drawing.Point(440, 42)
$discoveredPublisherLabel.AutoSize = $true
$tabDiscovered.Controls.Add($discoveredPublisherLabel)

$discoveredPublisherBox = New-Object System.Windows.Forms.ComboBox
$discoveredPublisherBox.Location = New-Object System.Drawing.Point(540, 39)
$discoveredPublisherBox.Width = 150
$discoveredPublisherBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
[void]$discoveredPublisherBox.Items.Add("<All Publishers>")
$discoveredPublisherBox.SelectedIndex = 0
$tabDiscovered.Controls.Add($discoveredPublisherBox)

$discoveredSortLabel = New-Object System.Windows.Forms.Label
$discoveredSortLabel.Text = "Sort by:"
$discoveredSortLabel.Location = New-Object System.Drawing.Point(440, 69)
$discoveredSortLabel.AutoSize = $true
$tabDiscovered.Controls.Add($discoveredSortLabel)

$discoveredSortBox = New-Object System.Windows.Forms.ComboBox
$discoveredSortBox.Location = New-Object System.Drawing.Point(540, 66)
$discoveredSortBox.Width = 150
$discoveredSortBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
[void]$discoveredSortBox.Items.Add("Device Count")
[void]$discoveredSortBox.Items.Add("Alphabetical")
$discoveredSortBox.SelectedIndex = 0
$tabDiscovered.Controls.Add($discoveredSortBox)

$discoveredListBox = New-Object System.Windows.Forms.CheckedListBox
$discoveredListBox.Location = New-Object System.Drawing.Point(20,110)
$discoveredListBox.Width = 710
$discoveredListBox.Height = 325
$discoveredListBox.CheckOnClick = $true
$discoveredListBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$tabDiscovered.Controls.Add($discoveredListBox)

$script:discoveredRaw = @()

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
$clearCacheButton.Text = "Clear Version Cache"
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
    
    Save-Settings
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
  $script:wingetVersionCache = @{}
  $script:diskCache = @{}
  $script:diskCacheLoaded = $false
  Remove-Item $script:versionCachePath -Force -ErrorAction SilentlyContinue
  Write-Log "Version cache cleared."
  Update-Status "Version cache cleared."
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
  Invoke-AsyncOperation -StatusText "Checking for updates..." -ScriptBlock {
    Test-AppUpdateAvailable
  } -OnComplete {
    param($updateResult)
    $checkUpdateButton.Enabled = $true
    # Consolidate error checking: Invoke-AsyncOperation wraps thrown exceptions as .Error;
    # Test-AppUpdateAvailable returns graceful errors as .ErrorMessage
    $errorDetail = if ($updateResult -and $updateResult.Error) { $updateResult.Error } `
                   elseif ($updateResult -and $updateResult.ErrorMessage) { $updateResult.ErrorMessage } `
                   else { $null }
    if ($errorDetail) {
      [System.Windows.Forms.MessageBox]::Show(
        "Could not check for updates.`n`nError: $errorDetail`n`nCheck your internet connection and try again.",
        "Update Check Failed",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
      )
      Update-Status "Update check failed (v$($script:appVersion)) – check internet connection"
      return
    }

    if ($updateResult -and $updateResult.UpdateAvailable) {
      $msg  = "A new version of WinTuner GUI is available!`n`n"
      $msg += "Current version: v$($script:appVersion)`n"
      $msg += "Latest version:  v$($updateResult.LatestVersion)`n`n"

      if ($updateResult.DownloadUrl) {
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

          $success = Invoke-AppSelfUpdate -DownloadUrl $updateResult.DownloadUrl -HashUrl $updateResult.HashUrl

          if ($success) {
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
          }
        } else {
          Update-Status "Update postponed by user"
        }
      } else {
        $msg += "No direct download available for this release.`n"
        $msg += "Please download manually from:`n$($updateResult.ReleaseUrl)"

        [System.Windows.Forms.MessageBox]::Show(
          $msg,
          "Update Available",
          [System.Windows.Forms.MessageBoxButtons]::OK,
          [System.Windows.Forms.MessageBoxIcon]::Information
        )
      }
    } else {
      $latestVer = if ($updateResult -and $updateResult.LatestVersion) { $updateResult.LatestVersion } else { "unknown" }
      $msg = "WinTuner GUI is up to date.`n`nLocal version:  v$($script:appVersion)`nGitHub version: v$latestVer"
      Update-Status "Up to date – Local: v$($script:appVersion) | GitHub: v$latestVer"
      [System.Windows.Forms.MessageBox]::Show(
        $msg,
        "No Update Available",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
      )
    }
  }
})

# Hashtable: AppName -> {PackageID, Version}
$script:packageMap = @{}

# Optional: user-chosen versions per PackageID
$script:selectedPackageVersions = @{}

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

$script:settingsPath = Join-Path ([Environment]::GetFolderPath('ApplicationData')) 'WinTunerGUI\settings.json'
$script:settings = @{ 
  RememberMe = $false
  LastUser = ""
  RecentUsers = @()
  MaxRecentUsers = 3
  WingetOverrides = @{}
  DefaultPackagePath = "C:\Temp"
  AutoCheckUpdates = $false
}

function Load-Settings {
  try {
    if (Test-Path $script:settingsPath) {
      $o = Get-Content -Path $script:settingsPath -Raw -ErrorAction Stop | ConvertFrom-Json
      if ($o) {
        $script:settings.RememberMe = [bool]$o.RememberMe
        $script:settings.LastUser = [string]$o.LastUser
        
        if ($o.PSObject.Properties['RecentUsers']) {
            $script:settings.RecentUsers = @([string[]]$o.RecentUsers)
        } else {
            $script:settings.RecentUsers = @()
        }
        if ($o.PSObject.Properties['MaxRecentUsers'] -and $o.MaxRecentUsers -gt 0) {
            $script:settings.MaxRecentUsers = [int]$o.MaxRecentUsers
        } else {
            $script:settings.MaxRecentUsers = 3
        }
        
        # New settings with defaults
        if ($o.PSObject.Properties['DefaultPackagePath']) {
          $script:settings.DefaultPackagePath = [string]$o.DefaultPackagePath
        } else {
          $script:settings.DefaultPackagePath = "C:\Temp"
        }
        
        if ($o.PSObject.Properties['AutoCheckUpdates']) {
          $script:settings.AutoCheckUpdates = [bool]$o.AutoCheckUpdates
        } else {
          $script:settings.AutoCheckUpdates = $false
        }
        
        if ($o.PSObject.Properties['WingetOverrides']) {
          # Convert PSCustomObject to hashtable
          $ht = @{}
          foreach ($p in $o.WingetOverrides.PSObject.Properties) { $ht[$p.Name] = [string]$p.Value }
          $script:settings.WingetOverrides = $ht
        } else { $script:settings.WingetOverrides = @{} }
      }
    }
  } catch {
    Write-Log "Warning: Failed to load settings from $($script:settingsPath): $($_.Exception.Message)"
    # Continue with default settings
  }
}

function Save-Settings {
  try {
    $dir = Split-Path -Parent $script:settingsPath
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    ($script:settings | ConvertTo-Json -Compress) | Set-Content -Path $script:settingsPath -Encoding utf8
  } catch {
    Write-Log "Error: Failed to save settings to $($script:settingsPath): $($_.Exception.Message)"
  }
}

Load-Settings
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
    Save-Settings
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
    if ($loginInfoLabel) { $loginInfoLabel.Text = "Logged in as: $($script:currentUserUpn)" }
    if ($rememberCheckBox) { $script:settings.RememberMe = [bool]$rememberCheckBox.Checked }
    if ($script:settings.RememberMe) { $script:settings.LastUser = $usernameBox.Text } else { $script:settings.LastUser = "" }
    Add-RecentUser -Upn $usernameBox.Text
    # Update dropdown list
    $usernameBox.Items.Clear()
    foreach ($u in @($script:settings.RecentUsers)) {
      if ($u) { [void]$usernameBox.Items.Add($u) }
    }
    Save-Settings
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
  if ([string]::IsNullOrWhiteSpace($appSearchBox.Text)) {
    [void][System.Windows.Forms.MessageBox]::Show(
      "App search can't be empty.",
      "Validation",
      [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Information
    )
    return
  }
  try {
    $searchButton.Enabled = $false
    Update-Status "Searching..."
    $results = Search-WtWinGetPackage -SearchQuery $appSearchBox.Text
    $dropdown.Items.Clear()
    $script:packageMap.Clear()
    foreach ($result in @($results)) {
      $displayText = "$($result.Name) — $($result.PackageID)"
      [void]$dropdown.Items.Add($displayText)
      $script:packageMap[$displayText] = @{
        PackageID = $result.PackageID
        Version   = $result.Version
      }
    }
    if ($dropdown.Items.Count -gt 0) { $dropdown.SelectedIndex = 0 }
    if ($dropdown.Items.Count -eq 0) {
      Update-Status "No results found for '$($appSearchBox.Text)'"
    } else {
      Update-Status "Search completed."
    }
  } finally {
    $searchButton.Enabled = $true
  }
})

$versionsButton.Add_Click({
  if (-not $dropdown.SelectedItem) { Update-Status "Please select a package first."; return }
  $appName  = $dropdown.SelectedItem
  $package  = $script:packageMap[$appName]
  if (-not $package -or -not $package.PackageID) { Update-Status "Selected item is invalid."; return }
  $packageID = $package.PackageID
  $versions = @(Get-WingetVersions -PackageId $packageID)
  if (-not $versions -or $versions.Count -eq 0) { Update-Status "No versions found for the selected package."; return }
  $chosen = Show-VersionPickerDialog -Title ("Select version for {0}" -f $packageID) -Versions $versions
  if ($chosen) {
    $script:selectedPackageVersions[$packageID] = $chosen
    Update-Status ("Selected version for {0}: {1}" -f $packageID, $chosen)
  } else {
    Update-Status "Version selection canceled."
  }
})

$createButton.Add_Click({
  if (-not $dropdown.SelectedItem) { Update-Status "Please select a package."; return }
  $appName  = $dropdown.SelectedItem
  $package  = $script:packageMap[$appName]
  if (-not $package -or -not $package.PackageID) { Update-Status "Selected item is invalid."; return }
  $packageID = $package.PackageID
  $folder    = [System.IO.Path]::GetFullPath($pathBox.Text.Trim())
  $forbiddenPaths = @(
    [Environment]::GetFolderPath('Windows'),
    [Environment]::GetFolderPath('System'),
    "$env:SystemRoot\System32",
    "$env:SystemRoot\SysWOW64",
    "$env:ProgramFiles",
    "${env:ProgramFiles(x86)}"
  )
  $isForbidden = $forbiddenPaths | Where-Object { $folder -eq $_ -or $folder.StartsWith($_ + '\') }
  if ($isForbidden) {
    [System.Windows.Forms.MessageBox]::Show(
      "The selected folder '$folder' is a protected system directory.`nPlease choose a different folder.",
      "Invalid Folder",
      [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    return
  }
  if (-not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
  $filePath  = Join-Path $folder "$packageID.wtpackage"
  
  if (Test-Path $filePath) {
    $res = [System.Windows.Forms.MessageBox]::Show(("A package file already exists:\n{0}\nOverwrite it?" -f $filePath), "Confirm overwrite", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($res -ne [System.Windows.Forms.DialogResult]::Yes) { Update-Status "Creation aborted by user (existing package)."; $uploadButton.Enabled = $true; return }
    try { 
      Remove-Item -Path $filePath -Force -ErrorAction Stop 
    } catch {
      Write-Log "Warning: Failed to delete existing package file ${filePath}: $($_.Exception.Message)"
      Update-Status "Warning: Could not delete existing package, continuing anyway..."
    }
  }

  try {
    $createButton.Enabled = $false
    $searchButton.Enabled = $false
    $versionsButton.Enabled = $false
    
    Update-Status "Creating package for $packageID..."
    $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
    $progressBar.MarqueeAnimationSpeed = 30
    $progressBar.Visible = $true
    [System.Windows.Forms.Application]::DoEvents()  # Update UI - TODO: refactor to use Invoke-AsyncOperation
    
    $desired = $null
    if ($script:selectedPackageVersions.ContainsKey($packageID)) { 
      $desired = $script:selectedPackageVersions[$packageID] 
    }
    
    $resPkg = New-WingetPackageWithFallback `
      -PackageId $packageID `
      -PackageFolder $folder `
      -DesiredVersion $desired `
      -LatestVersion $package.Version `
      -AllowUserRetry `
      -ErrorAction SilentlyContinue
    
    if ($resPkg -and $resPkg.Succeeded) {
      $effectiveVersion = $resPkg.EffectiveVersion
      if (-not $effectiveVersion) { $effectiveVersion = $package.Version }
      Update-Status ("Package created successfully (version {0})" -f $effectiveVersion)
      $uploadButton.Enabled = $true
      if ($effectiveVersion) { $script:builtVersions[$packageID] = $effectiveVersion }
    } else {
      Update-Status "Package creation failed"
    }
  } finally {
    $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $progressBar.Visible = $false
    $progressBar.Value = 0
    $createButton.Enabled = $true
    $searchButton.Enabled = $true
    $versionsButton.Enabled = $true
  }
})

$uploadButton.Add_Click({
    if (-not $script:isConnected) {
        [void][System.Windows.Forms.MessageBox]::Show(
            "Please login to your tenant first.",
            "Information",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return
    }
    if (-not $dropdown.SelectedItem) { Update-Status "Please select a package."; return }
    $appName  = $dropdown.SelectedItem
    $package  = $script:packageMap[$appName]
    if (-not $package) { Update-Status "Selected item is invalid."; return }
    $packageID = $package.PackageID
    $version   = $null
    if ($script:builtVersions -and $script:builtVersions.ContainsKey($packageID)) { 
        $version = $script:builtVersions[$packageID] 
    } else { 
        $version = $package.Version 
    }
    if ([string]::IsNullOrWhiteSpace($packageID)) { 
        try { $packageID = ($appName -split '—')[-1].Trim() } catch { } 
    }
    if ([string]::IsNullOrWhiteSpace($version))   { Update-Status "Version could not be determined."; return }
    if ([string]::IsNullOrWhiteSpace($packageID)) { Update-Status "Cannot upload: failed to resolve PackageId."; return }
    $folder = [System.IO.Path]::GetFullPath($pathBox.Text.Trim())
    $forbiddenPaths = @(
      [Environment]::GetFolderPath('Windows'),
      [Environment]::GetFolderPath('System'),
      "$env:SystemRoot\System32",
      "$env:SystemRoot\SysWOW64",
      "$env:ProgramFiles",
      "${env:ProgramFiles(x86)}"
    )
    $isForbidden = $forbiddenPaths | Where-Object { $folder -eq $_ -or $folder.StartsWith($_ + '\') }
    if ($isForbidden) {
      [System.Windows.Forms.MessageBox]::Show(
        "The selected folder '$folder' is a protected system directory.`nPlease choose a different folder.",
        "Invalid Folder",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
      )
      return
    }
    if (-not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
    
    try {
        $uploadButton.Enabled = $false
        $createButton.Enabled = $false
        
        Update-Status "Uploading $packageID (v$version) to tenant..."
        $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
        $progressBar.MarqueeAnimationSpeed = 30
        $progressBar.Visible = $true
        [System.Windows.Forms.Application]::DoEvents()  # TODO: refactor to use Invoke-AsyncOperation
        
        Deploy-WtWin32App -PackageId $packageID -Version $version -RootPackageFolder $folder -ErrorAction Stop
        
        Update-Status "Upload completed successfully"
        $uploadButton.Enabled = $false
        $appSearchBox.Text = ""
        $dropdown.Items.Clear()
        
        # Clear version cache for this package so updates will use latest version
        if ($script:selectedPackageVersions.ContainsKey($packageID)) {
            $script:selectedPackageVersions.Remove($packageID)
            Write-Log "Cleared cached version for $packageID after upload"
        }
    } catch {
        $errorMsg = $_.Exception.Message
        Update-Status "Upload failed: See log for details"
        Write-Log "Upload error: $errorMsg"
        
        # Show detailed error dialog
        $errorDetails = "Upload of $packageID (v$version) failed.`n`n"
        $errorDetails += "Error: $errorMsg`n`n"
        $errorDetails += "Possible solutions:`n"
        $errorDetails += "1. Check if app already exists in Intune (delete and retry)`n"
        $errorDetails += "2. Update WinTuner module: Update-Module WinTuner`n"
        $errorDetails += "3. Try a different app to test`n"
        $errorDetails += "4. Check Intune service health`n"
        $errorDetails += "`nActivity logged to WinTuner_GUI.log"
        
        [System.Windows.Forms.MessageBox]::Show(
            $errorDetails,
            "Upload Failed",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    } finally {
        $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
        $progressBar.Visible = $false
        $progressBar.Value = 0
        $uploadButton.Enabled = $true
        $createButton.Enabled = $true
    }
})


# ----------------------------------------------
# Check All / Uncheck All Buttons
# ----------------------------------------------
$checkAllButton.Add_Click({
  for ($i = 0; $i -lt $updateListBox.Items.Count; $i++) {
    $updateListBox.SetItemChecked($i, $true)
  }
  foreach ($app in $script:updateApps) { $app.Checked = $true }
  Update-Status "All apps checked ($($updateListBox.Items.Count) items)"
})

$uncheckAllButton.Add_Click({
  for ($i = 0; $i -lt $updateListBox.Items.Count; $i++) {
    $updateListBox.SetItemChecked($i, $false)
  }
  foreach ($app in $script:updateApps) { $app.Checked = $false }
  Update-Status "All apps unchecked"
})

# Save checked state when user checks/unchecks an item in the update list
$updateListBox.Add_ItemCheck({
  param($sender, $e)
  $itemName = $updateListBox.Items[$e.Index]
  $appObj = $script:updateApps | Where-Object { $_.Name -eq $itemName } | Select-Object -First 1
  if ($appObj) {
    $appObj.Checked = ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked)
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

  if ([string]::IsNullOrWhiteSpace($filterText)) {
    # No filter - show all apps
    foreach ($app in @($script:updateApps)) {
      if ($app -and $app.Name) {
        $idx = $updateListBox.Items.Add($app.Name)
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
        $idx = $updateListBox.Items.Add($app.Name)
        if ($app.Checked) { $updateListBox.SetItemChecked($idx, $true) }
      }
    }
  }
  $updateListBox.EndUpdate()
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
# UPDATED: Robust & verbose "Search Updates"
# ----------------------------------------------
$updateSearchButton.Add_Click({
  if (-not $script:isConnected) { 
    Update-Status "Please login to your tenant first."; 
    return 
  }

  try {
    $updateSearchButton.Enabled = $false
    Update-Status "Loading apps from Intune..."
    
    # Reset UI / cache
    $updateFilterBox.Text = ""  # Clear filter
    $updateListBox.Items.Clear()
    $script:updateApps = @()

    # 1) Load all apps
    $all = @()
    try {
      $all = @(Get-WtWin32Apps -Superseded:$false -ErrorAction Stop)
      Write-Log ("Loaded {0} apps from Intune" -f $all.Count)
    } catch {
      Write-Log ("Failed to load apps: {0}" -f $_.Exception.Message)
      Update-Status "Failed to load apps from Intune"
      return
    }

    if ($all.Count -eq 0) {
      Update-Status "No apps found in Intune"
      return
    }

    # 2) Filter apps that need checking
    $appsToCheck = @($all | Where-Object { $_ -and $_.CurrentVersion })
    Write-Log ("Checking {0} apps for updates..." -f $appsToCheck.Count)
    
    # Show progress bar
    $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $progressBar.Value = 0
    $progressBar.Maximum = $appsToCheck.Count
    $progressBar.Visible = $true

    $candidates = [System.Collections.Generic.List[object]]::new()
    $processedCount = 0
    $totalCount = $appsToCheck.Count
    
    foreach ($app in $appsToCheck) {
      $processedCount++
      
      # Update progress every app
      try {
        $progressBar.Value = $processedCount
        Update-Status ("Checking ({0}/{1}): {2}" -f $processedCount, $totalCount, $app.Name)
        [System.Windows.Forms.Application]::DoEvents()  # TODO: refactor to use Invoke-AsyncOperation
      } catch { }
      
      # Try to resolve winget ID
      $wingetId = Resolve-WingetIdForApp -App $app
      $verified = $false
      
      if ($wingetId) {
        # Check winget for latest version
        try {
          $wgVersions = @(Get-WingetVersions -PackageId $wingetId -ErrorAction SilentlyContinue)
          if ($wgVersions -and $wgVersions.Count -gt 0) {
            $wgLatest = $wgVersions[0]
            if ($wgLatest) {
              try { 
                $app.LatestVersion = $wgLatest 
              } catch { }
              
              if (Test-IsNewerVersion $wgLatest $app.CurrentVersion) {
                $candidates.Add($app)
                Write-Log ("Update available: {0} ({1} -> {2})" -f $app.Name, $app.CurrentVersion, $wgLatest)
              }
              $verified = $true
            }
          }
        } catch {
          # Silently skip winget errors
        }
      }
      
      # Fallback: use LatestVersion if winget check failed
      if (-not $verified -and $app.LatestVersion) {
        if (Test-IsNewerVersion $app.LatestVersion $app.CurrentVersion) {
          $candidates.Add($app)
          Write-Log ("Update available (fallback): {0} ({1} -> {2})" -f $app.Name, $app.CurrentVersion, $app.LatestVersion)
        }
      }
    }
       # 3) Populate dropdown and cache
    $count = 0
    $updateListBox.BeginUpdate()
    
    $script:updateApps = [System.Collections.Generic.List[object]]::new()
    
    foreach ($app in ($candidates | Sort-Object Name)) {
      if (-not $app -or -not $app.Name) { continue }
      # Ensure Checked property exists for filter state persistence
      if (-not ($app | Get-Member -Name Checked -MemberType NoteProperty)) {
        $app | Add-Member -NotePropertyName Checked -NotePropertyValue $false -Force
      }
      [void]$updateListBox.Items.Add($app.Name)
      $script:updateApps.Add($app)
      $count++
    }
    $updateListBox.EndUpdate()

    if ($count -gt 0) {
      Update-Status ("Search updates completed: {0} candidate(s) found. Check items to update." -f $count)
      # Enable check/uncheck buttons
      $checkAllButton.Enabled = $true
      $uncheckAllButton.Enabled = $true
    } else {
      Update-Status "No update candidates found."
      $checkAllButton.Enabled = $false
      $uncheckAllButton.Enabled = $false
    }
  } finally {
    $updateSearchButton.Enabled = $true
    $progressBar.Maximum = 100
    $progressBar.Value = 0
    $progressBar.Visible = $false
  }
})

# -----------------------------
# UPDATED: Update Checked Apps flow
# -----------------------------
$updateSelectedButton.Add_Click({
    # Get checked items
    $checkedApps = [System.Collections.Generic.List[object]]::new()

    Write-Log "Processing $($updateListBox.CheckedItems.Count) checked items from UI"
    Write-Log "global:updateApps cache has $($script:updateApps.Count) apps"

    foreach ($itemName in $updateListBox.CheckedItems) {
        Write-Log "Looking for app: '$itemName'"

        # Find matching app in cache
        $foundApp = $null
        foreach ($cachedApp in $script:updateApps) {
            if ($cachedApp -and $cachedApp.Name -eq $itemName) {
                $foundApp = $cachedApp
                break
            }
        }

        if ($foundApp) {
            Write-Log "Found: $($foundApp.Name) (Current: $($foundApp.CurrentVersion), Latest: $($foundApp.LatestVersion), GraphId: $($foundApp.GraphId))"
            $checkedApps.Add($foundApp)
        } else {
            Write-Log "WARNING: Could not find '$itemName' in cache!"
            Write-Log "Available apps in cache: $($script:updateApps.Name -join ', ')"
        }
    }

    if ($checkedApps.Count -eq 0) {
        Update-Status "No valid apps found. Try 'Search Updates' again."
        Write-Log "ERROR: 0 apps matched from $($updateListBox.CheckedItems.Count) checked items"
        return
    }

    Write-Log "Successfully matched $($checkedApps.Count) apps for update"

    $rootPackageFolder = [System.IO.Path]::GetFullPath($pathBox.Text.Trim())
    $forbiddenPaths = @(
      [Environment]::GetFolderPath('Windows'),
      [Environment]::GetFolderPath('System'),
      "$env:SystemRoot\System32",
      "$env:SystemRoot\SysWOW64",
      "$env:ProgramFiles",
      "${env:ProgramFiles(x86)}"
    )
    $isForbidden = $forbiddenPaths | Where-Object { $rootPackageFolder -eq $_ -or $rootPackageFolder.StartsWith($_ + '\') }
    if ($isForbidden) {
      [System.Windows.Forms.MessageBox]::Show(
        "The selected folder '$rootPackageFolder' is a protected system directory.`nPlease choose a different folder.",
        "Invalid Folder",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
      )
      return
    }
    if (-not (Test-Path $rootPackageFolder)) {
        New-Item -ItemType Directory -Path $rootPackageFolder -Force | Out-Null
    }

    try {
        Update-Status "Starting update for $($checkedApps.Count) checked apps..."
        $batchResult = Invoke-AppUpdateBatch -Apps $checkedApps -RootPackageFolder $rootPackageFolder
        Update-Status "Checked apps updated: $($batchResult.SuccessCount) successful, $($batchResult.FailedList.Count) failed"
    } catch {
        Update-Status "Update error: $($_.Exception.Message)"
        Write-Log "updateSelectedButton error: $($_.Exception.Message)"
    }
})


# -------------------------
# UPDATED: Update All flow
# -------------------------
$updateAllButton.Add_Click({
    $rootPackageFolder = [System.IO.Path]::GetFullPath($pathBox.Text.Trim())
    $forbiddenPaths = @(
      [Environment]::GetFolderPath('Windows'),
      [Environment]::GetFolderPath('System'),
      "$env:SystemRoot\System32",
      "$env:SystemRoot\SysWOW64",
      "$env:ProgramFiles",
      "${env:ProgramFiles(x86)}"
    )
    $isForbidden = $forbiddenPaths | Where-Object { $rootPackageFolder -eq $_ -or $rootPackageFolder.StartsWith($_ + '\') }
    if ($isForbidden) {
      [System.Windows.Forms.MessageBox]::Show(
        "The selected folder '$rootPackageFolder' is a protected system directory.`nPlease choose a different folder.",
        "Invalid Folder",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
      )
      return
    }
    if (-not (Test-Path $rootPackageFolder)) {
        New-Item -ItemType Directory -Path $rootPackageFolder -Force | Out-Null
    }

    # Build candidate list to show in confirmation dialog
    try {
        $updatedApps = @(Get-WtWin32Apps -Update $true -Superseded $false)
    } catch {
        Write-Log "Get-WtWin32Apps update threw: $($_)"
        $updatedApps = @()
    }

    $updatedApps = @(( $updatedApps | Where-Object {
        $_.LatestVersion -and $_.CurrentVersion -and (Test-IsNewerVersion $_.LatestVersion $_.CurrentVersion)
    } | Sort-Object Name ))

    if (-not $updatedApps -or $updatedApps.Count -eq 0) {
        Update-Status "No update candidates found."
        return
    }

    # Show confirmation dialog
    $appNames = ($updatedApps | Select-Object -ExpandProperty Name) -join "`r`n"
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "The following apps will be updated:`r`n$appNames",
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
        $batchResult = Invoke-AppUpdateBatch -Apps $updatedApps -RootPackageFolder $rootPackageFolder
        Update-Status "All Updates Completed: $($batchResult.SuccessCount) successful, $($batchResult.FailedList.Count) failed"
    } catch {
        Update-Status "Mass update error: $($_.Exception.Message)"
        Write-Log "updateAllButton error: $($_.Exception.Message)"
    }
})


$removeOldAppsButton.Add_Click({
  try {
    $supersededApps = @(Get-WtWin32Apps -Superseded:$true -ErrorAction Stop)
  } catch {
    Update-Status "Error fetching superseded apps: $($_.Exception.Message)"
    Write-Log "removeOldApps error: $($_.Exception.Message)"
    return
  }
  try {
    if ($supersededApps.Count -eq 0) { Update-Status "No Superseded Apps Found"; return }
    $appNames = ($supersededApps | Select-Object -ExpandProperty Name) -join "`r`n"
    $result = [System.Windows.Forms.MessageBox]::Show(
      "The following outdated apps will be removed:`r`n$appNames",
      "Confirmation",
      [System.Windows.Forms.MessageBoxButtons]::YesNo,
      [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
      $progressBar.Value = 0
      $progressBar.Visible = $true
      foreach ($app in @($supersededApps)) {
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
      }
      $progressBar.Maximum = 100
      $progressBar.Value = 100
      Update-Status "Deleted all superseded Apps..."
      try { $supersededSearchButton.PerformClick() } catch {}
    } else {
      Update-Status "Removal aborted."
    }
  } catch {
    Write-Log "Error loading superseded apps: $($_.Exception.Message)"
    Update-Status "Error: $($_.Exception.Message)"
  }
})

# Handler: Search superseded apps
$script:supersededApps = @()
$supersededSearchButton.Add_Click({
  try {
    Update-Status "Search for superseded apps..."
    $script:supersededApps = Get-WtWin32Apps -Superseded $true
    $supersededDropdown.Items.Clear()
    foreach ($app in @($script:supersededApps)) {
      $name = $app.Name
      $version = $app.CurrentVersion
      $display = "$name — $version"
      [void]$supersededDropdown.Items.Add($display)
    }
    if ($supersededDropdown.Items.Count -gt 0) { $supersededDropdown.SelectedIndex = 0 }
    Update-Status ("Search completed: {0} superseded Apps found." -f $supersededDropdown.Items.Count)
  } catch {
    Write-Log "Superseded search error: $($_.Exception.Message)"
    Update-Status ("Error while search: {0}" -f $_.Exception.Message)
  }
})


 

# Handler: Delete selected superseded app
$deleteSelectedAppButton.Add_Click({
  if (-not $script:supersededApps -or $supersededDropdown.SelectedIndex -lt 0) {
    Update-Status "Please first select a superseded app from the dropdown."
    return
  }
  $app = $script:supersededApps[$supersededDropdown.SelectedIndex]
  $result = [System.Windows.Forms.MessageBox]::Show(
    "Delete App '" + $app.Name + "'?",
    "Confirmation",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Question
  )
  if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
    try {
      Remove-WtWin32App -GraphId $app.GraphId -ErrorAction Stop
      Update-Status ("Deleted: {0}" -f $app.Name)
      try { $supersededSearchButton.PerformClick() } catch {}
    } catch {
      Update-Status ("Error while removal: {0}" -f $_.Exception.Message)
    }
  } else {
    Update-Status "Removal aborted."
  }
})

$logoutButton.Add_Click({
  try {
    Disconnect-WtWinTuner -ErrorAction Stop
  } catch {
    Write-Log "Logout warning: $($_.Exception.Message)"
  }
  try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
  $script:isConnected = $false
  $script:currentUserUpn = ""
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
            if (($item.DisplayName -notmatch "(?i)$escapedSearch") -and ($item.WingetApp.Name -notmatch "(?i)$escapedSearch")) {
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
            # Stellt den Haken (Checked-Status) wieder her, falls er vorher gesetzt war
            $discoveredListBox.SetItemChecked($idx, $obj.Checked)
        }
    }
    $discoveredListBox.EndUpdate()
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
    $itemText = $discoveredListBox.Items[$e.Index]
    $obj = $script:discoveredRaw | Where-Object { $_.DisplayText -eq $itemText } | Select-Object -First 1
    if ($obj) {
        $obj.Checked = ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked)
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

$scanDiscoveredButton.Add_Click({
  if (-not $script:isConnected) { Update-Status "Please login first."; return }

  # Speichere die originalen Streams und schalte sie stumm, um Threading-Crashes zu vermeiden
  $oldProgress = $ProgressPreference
  $oldInfo = $InformationPreference
  $ProgressPreference = 'SilentlyContinue'
  $InformationPreference = 'SilentlyContinue'

  try {
    $scanDiscoveredButton.Enabled = $false
    $deployDiscoveredButton.Enabled = $false
    $discoveredListBox.Items.Clear()
    $script:discoveredRaw = [System.Collections.Generic.List[object]]::new()
    
    $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
    $progressBar.Visible = $true
    [System.Windows.Forms.Application]::DoEvents()  # TODO: refactor to use Invoke-AsyncOperation

# --- GRAPH-AUTH BLOCK (FIXED) ---
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        Update-Status "Microsoft.Graph module not found..."
        [System.Windows.Forms.MessageBox]::Show(
            "Microsoft.Graph module not found.`n`nPlease install it first:`nInstall-Module Microsoft.Graph -Scope CurrentUser",
            "Module Not Found",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    # Liste ALLER benötigten Scopes für Discovered Apps
    $requiredScopes = @(
        "DeviceManagementApps.ReadWrite.All", 
        "DeviceManagementManagedDevices.Read.All", 
        "Directory.Read.All"
    )

    $mgContext = Get-MgContext -ErrorAction SilentlyContinue
    $needsAuth = $false

    if (-not $mgContext) {
        $needsAuth = $true
    } else {
        # Prüfen, ob ALLE erforderlichen Scopes im aktuellen Token vorhanden sind
        foreach ($s in $requiredScopes) {
            if ($mgContext.Scopes -notcontains $s) {
                $needsAuth = $true
                break
            }
        }

        $userMatch = ($mgContext.Account -eq $script:currentUserUpn)
        if (-not $userMatch) { $needsAuth = $true }

        if ($needsAuth) {
            Update-Status "Clearing old Graph session (Scope missing or wrong Tenant)..."
            [System.Windows.Forms.Application]::DoEvents()
            try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
        }
    }

    if ($needsAuth) {
        Update-Status "Authenticating with MS Graph for $($script:currentUserUpn)..."
        [System.Windows.Forms.Application]::DoEvents()
        $tenantDomain = $script:currentUserUpn.Split('@')[1]
        
        # Jetzt mit dem vollständigen Array an Scopes anmelden
        $null = Connect-MgGraph -TenantId $tenantDomain -Scopes $requiredScopes -NoWelcome -ErrorAction Stop *>&1
    }

    # 1. Vorhandene Apps checken (EXTREM SCHNELL DURCH "Resolve" STATT "Try-Resolve")
    Update-Status "Loading existing managed apps to filter them out..."
    [System.Windows.Forms.Application]::DoEvents()  # TODO: refactor to use Invoke-AsyncOperation
    $existingApps = @(Get-WtWin32Apps -Superseded:$false -ErrorAction SilentlyContinue 3>$null 4>$null)
    $existingPackageIds = [System.Collections.Generic.List[object]]::new()
    foreach ($eApp in $existingApps) {
		[System.Windows.Forms.Application]::DoEvents()  # TODO: refactor to use Invoke-AsyncOperation
        $id = Resolve-WtWingetId -AppOrResult $eApp
        if ($id) { $existingPackageIds.Add($id) }
    }

    # 2. Hole ALLE Discovered Apps aus Intune (inklusive Paginierung)
    Update-Status "Fetching ALL detected apps from Intune API (this might take a moment)..."
    [System.Windows.Forms.Application]::DoEvents()  # TODO: refactor to use Invoke-AsyncOperation
    
    $uri = "https://graph.microsoft.com/beta/deviceManagement/detectedApps?`$top=500&`$orderby=deviceCount desc"
    $detectedApps = [System.Collections.Generic.List[object]]::new()
    $maxPages = 100
    $pageCount = 0

    do {
        $response = Invoke-MgRestMethod -Uri $uri -Method GET -ErrorAction Stop
        if ($response.value) { $detectedApps.AddRange([object[]]$response.value) }
        $uri = $response.'@odata.nextLink'
        $pageCount++
        if ($pageCount -ge $maxPages) {
            Write-Log "Warning: Graph API pagination limit ($maxPages pages) reached. Some apps may not be shown."
            break
        }
    } while ($uri)

    if (-not $detectedApps -or $detectedApps.Count -eq 0) {
        Update-Status "No discovered apps found in Intune."
        return
    }

    $filteredApps = @($detectedApps | Where-Object { 
        $_.publisher -notmatch "(?i)Intel|HP|Dell|Lenovo|AMD|NVIDIA|Realtek|Synaptics|VMware" 
    })

    $total = $filteredApps.Count
    $current = 0
    $matchCount = 0

    $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $progressBar.Maximum = $total
    $progressBar.Value = 0

    foreach ($app in $filteredApps) {
		[System.Windows.Forms.Application]::DoEvents()  # TODO: refactor to use Invoke-AsyncOperation
        $current++
        $progressBar.Value = $current
        Update-Status "Analyzing ($current/$total): $($app.displayName)..."
        [System.Windows.Forms.Application]::DoEvents()  # TODO: refactor to use Invoke-AsyncOperation

        try {
            # 1. Entfernt restlos alles, was in Klammern steht (z.B. "(x64 de)", "(x86 en-US)")
            $searchName = $app.displayName -replace '\s*\([^)]*\)', ''
            # 2. Entfernt typische Versionsnummern, die aus Zahlen und Punkten bestehen
            $searchName = $searchName -replace '\s+[\d\.]+', ''
            $searchName = $searchName.Trim()

            if ([string]::IsNullOrWhiteSpace($searchName)) { continue }

            $wingetResults = @(Search-WtWinGetPackage -SearchQuery $searchName -ErrorAction SilentlyContinue 3>$null 4>$null)
            
            $bestMatch = $null
            $highestScore = 0
            
            foreach ($wgApp in $wingetResults) {
				[System.Windows.Forms.Application]::DoEvents()  # TODO: refactor to use Invoke-AsyncOperation
                $score = Get-StringSimilarity -str1 $app.displayName -str2 $wgApp.Name
                if ($score -gt $highestScore) {
                    $highestScore = $score
                    $bestMatch = $wgApp
                }
            }

            if ($bestMatch -and $highestScore -ge 50) {
                if ($existingPackageIds -contains $bestMatch.PackageID) { continue }

                # NEU: Prüfen, ob wir diese Winget-App (PackageID) schon in der Liste haben
                $existingEntry = $script:discoveredRaw | Where-Object { $_.WingetApp.PackageID -eq $bestMatch.PackageID } | Select-Object -First 1

                if ($existingEntry) {
                    # App existiert bereits in der Liste: Wir addieren die Geräteanzahl (DeviceCount)
                    $existingEntry.DeviceCount += $app.deviceCount
                    # Den Anzeigetext mit der neuen, kombinierten Anzahl aktualisieren
                    $existingEntry.DisplayText = "[$($existingEntry.DeviceCount) PCs] $($existingEntry.DisplayName) ($($existingEntry.Publisher))  -->  Winget: $($existingEntry.WingetApp.Name)"
                } else {
                    # App ist neu: Wir nutzen den sauberen Winget-Namen (ohne Versionsnummern aus Intune)
                    $cleanName = $bestMatch.Name 
                    $itemObj = [pscustomobject]@{
                        DisplayName = $cleanName
                        Publisher   = $app.publisher
                        DeviceCount = $app.deviceCount
                        WingetApp   = $bestMatch
                        Checked     = $false
                        DisplayText = "[$($app.deviceCount) PCs] $cleanName ($($app.publisher))  -->  Winget: $($bestMatch.Name)"
                    }
                    $script:discoveredRaw.Add($itemObj)
                    $matchCount++
                }
            }
        } catch {
            Write-Log "Failed to process '$($app.displayName)': $($_.Exception.Message)"
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

    if ($matchCount -gt 0) {
        Update-Status "Found $matchCount Winget match(es). Filter, sort, or deploy them!"
        $deployDiscoveredButton.Enabled = $true
        $checkAllDiscoveredButton.Enabled = $true
        $uncheckAllDiscoveredButton.Enabled = $true
    } else {
        Update-Status "No Winget matches found (or all are already managed)."
    }

  } catch {
    Update-Status "Error fetching discovered apps: $($_.Exception.Message)"
    Write-Log "Scan Discovered Error: $($_.Exception.Message)"
  } finally {
    $ProgressPreference = $oldProgress
    $InformationPreference = $oldInfo
    $scanDiscoveredButton.Enabled = $true
    $progressBar.Maximum = 100
    $progressBar.Value = 0
    $progressBar.Visible = $false
  }
})

$deployDiscoveredButton.Add_Click({
    $checkedItems = @($script:discoveredRaw | Where-Object { $_.Checked })
    if ($checkedItems.Count -eq 0) { 
        Update-Status "No apps checked."
        return 
    }

    $rootFolder = $script:settings.DefaultPackagePath
    if (-not $rootFolder) { $rootFolder = "C:\Temp" }
    if (-not (Test-Path $rootFolder)) { New-Item -ItemType Directory -Path $rootFolder -Force | Out-Null }

    $oldProgress = $ProgressPreference
    $oldInfo = $InformationPreference
    $ProgressPreference = 'SilentlyContinue'
    $InformationPreference = 'SilentlyContinue'

    try {
        $deployDiscoveredButton.Enabled = $false
        $scanDiscoveredButton.Enabled = $false
        $checkAllDiscoveredButton.Enabled = $false
        $uncheckAllDiscoveredButton.Enabled = $false
        
        $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
        $progressBar.Maximum = $checkedItems.Count
        $progressBar.Value = 0
        $progressBar.Visible = $true

        $successCount = 0
        $failedCount = 0
        $i = 0

        foreach ($item in $checkedItems) {
            $i++
            $progressBar.Value = $i
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
                
                $effVersion = if ($pkgRes.EffectiveVersion) { $pkgRes.EffectiveVersion } else { $version }

                Write-Log "Uploading new app to tenant: $packageId v$effVersion"
                Deploy-WtWin32App `
                    -PackageId $packageId `
                    -Version $effVersion `
                    -RootPackageFolder $rootFolder `
                    -ErrorAction Stop
                
                $successCount++
                Write-Log "Successfully deployed new app: $packageId"
            } catch {
                $failedCount++
                Write-Log "Failed to deploy $($wingetApp.Name): $($_.Exception.Message)"
            }
        }

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
        $deployDiscoveredButton.Enabled = $true
        $scanDiscoveredButton.Enabled = $true
        $checkAllDiscoveredButton.Enabled = $true
        $uncheckAllDiscoveredButton.Enabled = $true
        $progressBar.Maximum = 100
        $progressBar.Value = 0
        $progressBar.Visible = $false
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
            Save-Settings 
        } 
    } catch {}

    # 2. Wenn bereits geschlossen wird, ignorieren
    if ($script:_closingInProgress) { return }
    $script:_closingInProgress = $true

    # 3. Falls verbunden, regulär abmelden
    if ($script:isConnected) {
        try {
            $form.Enabled = $false
            if ($statusLabel) { 
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
  Invoke-AsyncOperation -StatusText "Checking for updates..." -ScriptBlock {
    Test-AppUpdateAvailable
  } -OnComplete {
    param($updateResult)
    if ($updateResult -and -not $updateResult.Error -and -not $updateResult.ErrorMessage) {
      if ($updateResult.UpdateAvailable) {
        try {
          $msg  = "A new version of WinTuner GUI is available!`n`n"
          $msg += "Current version: v$($script:appVersion)`n"
          $msg += "Latest version:  v$($updateResult.LatestVersion)`n`n"

          if ($updateResult.DownloadUrl) {
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

              $success = Invoke-AppSelfUpdate -DownloadUrl $updateResult.DownloadUrl -HashUrl $updateResult.HashUrl

              if ($success) {
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
              }
            } else {
              Update-Status "Update available: v$($updateResult.LatestVersion) - Go to Settings to update later."
            }
          } else {
            $msg += "No direct download available for this release.`n"
            $msg += "Please download manually from:`n$($updateResult.ReleaseUrl)"

            [System.Windows.Forms.MessageBox]::Show(
              $msg,
              "Update Available",
              [System.Windows.Forms.MessageBoxButtons]::OK,
              [System.Windows.Forms.MessageBoxIcon]::Information
            )
          }
        } catch {
          try { Write-Log "Startup update dialog error: $($_.Exception.Message)" } catch {}
        }
      } else {
        $latestVer = if ($updateResult -and $updateResult.LatestVersion) { $updateResult.LatestVersion } else { "unknown" }
        Update-Status "WinTuner GUI is up to date – Local: v$($script:appVersion) | GitHub: v$latestVer"
      }
    } else {
      # Error case – still show a clear status
      Update-Status "Update check failed – running v$($script:appVersion)"
    }
  }
})

# Tooltips for main buttons
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 5000
$toolTip.InitialDelay = 500
$toolTip.ReshowDelay  = 500
$toolTip.ShowAlways   = $true

$toolTip.SetToolTip($searchButton,          "Search the WinGet repository for applications")
$toolTip.SetToolTip($versionsButton,        "Select a specific version for the selected app")
$toolTip.SetToolTip($browseButton,          "Choose the local folder to store package files")
if ($createButton)          { $toolTip.SetToolTip($createButton,          "Create the .wtpackage file locally") }
if ($uploadButton)          { $toolTip.SetToolTip($uploadButton,          "Upload and deploy the package to Microsoft Intune") }
if ($updateSearchButton)    { $toolTip.SetToolTip($updateSearchButton,    "Scan all Intune Win32 apps for available WinGet updates") }
if ($updateAllButton)       { $toolTip.SetToolTip($updateAllButton,       "Update all apps with available updates") }
if ($updateSelectedButton)  { $toolTip.SetToolTip($updateSelectedButton,  "Update only the checked apps in the list") }
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
if ($supersededSearchButton){ $toolTip.SetToolTip($supersededSearchButton,"Search for outdated (superseded) app versions in Intune") }
if ($deleteSelectedAppButton){ $toolTip.SetToolTip($deleteSelectedAppButton, "Delete the app currently selected in the dropdown from Intune") }
if ($removeOldAppsButton)   { $toolTip.SetToolTip($removeOldAppsButton,   "Delete all superseded app versions from Intune at once") }

# tabDiscovered
if ($deployDiscoveredButton){ $toolTip.SetToolTip($deployDiscoveredButton,"Deploy the checked discovered apps to Microsoft Intune") }
if ($checkAllDiscoveredButton)  { $toolTip.SetToolTip($checkAllDiscoveredButton,   "Check all apps in the discovered apps list") }
if ($uncheckAllDiscoveredButton){ $toolTip.SetToolTip($uncheckAllDiscoveredButton, "Uncheck all apps in the discovered apps list") }

# tabSettings
if ($browsePathButton)         { $toolTip.SetToolTip($browsePathButton,         "Open a folder browser to choose the default package folder") }
if ($autoCheckUpdatesCheckbox) { $toolTip.SetToolTip($autoCheckUpdatesCheckbox, "Automatically scan for app updates each time you log in") }
if ($rememberMeCheckbox)       { $toolTip.SetToolTip($rememberMeCheckbox,       "Save your username so it is pre-filled on the next launch") }
if ($saveSettingsButton)       { $toolTip.SetToolTip($saveSettingsButton,       "Save all settings to disk") }
if ($clearCacheButton)         { $toolTip.SetToolTip($clearCacheButton,         "Clear the locally cached WinGet version list") }
if ($checkUpdateButton)        { $toolTip.SetToolTip($checkUpdateButton,        "Check GitHub for a newer version of WinTuner GUI") }

# Run the form mit finalem Sicherheitsnetz
try {
    [System.Windows.Forms.Application]::Run($form)
} catch {
    # Fängt ab, falls das Skript als Ganzes unerwartet beendet wird
    Write-FileLog "FATAL SCRIPT CRASH: $($_.Exception.Message)`n$($_.ScriptStackTrace)"
}
