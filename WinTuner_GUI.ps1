# WinTuner GUI by Manuel Höfler  (patched + deploy fix + robust update search + regex quoting fix)
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
$ErrorActionPreference = 'SilentlyContinue'
$WarningPreference = 'SilentlyContinue'
$InformationPreference = 'SilentlyContinue'
$VerbosePreference = 'SilentlyContinue'
$DebugPreference = 'SilentlyContinue'
$ProgressPreference = 'SilentlyContinue'

# Redirect all output streams to prevent threading issues
$PSDefaultParameterValues = @{
  '*:ErrorAction' = 'SilentlyContinue'
  '*:WarningAction' = 'SilentlyContinue'
  '*:InformationAction' = 'SilentlyContinue'
  '*:ProgressAction' = 'SilentlyContinue'
  '*:Verbose' = $false
  '*:Debug' = $false
}

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

function Try-ResolveWingetIdForApp {
  param([object]$App)
  $id = Resolve-WtWingetId -AppOrResult $App
  if (-not [string]::IsNullOrWhiteSpace($id)) { return $id }
  try {
    $res = @(Search-WtWinGetPackage -SearchQuery $App.Name)
  } catch { $res = @() }
  if ($res -and $res.Count -gt 0) {
    $exact = $res | Where-Object { $_.Name -and ($_.Name -eq $App.Name) } | Select-Object -First 1
    if ($exact -and $exact.PackageID) { return [string]$exact.PackageID }
    $first = $res | Select-Object -First 1
    if ($first -and $first.PackageID) { return [string]$first.PackageID }
  }
  return $null
}

function Get-PreviousWingetVersion {
  param([string]$PackageId, [string]$LatestVersion)
  try { $output = & winget show --id $PackageId --versions 2>$null } catch { return $null }
  if (-not $output) { return $null }
  $cand = @()
  foreach ($line in @($output)) {
    $t = ($line -replace '^[\s\-•]+','').Trim()
    if (-not $t) { continue }
    if ($t -match '^(\d+)(\.[0-9A-Za-z]+)*([\-+._][0-9A-Za-z]+)*$') { $cand += $t }
  }
  if (-not $cand -or $cand.Count -eq 0) { return $null }
  $unique = @($cand | Select-Object -Unique)
  if ($LatestVersion) { $unique = @($unique | Where-Object { $_ -ne $LatestVersion }) }
  $parsed = foreach ($v in $unique) {
    $ok = $false; $vo = $null
    try { $vo = [version]$v; $ok = $true } catch {}
    [pscustomobject]@{ Text = $v; Parsed = $vo; Numeric = $ok }
  }
  $sorted = @()
  if ($parsed | Where-Object Numeric) { $sorted = @($parsed | Where-Object Numeric | Sort-Object Parsed -Descending) }
  else { $sorted = @($parsed | Sort-Object Text -Descending) }
  if ($sorted.Count -gt 0) { return $sorted[0].Text }
  return $null
}

function Get-WingetVersions {
  param([string]$PackageId)
  
  # Check cache first (speeds up repeated searches)
  if ($global:wingetVersionCache.ContainsKey($PackageId)) {
    return $global:wingetVersionCache[$PackageId]
  }
  
  # Query winget
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
  
  # Cache result for future use
  $global:wingetVersionCache[$PackageId] = $result
  
  return $result
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
        try { New-WtWingetPackage -PackageId $PackageId -PackageFolder $PackageFolder -Version $prev -ErrorAction Stop; return [pscustomobject]@{ Succeeded=$true; EffectiveVersion=$prev } } catch { return [pscustomobject]@{ Succeeded=$false; EffectiveVersion=$null } }
      } else { return [pscustomobject]@{ Succeeded=$false; EffectiveVersion=$null } }
    } elseif ($m -match 'Hash mismatch') {
      if ($AllowUserRetry) {
        $res = [System.Windows.Forms.MessageBox]::Show("Hash mismatch detected. Retry download? Click Yes to retry, No to use previous available version, or Cancel to abort.", "Hash mismatch", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel, [System.Windows.Forms.MessageBoxIcon]::Warning, [System.Windows.Forms.MessageBoxDefaultButton]::Button1)
        if ($res -eq [System.Windows.Forms.DialogResult]::Yes) {
          try { if ($attemptVersion) { New-WtWingetPackage -PackageId $PackageId -PackageFolder $PackageFolder -Version $attemptVersion -ErrorAction Stop } else { New-WtWingetPackage -PackageId $PackageId -PackageFolder $PackageFolder -ErrorAction Stop }; return [pscustomobject]@{ Succeeded=$true; EffectiveVersion=$attemptVersion } } catch { return [pscustomobject]@{ Succeeded=$false; EffectiveVersion=$null } }
        } elseif ($res -eq [System.Windows.Forms.DialogResult]::No) {
          $latest = $attemptVersion; if (-not $latest) { $latest = $LatestVersion }
          $prev = Get-PreviousWingetVersion -PackageId $PackageId -LatestVersion $latest
          # Only allow previous if it's newer than current tenant version (if known)
          if ($prev -and ( -not $InstalledVersion -or (Test-IsNewerVersion $prev $InstalledVersion) )) {
            try { New-WtWingetPackage -PackageId $PackageId -PackageFolder $PackageFolder -Version $prev -ErrorAction Stop; return [pscustomobject]@{ Succeeded=$true; EffectiveVersion=$prev } } catch { return [pscustomobject]@{ Succeeded=$false; EffectiveVersion=$null } }
          } else { return [pscustomobject]@{ Succeeded=$false; EffectiveVersion=$null } }
        } else { return [pscustomobject]@{ Succeeded=$false; EffectiveVersion=$null } }
      } else { return [pscustomobject]@{ Succeeded=$false; EffectiveVersion=$null } }
    } else { return [pscustomobject]@{ Succeeded=$false; EffectiveVersion=$null } }
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
  $candidates = @()
  foreach ($app in @($all)) {
    if (-not $app -or -not $app.CurrentVersion) { continue }
    $wingetId = Try-ResolveWingetIdForApp -App $app
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
      $candidates += $app
    }
  }
  return ,$candidates
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
      $result.Message = "Package creation failed for $AppName"
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

# Dark mode theme colors
$global:darkTheme = @{
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
$global:lightTheme = @{
  BackColor       = [System.Drawing.Color]::FromArgb(240, 240, 240)
  ForeColor       = [System.Drawing.Color]::FromArgb(0, 0, 0)
  ButtonBackColor = [System.Drawing.Color]::FromArgb(225, 225, 225)
  ButtonForeColor = [System.Drawing.Color]::FromArgb(0, 0, 0)
  TextBoxBackColor= [System.Drawing.Color]::FromArgb(255, 255, 255)
  TextBoxForeColor= [System.Drawing.Color]::FromArgb(0, 0, 0)
  TabBackColor    = [System.Drawing.Color]::FromArgb(240, 240, 240)
  TabForeColor    = [System.Drawing.Color]::FromArgb(0, 0, 0)
}

$global:isDarkMode   = $true
$global:currentTheme = $global:darkTheme

# Function to apply theme to all controls
function Apply-Theme {
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
    Apply-Theme -control $childControl -theme $theme
  }
}

# Function to toggle theme
function Toggle-Theme {
  $global:isDarkMode   = -not $global:isDarkMode
  $global:currentTheme = if ($global:isDarkMode) { $global:darkTheme } else { $global:lightTheme }
  Apply-Theme -control $form -theme $global:currentTheme
  # Button text indicates the action (target)
  $themeToggleButton.Text = if ($global:isDarkMode) { "Light Mode" } else { "Dark Mode" }
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
      $logPath = Join-Path $base 'WinTuner_GUI.log'
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
  $bw.Add_DoWork({
    param($sender, $e)
    try {
      $e.Result = & $ScriptBlock
    } catch {
      $e.Result = @{ Error = $_.Exception.Message }
      Write-Log "Async operation error: $($_.Exception.Message)"
    }
  })
  
  # On completion (runs on UI thread)
  $bw.Add_RunWorkerCompleted({
    param($sender, $e)
    
    # Restore progress bar to normal
    $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
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
  })
  
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

# Helper: check if WinTuner is connected (simple smoke test)
function Test-WtConnected {
  try {
    $null = Get-WtWin32Apps -Update:$false -Superseded:$false -ErrorAction Stop | Select-Object -First 1 | Out-Null
    return $true
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
  } else {
    $loginButton.Visible = $true
    $usernameBox.Visible = $true
    $usernameLabel.Visible = $true
    if ($usernameError) { $usernameError.Visible = $true }
    $tabControl.Visible = $true
    $logoutButton.Visible = $false
  }
  if ($rememberCheckBox) { $rememberCheckBox.Visible = -not $Connected }
  if ($updateSearchButton) { $updateSearchButton.Enabled = $Connected }
  if ($updateSelectedButton) { $updateSelectedButton.Enabled = $Connected }
  if ($updateAllButton) { $updateAllButton.Enabled = $Connected }
  if ($supersededSearchButton) { $supersededSearchButton.Enabled = $Connected }
  if ($deleteSelectedAppButton) { $deleteSelectedAppButton.Enabled = $Connected }
  if ($removeOldAppsButton) { $removeOldAppsButton.Enabled = $Connected }
  if ($loadRollbackAppsButton) { $loadRollbackAppsButton.Enabled = $Connected }
  if ($mapWingetIdButton) { $mapWingetIdButton.Enabled = $Connected }
  if ($loginInfoLabel) {
    $loginInfoLabel.Visible = $Connected
    if ($Connected -and $global:currentUserUpn) { $loginInfoLabel.Text = "Logged in as: $($global:currentUserUpn)" }
  }
}

$global:isConnected = $false
$global:currentUserUpn = ""

# Track effective built versions per PackageId
$global:builtVersions = @{}

# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = "WinTuner GUI"
$form.Size = New-Object System.Drawing.Size(900, 850)
$form.Padding = '5,5,5,5'

# Theme toggle button (top right)
$themeToggleButton = New-Object System.Windows.Forms.Button
$themeToggleButton.Text = "Light Mode"  # indicates action from dark -> light
$themeToggleButton.Location = New-Object System.Drawing.Point(800, 5)
$themeToggleButton.Size = New-Object System.Drawing.Size(80, 25)
$themeToggleButton.Add_Click({ Toggle-Theme })
$form.Controls.Add($themeToggleButton)

# Username label and textbox
$usernameLabel = New-Object System.Windows.Forms.Label
$usernameLabel.Text = "Username:"
$usernameLabel.Location = New-Object System.Drawing.Point(10,20)
$usernameLabel.AutoSize = $true
$form.Controls.Add($usernameLabel)

$usernameBox = New-Object System.Windows.Forms.TextBox
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
$usernameBox.Location = New-Object System.Drawing.Point(100,20)
$usernameBox.Width = 450
$usernameBox.BorderStyle = 'FixedSingle'
$form.Controls.Add($usernameBox)

# Validation hint label for username
$usernameError = New-Object System.Windows.Forms.Label
$usernameError.Text = ""
$usernameError.Location = New-Object System.Drawing.Point(100,45)
$usernameError.AutoSize = $true
$usernameError.ForeColor = [System.Drawing.Color]::FromArgb(220,80,80)
$form.Controls.Add($usernameError)

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
$form.Controls.Add($statusLabel)

# Output textbox (Log area below tabs and progress bar)
$outputBox = New-Object System.Windows.Forms.TextBox
$outputBox.Location = New-Object System.Drawing.Point(10, 680)
$outputBox.Size = New-Object System.Drawing.Size(760, 60)
$outputBox.Multiline = $true
$outputBox.ScrollBars = "Vertical"
$outputBox.ReadOnly = $true
$form.Controls.Add($outputBox)

# Progress bar (appears between tabs and log when active)
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 655)
$progressBar.Width = 760
$progressBar.Height = 20
$progressBar.Visible = $false
$form.Controls.Add($progressBar)

# Logout button
$logoutButton = New-Object System.Windows.Forms.Button
$logoutButton.Text = "Tenant Logout"
$logoutButton.Location = New-Object System.Drawing.Point(570,20)
$logoutButton.Width = 180
$logoutButton.Visible = $false
$form.Controls.Add($logoutButton)

$loginInfoLabel = New-Object System.Windows.Forms.Label
$loginInfoLabel.Text = ""
$loginInfoLabel.Location = New-Object System.Drawing.Point(100,20)
$loginInfoLabel.AutoSize = $true
$loginInfoLabel.Visible = $false
$form.Controls.Add($loginInfoLabel)

# TabControl
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(10, 90)
$tabControl.Size = New-Object System.Drawing.Size(760, 560)
$tabControl.Visible = $true
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
$uploadButton.Visible = $false
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
$deleteSelectedAppButton.Text = "Ausgewählte App löschen"
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
# Rollback Section
# ==================================================
$rollbackHeaderLabel = New-Object System.Windows.Forms.Label
$rollbackHeaderLabel.Text = "Rollback to previous version"
$rollbackHeaderLabel.Location = New-Object System.Drawing.Point(100,430)
$rollbackHeaderLabel.AutoSize = $true
$rollbackHeaderLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$tabUpdate.Controls.Add($rollbackHeaderLabel)

# Dropdown: Select app to rollback (shows all versions)
$rollbackAppDropdown = New-Object System.Windows.Forms.ComboBox
$rollbackAppDropdown.Location = New-Object System.Drawing.Point(100,460)
$rollbackAppDropdown.Width = 300
$rollbackAppDropdown.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$tabUpdate.Controls.Add($rollbackAppDropdown)

# Dropdown for available versions
$rollbackVersionLabel = New-Object System.Windows.Forms.Label
$rollbackVersionLabel.Text = "Keep Version:"
$rollbackVersionLabel.Location = New-Object System.Drawing.Point(410,463)
$rollbackVersionLabel.AutoSize = $true
$tabUpdate.Controls.Add($rollbackVersionLabel)

$rollbackVersionDropdown = New-Object System.Windows.Forms.ComboBox
$rollbackVersionDropdown.Location = New-Object System.Drawing.Point(510,460)
$rollbackVersionDropdown.Width = 240
$rollbackVersionDropdown.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$tabUpdate.Controls.Add($rollbackVersionDropdown)

# Event: When app selected, load versions
$rollbackAppDropdown.Add_SelectedIndexChanged({
  if ($rollbackAppDropdown.SelectedIndex -ge 0) {
    # Extract app name (remove "(X versions)" suffix)
    $selectedText = $rollbackAppDropdown.Text
    $selectedAppName = $selectedText -replace ' \(\d+ versions\)$', ''
    Update-Status "Loading versions for $selectedAppName..."
    
    # Get all versions of this app from Intune
    $allVersions = @(Get-WtWin32Apps -Superseded:$false) + @(Get-WtWin32Apps -Superseded:$true) | Where-Object { $_.Name -eq $selectedAppName }
    $global:rollbackVersions = @($allVersions)
    
    $rollbackVersionDropdown.Items.Clear()
    
    if ($allVersions.Count -eq 0) {
      [void]$rollbackVersionDropdown.Items.Add("No versions found")
      $executeRollbackButton.Enabled = $false
      Update-Status "No versions found for $selectedAppName"
      return
    }
    
    # Sort by version descending (newest first)
    $sortedVersions = $allVersions | Sort-Object { 
      try { [version]$_.CurrentVersion } catch { $_.CurrentVersion }
    } -Descending
    
    $global:rollbackVersions = @($sortedVersions)
    
    # Mark newest as NEW, all others as OLD
    $isFirst = $true
    foreach ($ver in $sortedVersions) {
      if ($isFirst) {
        $status = "Current (NEW)"
        $isFirst = $false
      } else {
        $status = "Old (ROLLBACK)"
      }
      $display = "$($ver.CurrentVersion) — $status"
      [void]$rollbackVersionDropdown.Items.Add($display)
    }
    
    if ($rollbackVersionDropdown.Items.Count -gt 0) {
      $rollbackVersionDropdown.SelectedIndex = 0
      $executeRollbackButton.Enabled = $true
    }
    
    Update-Status "Loaded $($allVersions.Count) version(s) for $selectedAppName"
  }
})

# Load Apps button
$loadRollbackAppsButton = New-Object System.Windows.Forms.Button
$loadRollbackAppsButton.Text = "Load Apps for Rollback"
$loadRollbackAppsButton.Location = New-Object System.Drawing.Point(100,495)
$loadRollbackAppsButton.Width = 200
$loadRollbackAppsButton.Enabled = $false
$tabUpdate.Controls.Add($loadRollbackAppsButton)

# Execute Rollback button
$executeRollbackButton = New-Object System.Windows.Forms.Button
$executeRollbackButton.Text = "🔄 Execute Rollback"
$executeRollbackButton.Location = New-Object System.Drawing.Point(310,495)
$executeRollbackButton.Width = 190
$executeRollbackButton.Enabled = $false
$tabUpdate.Controls.Add($executeRollbackButton)

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
$defaultPathTextBox.Text = if ($global:settings.DefaultPackagePath) { $global:settings.DefaultPackagePath } else { "C:\Temp" }
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
$autoCheckUpdatesCheckbox.Checked = if ($global:settings.AutoCheckUpdates) { $global:settings.AutoCheckUpdates } else { $false }
$tabSettings.Controls.Add($autoCheckUpdatesCheckbox)

# RememberMe Checkbox (moved to settings)
$rememberMeCheckbox = New-Object System.Windows.Forms.CheckBox
$rememberMeCheckbox.Text = "Remember last username"
$rememberMeCheckbox.Location = New-Object System.Drawing.Point(20,130)
$rememberMeCheckbox.AutoSize = $true
$rememberMeCheckbox.Checked = if ($global:settings.RememberMe) { $global:settings.RememberMe } else { $false }
$tabSettings.Controls.Add($rememberMeCheckbox)

# Save Settings Button
$saveSettingsButton = New-Object System.Windows.Forms.Button
$saveSettingsButton.Text = "💾 Save Settings"
$saveSettingsButton.Location = New-Object System.Drawing.Point(20,180)
$saveSettingsButton.Width = 150
$saveSettingsButton.Height = 35
$tabSettings.Controls.Add($saveSettingsButton)

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
    $global:settings.DefaultPackagePath = $defaultPathTextBox.Text
    $global:settings.AutoCheckUpdates = $autoCheckUpdatesCheckbox.Checked
    $global:settings.RememberMe = $rememberMeCheckbox.Checked
    
    # Update pathBox on WinGet Apps tab with new default
    if ($pathBox) {
      $pathBox.Text = $global:settings.DefaultPackagePath
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

# Hashtable: AppName -> {PackageID, Version}
$global:packageMap = @{}

# Optional: user-chosen versions per PackageID
$global:selectedPackageVersions = @{}

# Cache for winget searches to speed up repeated searches
$global:wingetVersionCache = @{}  # PackageId -> @(versions)
$global:lastUpdateSearch = $null  # Timestamp of last update search

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
  $errMsg = "CRITICAL: Failed to import WinTuner module: $($_.Exception.Message)"
  Write-Log $errMsg
  Update-Status $errMsg
  [System.Windows.Forms.MessageBox]::Show(
    "Failed to import WinTuner module. Please install it manually:\n\nInstall-Module WinTuner -Scope CurrentUser\n\nError: $($_.Exception.Message)",
    "Module Error",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error
  )
}
Update-Status "Module imported."

# Login button
$loginButton = New-Object System.Windows.Forms.Button
$loginButton.Text = "Login to Tenant"
$loginButton.Location = New-Object System.Drawing.Point(570,20)
$loginButton.Width = 180
$form.Controls.Add($loginButton)

# initialize login button enabled state based on username validation
$loginButton.Enabled = (Test-ValidM365UserName -UserName $usernameBox.Text)

$rememberCheckBox = New-Object System.Windows.Forms.CheckBox
$rememberCheckBox.Text = "Remember me"
$rememberCheckBox.Location = New-Object System.Drawing.Point(570,50)
$rememberCheckBox.AutoSize = $true
$rememberCheckBox.Checked = $false
$form.Controls.Add($rememberCheckBox)

$global:settingsPath = Join-Path ([Environment]::GetFolderPath('ApplicationData')) 'WinTunerGUI\settings.json'
$global:settings = @{ 
  RememberMe = $false
  LastUser = ""
  WingetOverrides = @{}
  DefaultPackagePath = "C:\Temp"
  AutoCheckUpdates = $false
}

function Load-Settings {
  try {
    if (Test-Path $global:settingsPath) {
      $o = Get-Content -Path $global:settingsPath -Raw -ErrorAction Stop | ConvertFrom-Json
      if ($o) {
        $global:settings.RememberMe = [bool]$o.RememberMe
        $global:settings.LastUser = [string]$o.LastUser
        
        # New settings with defaults
        if ($o.PSObject.Properties['DefaultPackagePath']) {
          $global:settings.DefaultPackagePath = [string]$o.DefaultPackagePath
        } else {
          $global:settings.DefaultPackagePath = "C:\Temp"
        }
        
        if ($o.PSObject.Properties['AutoCheckUpdates']) {
          $global:settings.AutoCheckUpdates = [bool]$o.AutoCheckUpdates
        } else {
          $global:settings.AutoCheckUpdates = $false
        }
        
        if ($o.PSObject.Properties['WingetOverrides']) {
          # Convert PSCustomObject to hashtable
          $ht = @{}
          foreach ($p in $o.WingetOverrides.PSObject.Properties) { $ht[$p.Name] = [string]$p.Value }
          $global:settings.WingetOverrides = $ht
        } else { $global:settings.WingetOverrides = @{} }
      }
    }
  } catch {
    Write-Log "Warning: Failed to load settings from $($global:settingsPath): $($_.Exception.Message)"
    # Continue with default settings
  }
}

function Save-Settings {
  try {
    $dir = Split-Path -Parent $global:settingsPath
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    ($global:settings | ConvertTo-Json -Compress) | Set-Content -Path $global:settingsPath -Encoding utf8
  } catch {
    Write-Log "Error: Failed to save settings to $($global:settingsPath): $($_.Exception.Message)"
  }
}

Load-Settings
$rememberCheckBox.Checked = [bool]$global:settings.RememberMe
if ($global:settings.RememberMe -and $global:settings.LastUser) { $usernameBox.Text = $global:settings.LastUser } else { $usernameBox.Text = "" }

# Initialize pathBox with saved default package path
if ($pathBox) {
  if ($global:settings.DefaultPackagePath) {
    $pathBox.Text = $global:settings.DefaultPackagePath
  } else {
    $pathBox.Text = "C:\Temp"
  }
}

$rememberCheckBox.Add_CheckedChanged({
  try {
    $global:settings.RememberMe = [bool]$rememberCheckBox.Checked
    if ($global:settings.RememberMe) { $global:settings.LastUser = $usernameBox.Text } else { $global:settings.LastUser = "" }
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
  try {
    Update-Status "Connecting to tenant..."
    $global:isConnected = $false
    $null = Connect-WtWinTuner -Username $usernameBox.Text -ErrorAction Stop
    if (-not (Test-WtConnected)) { throw "Authentication error or failed." }
    $global:isConnected = $true
    Update-Status "Login success."
    $global:currentUserUpn = $usernameBox.Text
    if ($loginInfoLabel) { $loginInfoLabel.Text = "Logged in as: $($global:currentUserUpn)" }
    if ($rememberCheckBox) { $global:settings.RememberMe = [bool]$rememberCheckBox.Checked }
    if ($global:settings.RememberMe) { $global:settings.LastUser = $usernameBox.Text } else { $global:settings.LastUser = "" }
    Save-Settings
    Set-ConnectedUIState -Connected $true
    
    # Auto-check for updates if enabled
    if ($global:settings.AutoCheckUpdates) {
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
    Update-Status ("Login canceled/failed: {0}" -f $_.Exception.Message)
    Set-ConnectedUIState -Connected $false
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
    $global:packageMap.Clear()
    foreach ($result in @($results)) {
      $displayText = "$($result.Name) — $($result.PackageID)"
      [void]$dropdown.Items.Add($displayText)
      $global:packageMap[$displayText] = @{
        PackageID = $result.PackageID
        Version   = $result.Version
      }
    }
    if ($dropdown.Items.Count -gt 0) { $dropdown.SelectedIndex = 0 }
    Update-Status "Search completed."
  } finally {
    $searchButton.Enabled = $true
  }
})

$versionsButton.Add_Click({
  if (-not $dropdown.SelectedItem) { Update-Status "Please select a package first."; return }
  $appName  = $dropdown.SelectedItem
  $package  = $global:packageMap[$appName]
  if (-not $package -or -not $package.PackageID) { Update-Status "Selected item is invalid."; return }
  $packageID = $package.PackageID
  $versions = @(Get-WingetVersions -PackageId $packageID)
  if (-not $versions -or $versions.Count -eq 0) { Update-Status "No versions found for the selected package."; return }
  $chosen = Show-VersionPickerDialog -Title ("Select version for {0}" -f $packageID) -Versions $versions
  if ($chosen) {
    $global:selectedPackageVersions[$packageID] = $chosen
    Update-Status ("Selected version for {0}: {1}" -f $packageID, $chosen)
  } else {
    Update-Status "Version selection canceled."
  }
})

$createButton.Add_Click({
  if (-not $dropdown.SelectedItem) { Update-Status "Please select a package."; return }
  $appName  = $dropdown.SelectedItem
  $package  = $global:packageMap[$appName]
  if (-not $package -or -not $package.PackageID) { Update-Status "Selected item is invalid."; return }
  $packageID = $package.PackageID
  $folder    = $pathBox.Text
  if (-not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
  $filePath  = Join-Path $folder "$packageID.wtpackage"
  
  if (Test-Path $filePath) {
    $res = [System.Windows.Forms.MessageBox]::Show(("A package file already exists:\n{0}\nOverwrite it?" -f $filePath), "Confirm overwrite", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($res -ne [System.Windows.Forms.DialogResult]::Yes) { Update-Status "Creation aborted by user (existing package)."; $uploadButton.Visible = $true; return }
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
    [System.Windows.Forms.Application]::DoEvents()  # Update UI
    
    $desired = $null
    if ($global:selectedPackageVersions.ContainsKey($packageID)) { 
      $desired = $global:selectedPackageVersions[$packageID] 
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
      $uploadButton.Visible = $true
      if ($effectiveVersion) { $global:builtVersions[$packageID] = $effectiveVersion }
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
    if (-not $global:isConnected) {
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
    $package  = $global:packageMap[$appName]
    if (-not $package) { Update-Status "Selected item is invalid."; return }
    $packageID = $package.PackageID
    $version   = $null
    if ($global:builtVersions -and $global:builtVersions.ContainsKey($packageID)) { 
        $version = $global:builtVersions[$packageID] 
    } else { 
        $version = $package.Version 
    }
    if ([string]::IsNullOrWhiteSpace($packageID)) { 
        try { $packageID = ($appName -split '—')[-1].Trim() } catch { } 
    }
    if ([string]::IsNullOrWhiteSpace($version))   { Update-Status "Version could not be determined."; return }
    if ([string]::IsNullOrWhiteSpace($packageID)) { Update-Status "Cannot upload: failed to resolve PackageId."; return }
    $folder = $pathBox.Text
    if (-not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
    
    try {
        $uploadButton.Enabled = $false
        $createButton.Enabled = $false
        
        Update-Status "Uploading $packageID (v$version) to tenant..."
        $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
        $progressBar.MarqueeAnimationSpeed = 30
        $progressBar.Visible = $true
        [System.Windows.Forms.Application]::DoEvents()
        
        Deploy-WtWin32App -PackageId $packageID -Version $version -RootPackageFolder $folder -ErrorAction Stop
        
        Update-Status "Upload completed successfully"
        $uploadButton.Visible = $false
        $appSearchBox.Text = ""
        $dropdown.Items.Clear()
        
        # Clear version cache for this package so updates will use latest version
        if ($global:selectedPackageVersions.ContainsKey($packageID)) {
            $global:selectedPackageVersions.Remove($packageID)
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
  Update-Status "All apps checked ($($updateListBox.Items.Count) items)"
})

$uncheckAllButton.Add_Click({
  for ($i = 0; $i -lt $updateListBox.Items.Count; $i++) {
    $updateListBox.SetItemChecked($i, $false)
  }
  Update-Status "All apps unchecked"
})


# ----------------------------------------------
# Update List Filter - filters as you type
# ----------------------------------------------
$updateFilterBox.Add_TextChanged({
  $filterText = $updateFilterBox.Text.Trim()
  
  # Clear and repopulate list with filtered items
  $updateListBox.Items.Clear()
  
  if ([string]::IsNullOrWhiteSpace($filterText)) {
    # No filter - show all apps
    foreach ($app in @($global:updateApps)) {
      if ($app -and $app.Name) {
        [void]$updateListBox.Items.Add($app.Name)
      }
    }
  } else {
    # Filter apps by name (case-insensitive)
    $filtered = $global:updateApps | Where-Object { 
      $_.Name -like "*$filterText*" 
    }
    foreach ($app in @($filtered)) {
      if ($app -and $app.Name) {
        [void]$updateListBox.Items.Add($app.Name)
      }
    }
  }
  
  # Update status with filter info
  if (-not [string]::IsNullOrWhiteSpace($filterText)) {
    $statusLabel.Text = "Filter: $($updateListBox.Items.Count) apps match '$filterText'"
  }
})


# ----------------------------------------------
# UPDATED: Robust & verbose "Search Updates"
# ----------------------------------------------
$updateSearchButton.Add_Click({
  if (-not $global:isConnected) { 
    Update-Status "Please login to your tenant first."; 
    return 
  }

  try {
    $updateSearchButton.Enabled = $false
    Update-Status "Loading apps from Intune..."
    
    # Reset UI / cache
    $updateFilterBox.Text = ""  # Clear filter
    $updateListBox.Items.Clear()
    $global:updateApps = @()

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

    $candidates = @()
    $processedCount = 0
    $totalCount = $appsToCheck.Count
    
    foreach ($app in $appsToCheck) {
      $processedCount++
      
      # Update progress every app
      try {
        $progressBar.Value = $processedCount
        Update-Status ("Checking ({0}/{1}): {2}" -f $processedCount, $totalCount, $app.Name)
        [System.Windows.Forms.Application]::DoEvents()
      } catch { }
      
      # Try to resolve winget ID
      $wingetId = Try-ResolveWingetIdForApp -App $app
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
                $candidates += $app
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
          $candidates += $app
          Write-Log ("Update available (fallback): {0} ({1} -> {2})" -f $app.Name, $app.CurrentVersion, $app.LatestVersion)
        }
      }
    }

    # 3) Populate dropdown and cache
    $count = 0
    foreach ($app in ($candidates | Sort-Object Name)) {
      if (-not $app -or -not $app.Name) { continue }
      [void]$updateListBox.Items.Add($app.Name)
      $global:updateApps += $app
      $count++
    }

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
    $progressBar.Visible = $false
    $progressBar.Value = 0
  }
})

# -----------------------------
# UPDATED: Update Checked Apps flow
# -----------------------------
$updateSelectedButton.Add_Click({
    # Get checked items
    $checkedApps = @()
    
    Write-Log "Processing $($updateListBox.CheckedItems.Count) checked items from UI"
    Write-Log "global:updateApps cache has $($global:updateApps.Count) apps"
    
    foreach ($itemName in $updateListBox.CheckedItems) {
        Write-Log "Looking for app: '$itemName'"
        
        # Find matching app in cache
        $foundApp = $null
        foreach ($cachedApp in $global:updateApps) {
            if ($cachedApp -and $cachedApp.Name -eq $itemName) {
                $foundApp = $cachedApp
                break
            }
        }
        
        if ($foundApp) { 
            Write-Log "Found: $($foundApp.Name) (Current: $($foundApp.CurrentVersion), Latest: $($foundApp.LatestVersion), GraphId: $($foundApp.GraphId))"
            $checkedApps += $foundApp
        } else {
            Write-Log "WARNING: Could not find '$itemName' in cache!"
            Write-Log "Available apps in cache: $($global:updateApps.Name -join ', ')"
        }
    }
    
    if ($checkedApps.Count -eq 0) { 
        Update-Status "No valid apps found. Try 'Search Updates' again."
        Write-Log "ERROR: 0 apps matched from $($updateListBox.CheckedItems.Count) checked items"
        return 
    }
    
    Write-Log "Successfully matched $($checkedApps.Count) apps for update"
    
    $rootPackageFolder = $pathBox.Text
    if (-not (Test-Path $rootPackageFolder)) { 
        New-Item -ItemType Directory -Path $rootPackageFolder -Force | Out-Null 
    }
    
    try {
        $updateSelectedButton.Enabled = $false
        $updateAllButton.Enabled = $false
        $updateSearchButton.Enabled = $false
        $checkAllButton.Enabled = $false
        $uncheckAllButton.Enabled = $false
        
        Update-Status "Starting update for $($checkedApps.Count) checked apps..."
        $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
        $progressBar.MarqueeAnimationSpeed = 30
        $progressBar.Visible = $true
        
        $successCount = 0
        $failedCount = 0
        $totalCount = $checkedApps.Count
        $currentIndex = 0
        
        # Process each app - extract properties to avoid object passing issues
        foreach ($app in $checkedApps) {
            $currentIndex++
            Update-Status ("Updating ({0}/{1}): {2}" -f $currentIndex, $totalCount, $app.Name)
            [System.Windows.Forms.Application]::DoEvents()
            
            # Extract properties from WtWin32App object
            $appName = $app.Name
            $appCurrentVersion = $app.CurrentVersion
            $appLatestVersion = $app.LatestVersion
            $appGraphId = $app.GraphId
            
            # Get PackageIdentifier - use Try-ResolveWingetIdForApp
            $appPackageId = Try-ResolveWingetIdForApp -App $app
            
            Write-Log "Calling Update-SingleApp with: Name='$appName', Current='$appCurrentVersion', Latest='$appLatestVersion', GraphId='$appGraphId', PackageId='$appPackageId'"
            
            # Call with individual parameters instead of object
            $result = Update-SingleApp `
                -AppName $appName `
                -CurrentVersion $appCurrentVersion `
                -LatestVersion $appLatestVersion `
                -GraphId $appGraphId `
                -PackageIdentifier $appPackageId `
                -RootPackageFolder $rootPackageFolder
            
            if ($result.Success) {
                $successCount++
                Write-Log "Successfully updated: $appName"
            } else {
                $failedCount++
                Write-Log "Failed to update: $appName - $($result.Message)"
            }
        }
        
        Update-Status "Checked apps updated: $successCount successful, $failedCount failed"
        
        # Refresh update list
        try { $updateSearchButton.PerformClick() } catch {}
        
    } catch {
        Update-Status "Update error: $($_.Exception.Message)"
        Write-Log "updateSelectedButton error: $($_.Exception.Message)"
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
})


# -------------------------
# UPDATED: Update All flow
# -------------------------
$updateAllButton.Add_Click({
    $rootPackageFolder = $pathBox.Text
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
        $updateAllButton.Enabled = $false
        $updateSelectedButton.Enabled = $false
        $updateSearchButton.Enabled = $false
        $checkAllButton.Enabled = $false
        $uncheckAllButton.Enabled = $false
        
        Update-Status "Starting mass update for $($updatedApps.Count) apps..."
        $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
        $progressBar.MarqueeAnimationSpeed = 30
        $progressBar.Visible = $true
        
        $successCount = 0
        $failedCount = 0
        $totalCount = $updatedApps.Count
        $currentIndex = 0
        
        foreach ($app in @($updatedApps)) {
            $currentIndex++
            Update-Status ("Updating ({0}/{1}): {2}" -f $currentIndex, $totalCount, $app.Name)
            [System.Windows.Forms.Application]::DoEvents()  # Keep UI responsive
            
            Write-Log "Processing update for: $($app.Name)"
            
            # Execute update workflow using shared function
            $result = Update-SingleApp -App $app -RootPackageFolder $rootPackageFolder
            
            if ($result.Success) {
                $successCount++
                Write-Log "Successfully updated: $($app.Name)"
            } else {
                $failedCount++
                Write-Log "Failed to update: $($app.Name) - $($result.Message)"
            }
        }
        
        Update-Status "All Updates Completed: $successCount successful, $failedCount failed"
        
        # Refresh update list
        try { $updateSearchButton.PerformClick() } catch {}
        
    } catch {
        Update-Status "Mass update error: $($_.Exception.Message)"
        Write-Log "updateAllButton error: $($_.Exception.Message)"
    } finally {
        $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
        $progressBar.Visible = $false
        $progressBar.Value = 0
        $updateAllButton.Enabled = $true
        $updateSelectedButton.Enabled = $true
        $updateSearchButton.Enabled = $true
        $checkAllButton.Enabled = $true
        $uncheckAllButton.Enabled = $true
    }
})


$removeOldAppsButton.Add_Click({
  $oldApps = Get-WtWin32Apps -Superseded $true
  if ($oldApps.Count -eq 0) { Update-Status "No Superseded Apps Found"; return }
  $appNames = ($oldApps | Select-Object -ExpandProperty Name) -join "`r`n"
  $result = [System.Windows.Forms.MessageBox]::Show(
    "The following outdated apps will be removed:`r`n$appNames",
    "Bestätigung",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Question
  )
  if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
    $progressBar.Value = 0
    $progressBar.Visible = $true
    foreach ($app in @($oldApps)) {
      try {
        Remove-WtWin32App -GraphId $app.GraphId -ErrorAction Stop
        Update-Status ("Entfernt: {0}" -f $app.Name)
      } catch {
        Update-Status ("Error removing {0}: {1}" -f $app.Name, $_.Exception.Message)
      }
    }
    $progressBar.Value = 100
    Update-Status "Deleted all superseded Apps..."
    try { $supersededSearchButton.PerformClick() } catch {}
  } else {
    Update-Status "Removal aborted."
  }
})

# Handler: Search superseded apps
$global:supersededApps = @()
$supersededSearchButton.Add_Click({
  try {
    Update-Status "Search for superseded apps..."
    $global:supersededApps = Get-WtWin32Apps -Superseded $true
    $supersededDropdown.Items.Clear()
    foreach ($app in @($global:supersededApps)) {
      $name = $app.Name
      $version = $app.CurrentVersion
      $display = "$name — $version"
      [void]$supersededDropdown.Items.Add($display)
    }
    if ($supersededDropdown.Items.Count -gt 0) { $supersededDropdown.SelectedIndex = 0 }
    Update-Status ("Search completed: {0} superseded Apps found." -f $supersededDropdown.Items.Count)
  } catch {
    Update-Status ("Error while search: {0}" -f $_.Exception.Message)
  }
})

# ==================================================
# Rollback Handlers
# ==================================================
$global:rollbackApps = @()

# Load apps for rollback
$loadRollbackAppsButton.Add_Click({
  try {
    Update-Status "Loading apps with multiple versions..."
    
    # Get all apps (current + superseded)
    $allApps = @(Get-WtWin32Apps -Superseded:$false) + @(Get-WtWin32Apps -Superseded:$true)
    
    # Group by name and filter those with > 1 version
    $appsGrouped = $allApps | Group-Object -Property Name
    $appsWithMultipleVersions = @($appsGrouped | Where-Object { $_.Count -gt 1 } | Select-Object -ExpandProperty Name | Sort-Object)
    
    $rollbackAppDropdown.Items.Clear()
    foreach ($appName in $appsWithMultipleVersions) {
      $versionCount = ($appsGrouped | Where-Object { $_.Name -eq $appName }).Count
      $display = "$appName ($versionCount versions)"
      [void]$rollbackAppDropdown.Items.Add($display)
    }
    
    if ($rollbackAppDropdown.Items.Count -gt 0) {
      $rollbackAppDropdown.SelectedIndex = 0
    } else {
      Update-Status "No apps with multiple versions found."
    }
    
    Update-Status ("Loaded {0} apps with multiple versions." -f $rollbackAppDropdown.Items.Count)
  } catch {
    Update-Status ("Error loading apps: {0}" -f $_.Exception.Message)
  }
})

# Execute rollback - Delete all versions except the selected one
$global:rollbackVersions = @()
$global:allAppVersions = @()

$executeRollbackButton.Add_Click({
  if ($rollbackAppDropdown.SelectedIndex -lt 0) {
    Update-Status "Please select an app first."
    return
  }
  
  if ($rollbackVersionDropdown.SelectedIndex -lt 0) {
    Update-Status "Please select version to keep."
    return
  }
  
  # Extract app name (remove "(X versions)" suffix)
  $selectedText = $rollbackAppDropdown.Text
  $appName = $selectedText -replace ' \(\d+ versions\)$', ''
  $selectedIndex = $rollbackVersionDropdown.SelectedIndex
  $targetVersion = $global:rollbackVersions[$selectedIndex]
  $targetVersionNumber = $targetVersion.CurrentVersion
  
  # If index > 0, it's an old version (needs full rollback)
  # If index = 0, it's the newest version (just cleanup)
  $isOldVersion = ($selectedIndex -gt 0)
  
  Write-Log "Rollback: App='$appName', Target='v$targetVersionNumber', Index=$selectedIndex, IsOld=$isOldVersion"
  
  # Get all other versions
  $allVersions = @($global:rollbackVersions)
  $otherVersions = @($allVersions | Where-Object { $_.GraphId -ne $targetVersion.GraphId })
  
  if ($otherVersions.Count -eq 0) {
    Update-Status "Only one version exists - rollback not needed."
    [System.Windows.Forms.MessageBox]::Show(
      "$appName has only one version (v$targetVersionNumber).`n`nRollback is not needed.",
      "Single Version",
      [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Information
    )
    return
  }
  
  # Confirmation dialog - different message based on old/new
  $confirmMsg = "Rollback: $appName → v$targetVersionNumber`n`n"
  
  if ($isOldVersion) {
    $confirmMsg += "🔄 Target is OLD version (Rollback)`n`n"
    $confirmMsg += "This will:`n"
    $confirmMsg += "1. Delete ALL versions:`n"
    foreach ($v in $allVersions) {
      $confirmMsg += "   - $($v.CurrentVersion)`n"
    }
    $confirmMsg += "2. Re-deploy v$targetVersionNumber`n`n"
    $confirmMsg += "❌ ALL ASSIGNMENTS WILL BE LOST!`n"
    $confirmMsg += "⚠️ You must RE-ASSIGN after rollback!`n`n"
  } else {
    $confirmMsg += "✅ Target is NEWEST version (Cleanup)`n`n"
    $confirmMsg += "This will:`n"
    $confirmMsg += "1. Delete older versions:`n"
    foreach ($v in $otherVersions) {
      $confirmMsg += "   - $($v.CurrentVersion)`n"
    }
    $confirmMsg += "2. Keep v$targetVersionNumber`n`n"
    $confirmMsg += "✅ Assignments will be PRESERVED!`n`n"
  }
  
  $confirmMsg += "Continue?"
  
  $confirm = [System.Windows.Forms.MessageBox]::Show(
    $confirmMsg,
    "Confirm Rollback",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Warning
  )
  
  if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
    Update-Status "Rollback canceled."
    return
  }
  
  try {
    $executeRollbackButton.Enabled = $false
    $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
    $progressBar.Visible = $true
    
    if ($isOldVersion) {
      # Scenario A: Target is old version - REMOVE SUPERSEDENCE and DELETE NEWER
      Update-Status "Rollback to OLD version: Removing supersedence relationship..."
      Write-Log "Rollback: Target is old version - removing supersedence and deleting newer versions"
      
      # Step 1: Remove supersedence relationships to make old version standalone
      $newerVersions = @($allVersions | Where-Object { $_.GraphId -ne $targetVersion.GraphId })
      
      foreach ($newerVer in $newerVersions) {
        try {
          Update-Status "Removing supersedence from v$($newerVer.CurrentVersion)..."
          Write-Log "Removing supersedence relationship for $($newerVer.Name) v$($newerVer.CurrentVersion) (GraphId: $($newerVer.GraphId))"
          
          # Try direct REST API call using Invoke-RestMethod with existing token
          $uri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$($newerVer.GraphId)"
          
          # Get current app details first to preserve properties
          $currentApp = Get-WtWin32Apps | Where-Object { $_.GraphId -eq $newerVer.GraphId } | Select-Object -First 1
          
          if ($currentApp) {
            # Create minimal update body to clear relationships
            $updateBody = @{
              "@odata.type" = "#microsoft.graph.win32LobApp"
              supersedingAppRelationships = @()
              supersededAppRelationships = @()
            } | ConvertTo-Json -Depth 10
            
            # Get auth token from WinTuner session
            try {
              $token = Get-MgContext | Select-Object -ExpandProperty AuthType
              if (-not $token) {
                throw "No auth token available"
              }
              
              $headers = @{
                "Authorization" = "Bearer $token"
                "Content-Type" = "application/json"
              }
              
              Invoke-RestMethod -Uri $uri -Method PATCH -Headers $headers -Body $updateBody -ErrorAction Stop
              Write-Log "Supersedence relationship removed from v$($newerVer.CurrentVersion) via REST API"
              
            } catch {
              Write-Log "Direct REST API failed: $($_.Exception.Message)"
              Write-Log "Supersedence removal may have failed - will try to delete anyway"
            }
          } else {
            Write-Log "Could not find app details for GraphId $($newerVer.GraphId)"
          }
          
        } catch {
          Write-Log "Could not remove supersedence from v$($newerVer.CurrentVersion): $($_.Exception.Message)"
          Write-Log "Will attempt deletion anyway (may fail if still parent)"
        }
      }
      
      # Small delay to let relationship updates propagate
      Start-Sleep -Seconds 2
      
      # Step 2: Delete ALL versions in correct order (oldest to newest = children before parents)
      # This includes the target version temporarily, then we'll redeploy it
      Update-Status "Deleting all versions in correct order..."
      Write-Log "Since supersedence removal may have failed, deleting all versions and redeploying target"
      
      # Sort all versions: oldest first (so children get deleted before parents)
      $allVersionsSorted = @($allVersions | Sort-Object { try { [version]$_.CurrentVersion } catch { $_.CurrentVersion } })
      
      $deletedCount = 0
      $failedCount = 0
      
      foreach ($ver in $allVersionsSorted) {
        try {
          Update-Status "Deleting v$($ver.CurrentVersion)..."
          Remove-WtWin32App -GraphId $ver.GraphId -ErrorAction Stop
          Write-Log "Deleted: v$($ver.CurrentVersion)"
          $deletedCount++
        } catch {
          Write-Log "Failed to delete v$($ver.CurrentVersion): $($_.Exception.Message)"
          $failedCount++
        }
      }
      
      # Step 3: Re-deploy the target version
      if ($deletedCount -gt 0) {
        Update-Status "Re-deploying target version v$targetVersionNumber..."
        Write-Log "Re-deploying $appName v$targetVersionNumber"
        
        try {
          $wingetId = Try-ResolveWingetIdForApp -App $targetVersion
          if ([string]::IsNullOrWhiteSpace($wingetId)) {
            throw "Cannot determine WingetId for $appName"
          }
          
          $rootFolder = $pathBox.Text
          $resPkg = New-WingetPackageWithFallback `
            -PackageId $wingetId `
            -PackageFolder $rootFolder `
            -DesiredVersion $targetVersionNumber `
            -LatestVersion $targetVersionNumber `
            -InstalledVersion "" `
            -ErrorAction Stop
          
          if (-not $resPkg -or -not $resPkg.Succeeded) {
            throw "Package creation failed"
          }
          
          Deploy-WtWin32App `
            -PackageId $wingetId `
            -Version $targetVersionNumber `
            -RootPackageFolder $rootFolder `
            -ErrorAction Stop
          
          Write-Log "Target version v$targetVersionNumber re-deployed successfully"
          
        } catch {
          Write-Log "ERROR: Failed to re-deploy v${targetVersionNumber}: $($_.Exception.Message)"
          throw "Rollback partially completed but re-deployment failed. You may need to manually deploy v$targetVersionNumber"
        }
      }
      
      Update-Status "Rollback completed: $appName v$targetVersionNumber deployed"
      Write-Log "Rollback completed: Deleted $deletedCount version(s), $failedCount failed, re-deployed v$targetVersionNumber"
      
      [System.Windows.Forms.MessageBox]::Show(
        "Rollback successful!`n`n$appName v$targetVersionNumber has been deployed.`n`nDeleted: $deletedCount version(s)`nFailed: $failedCount`n`n⚠️ You need to RE-ASSIGN this app to groups!",
        "Rollback Complete - Re-assign Required",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
      )
      
    } else {
      # Scenario B: Target is current (new) - DELETE OTHERS, KEEP TARGET
      Update-Status "Keeping v$targetVersionNumber, deleting others..."
      Write-Log "Rollback: Target is current - deleting $($otherVersions.Count) other versions, keeping v$targetVersionNumber"
      
      $deletedCount = 0
      $failedCount = 0
      
      # Delete superseded first (children), then current (parents)
      $supersededToDelete = @($otherVersions | Where-Object { $_.Superseded })
      $currentToDelete = @($otherVersions | Where-Object { -not $_.Superseded })
      
      foreach ($ver in $supersededToDelete) {
        try {
          Update-Status "Deleting old version $($ver.CurrentVersion)..."
          Remove-WtWin32App -GraphId $ver.GraphId -ErrorAction Stop
          Write-Log "Deleted: $($ver.CurrentVersion) (superseded)"
          $deletedCount++
        } catch {
          Write-Log "Failed to delete $($ver.CurrentVersion): $($_.Exception.Message)"
          $failedCount++
        }
      }
      
      foreach ($ver in $currentToDelete) {
        try {
          Update-Status "Deleting version $($ver.CurrentVersion)..."
          Remove-WtWin32App -GraphId $ver.GraphId -ErrorAction Stop
          Write-Log "Deleted: $($ver.CurrentVersion) (current)"
          $deletedCount++
        } catch {
          Write-Log "Failed to delete $($ver.CurrentVersion): $($_.Exception.Message)"
          $failedCount++
        }
      }
      
      Update-Status "Rollback completed: $appName v$targetVersionNumber is now the only version"
      Write-Log "Rollback completed: Deleted $deletedCount, failed $failedCount, kept v$targetVersionNumber"
      
      [System.Windows.Forms.MessageBox]::Show(
        "Rollback successful!`n`n$appName v$targetVersionNumber is now the only version.`n`nDeleted: $deletedCount`nFailed: $failedCount`n`n✅ Assignments preserved!",
        "Rollback Complete",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
      )
    }
    
    # Refresh lists
    $rollbackAppDropdown.Items.Clear()
    $rollbackVersionDropdown.Items.Clear()
    $executeRollbackButton.Enabled = $false
    
  } catch {
    $errorMsg = $_.Exception.Message
    Update-Status "Rollback failed: $errorMsg"
    Write-Log "Rollback error: $errorMsg"
    
    [System.Windows.Forms.MessageBox]::Show(
      "Rollback failed!`n`n$errorMsg",
      "Rollback Error",
      [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Error
    )
  } finally {
    $progressBar.Visible = $false
    $executeRollbackButton.Enabled = $true
  }
})

# Handler: Delete selected superseded app
$deleteSelectedAppButton.Add_Click({
  if (-not $global:supersededApps -or $supersededDropdown.SelectedIndex -lt 0) {
    Update-Status "Please first select a superseded app from the dropdown."
    return
  }
  $app = $global:supersededApps[$supersededDropdown.SelectedIndex]
  $result = [System.Windows.Forms.MessageBox]::Show(
    "Delete App '" + $app.Name + "'?",
    "Bestätigung",
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
  Disconnect-WtWinTuner
  $global:isConnected = $false
  $global:currentUserUpn = ""
  if ($loginInfoLabel) { $loginInfoLabel.Text = "" }
  Update-Status "Logout success."
  Set-ConnectedUIState -Connected $false
})

# Apply initial theme (Dark by default)
Apply-Theme -control $form -theme $global:currentTheme

# Safe logger for closing context
function Write-FileLog {
  param([string]$message)
  try { Write-Log $message } catch {}
}

# Re-entrancy protection for closing
$global:_closingInProgress = $false
$form.Add_FormClosing({
  param($sender, [System.Windows.Forms.FormClosingEventArgs]$e)
  try { if ($global:settings) { if ($global:settings.RememberMe) { $global:settings.LastUser = $usernameBox.Text } else { $global:settings.LastUser = "" }; Save-Settings } } catch {}
  if ($global:_closingInProgress) { return }
  if (-not $global:isConnected) { return }

  $e.Cancel = $true
  $global:_closingInProgress = $true
  try {
    $form.Enabled = $false
    if ($statusLabel) { $statusLabel.Text = "Closing... signing out from tenant" }
  } catch {}

  Write-FileLog 'Shutdown: starting tenant disconnect (BackgroundWorker).'

  # 5s Timeout on UI thread
  $script:_wtCloseTimer = New-Object System.Windows.Forms.Timer
  $script:_wtCloseTimer.Interval = 5000
  $script:_wtCloseTimer.Add_Tick({
    $script:_wtCloseTimer.Stop()
    try { Write-FileLog 'Shutdown: disconnect timeout after 5s. Forcing close.' } catch {}
    try { $form.Close() } catch {}
  })
  $script:_wtCloseTimer.Start()

  # Disconnect in background
  $bw = New-Object System.ComponentModel.BackgroundWorker
  $bw.WorkerSupportsCancellation = $false
  $bw.Add_DoWork({ 
    try { 
      Disconnect-WtWinTuner 
    } catch { 
      try { Write-FileLog "Warning: Disconnect-WtWinTuner failed: $($_.Exception.Message)" } catch {}
    } 
  })
  $bw.Add_RunWorkerCompleted({
    try { Write-FileLog 'Shutdown: disconnect finished. Closing form.' } catch {}
    try { if ($script:_wtCloseTimer) { $script:_wtCloseTimer.Stop() } } catch {}
    $global:isConnected = $false
    try { $form.Close() } catch {}
  })
  $bw.RunWorkerAsync()
})

# Run the form
[System.Windows.Forms.Application]::Run($form)
