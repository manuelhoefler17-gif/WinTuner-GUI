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

$__WinTunerMain = @'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Enable visual styles BEFORE creating controls
[System.Windows.Forms.Application]::EnableVisualStyles()

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
  try { $output = & winget show --id $PackageId --versions 2>$null } catch { return @() }
  if (-not $output) { return @() }
  $cand = @()
  foreach ($line in @($output)) {
    $t = ($line -replace '^[\s\-•]+','').Trim()
    if (-not $t) { continue }
    if ($t -match '^(\d+)(\.[0-9A-Za-z]+)*([\-+._][0-9A-Za-z]+)*$') { $cand += $t }
  }
  $unique = @($cand | Select-Object -Unique)
  $parsed = foreach ($v in $unique) { $ok = $false; $vo = $null; try { $vo = [version]$v; $ok = $true } catch {}; [pscustomobject]@{ Text = $v; Parsed = $vo; Numeric = $ok } }
  if ($parsed | Where-Object Numeric) { return @($parsed | Where-Object Numeric | Sort-Object Parsed -Descending | Select-Object -ExpandProperty Text) }
  else { return @($parsed | Sort-Object Text -Descending | Select-Object -ExpandProperty Text) }
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
      try { $app.LatestVersion = $usedLatest } catch {}
      $candidates += $app
    }
  }
  return ,$candidates
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

# Logging function
function Write-Log {
  param([string]$message)
  $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
  $base = if ($PSScriptRoot) { $PSScriptRoot } elseif ($MyInvocation -and $MyInvocation.MyCommand -and $MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path } else { (Get-Location).Path }
  $logPath = Join-Path $base 'WinTuner_GUI.log'
  Add-Content -Path $logPath -Value "$timestamp - $message" -Encoding utf8
  if ($outputBox) { $outputBox.AppendText("$timestamp - $message`r`n") }
}

# Status update function
function Update-Status {
  param([string]$status)
  $statusLabel.Text = $status
  Write-Log $status
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
$form.Size = New-Object System.Drawing.Size(900, 700)
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
$statusLabel.Location = New-Object System.Drawing.Point(10, 620)
$statusLabel.Width = 750
$form.Controls.Add($statusLabel)

# Output textbox
$outputBox = New-Object System.Windows.Forms.TextBox
$outputBox.Location = New-Object System.Drawing.Point(10, 520)
$outputBox.Size = New-Object System.Drawing.Size(760, 90)
$outputBox.Multiline = $true
$outputBox.ScrollBars = "Vertical"
$outputBox.ReadOnly = $true
$form.Controls.Add($outputBox)

# Progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(100, 490)
$progressBar.Width = 650
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
$tabControl.Size = New-Object System.Drawing.Size(760, 420)
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

$updateDropdown = New-Object System.Windows.Forms.ComboBox
$updateDropdown.Location = New-Object System.Drawing.Point(100,90)
$updateDropdown.Width = 650
$updateDropdown.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$tabUpdate.Controls.Add($updateDropdown)

$updateSelectedButton = New-Object System.Windows.Forms.Button
$updateSelectedButton.Text = "Update selected apps"
$updateSelectedButton.Location = New-Object System.Drawing.Point(100,130)
$updateSelectedButton.Width = 250
$updateSelectedButton.Enabled = $false
$tabUpdate.Controls.Add($updateSelectedButton)

$updateAllButton = New-Object System.Windows.Forms.Button
$updateAllButton.Text = "Update all apps"
$updateAllButton.Location = New-Object System.Drawing.Point(420,130)
$updateAllButton.Width = 250
$updateAllButton.Enabled = $false
$tabUpdate.Controls.Add($updateAllButton)

# Label over "Search Superseded Apps"
$supersededHeaderLabel = New-Object System.Windows.Forms.Label
$supersededHeaderLabel.Text = "Search for superseded Apps"
$supersededHeaderLabel.Location = New-Object System.Drawing.Point(100,170)
$supersededHeaderLabel.AutoSize = $true
$supersededHeaderLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$tabUpdate.Controls.Add($supersededHeaderLabel)

$supersededSearchButton = New-Object System.Windows.Forms.Button
$supersededSearchButton.Text = "Search Superseded Apps"
$supersededSearchButton.Location = New-Object System.Drawing.Point(100,200)
$supersededSearchButton.Width = 250
$supersededSearchButton.Enabled = $false
$tabUpdate.Controls.Add($supersededSearchButton)

# Dropdown: Superseded Apps (Name + Version)
$supersededDropdown = New-Object System.Windows.Forms.ComboBox
$supersededDropdown.Location = New-Object System.Drawing.Point(100,240)
$supersededDropdown.Width = 650
$supersededDropdown.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$tabUpdate.Controls.Add($supersededDropdown)

# Button: Delete selected app
$deleteSelectedAppButton = New-Object System.Windows.Forms.Button
$deleteSelectedAppButton.Text = "Ausgewählte App löschen"
$deleteSelectedAppButton.Location = New-Object System.Drawing.Point(100,280)
$deleteSelectedAppButton.Width = 250
$deleteSelectedAppButton.Enabled = $false
$tabUpdate.Controls.Add($deleteSelectedAppButton)

$removeOldAppsButton = New-Object System.Windows.Forms.Button
$removeOldAppsButton.Text = "Delete all Superseded Apps"
$removeOldAppsButton.Location = New-Object System.Drawing.Point(360,280)
$removeOldAppsButton.Width = 250
$removeOldAppsButton.Enabled = $false
$tabUpdate.Controls.Add($removeOldAppsButton)

# Hashtable: AppName -> {PackageID, Version}
$global:packageMap = @{}

# Optional: user-chosen versions per PackageID
$global:selectedPackageVersions = @{}

# Module check
Update-Status "Checking WinTuner Module..."
try {
  if (Get-Module -ListAvailable -Name WinTuner) {
    Update-Status "Module found, searching for updates..."
    try { Update-Module -Name WinTuner -ErrorAction Stop } catch {}
  } else {
    Update-Status "Module not found, installing..."
    try { Install-Module -Name WinTuner -Scope CurrentUser -Repository PSGallery -Force -ErrorAction Stop } catch { Update-Status ("Module install failed: {0}" -f $_.Exception.Message) }
  }
} catch {
  Update-Status ("Module install/update error: {0}" -f $_.Exception.Message)
}
try { Import-Module WinTuner -ErrorAction Stop } catch {}
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
$global:settings = @{ RememberMe = $false; LastUser = ""; WingetOverrides = @{} }

function Load-Settings {
  try {
    if (Test-Path $global:settingsPath) {
      $o = Get-Content -Path $global:settingsPath -Raw -ErrorAction Stop | ConvertFrom-Json
      if ($o) {
        $global:settings.RememberMe = [bool]$o.RememberMe
        $global:settings.LastUser = [string]$o.LastUser
        if ($o.PSObject.Properties['WingetOverrides']) {
          # Convert PSCustomObject to hashtable
          $ht = @{}
          foreach ($p in $o.WingetOverrides.PSObject.Properties) { $ht[$p.Name] = [string]$p.Value }
          $global:settings.WingetOverrides = $ht
        } else { $global:settings.WingetOverrides = @{} }
      }
    }
  } catch {}
}

function Save-Settings {
  try {
    $dir = Split-Path -Parent $global:settingsPath
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    ($global:settings | ConvertTo-Json -Compress) | Set-Content -Path $global:settingsPath -Encoding utf8
  } catch {}
}

Load-Settings
$rememberCheckBox.Checked = [bool]$global:settings.RememberMe
if ($global:settings.RememberMe -and $global:settings.LastUser) { $usernameBox.Text = $global:settings.LastUser } else { $usernameBox.Text = "" }

$rememberCheckBox.Add_CheckedChanged({
  try {
    $global:settings.RememberMe = [bool]$rememberCheckBox.Checked
    if ($global:settings.RememberMe) { $global:settings.LastUser = $usernameBox.Text } else { $global:settings.LastUser = "" }
    Save-Settings
  } catch {}
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
    if ($loginInfoLabel) { $loginInfoLabel.Text = "Angemeldet als: $($global:currentUserUpn)" }
    if ($rememberCheckBox) { $global:settings.RememberMe = [bool]$rememberCheckBox.Checked }
    if ($global:settings.RememberMe) { $global:settings.LastUser = $usernameBox.Text } else { $global:settings.LastUser = "" }
    Save-Settings
    Set-ConnectedUIState -Connected $true
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
    try { Remove-Item -Path $filePath -Force -ErrorAction Stop } catch {}
  }

  try {
    $createButton.Enabled = $false
    Update-Status "Creating package..."
    $progressBar.Value = 0
    $progressBar.Visible = $true
    $progressBar.Value = 30
    Start-Sleep -Milliseconds 300
    $desired = $null; if ($global:selectedPackageVersions.ContainsKey($packageID)) { $desired = $global:selectedPackageVersions[$packageID] }
    $resPkg = New-WingetPackageWithFallback -PackageId $packageID -PackageFolder $folder -DesiredVersion $desired -LatestVersion $package.Version -AllowUserRetry -ErrorAction SilentlyContinue
    if ($resPkg -and $resPkg.Succeeded) {
      $progressBar.Value = 100
      $effectiveVersion = $resPkg.EffectiveVersion
      if (-not $effectiveVersion) { $effectiveVersion = $package.Version }
      Update-Status ("Success creating package (version {0})." -f $effectiveVersion)
      $uploadButton.Visible = $true
      if ($effectiveVersion) { $global:builtVersions[$packageID] = $effectiveVersion }
    } else {
      return
    }
  } finally {
    $createButton.Enabled = $true
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
    if ($global:builtVersions -and $global:builtVersions.ContainsKey($packageID)) { $version = $global:builtVersions[$packageID] } else { $version = $package.Version }
    if ([string]::IsNullOrWhiteSpace($packageID)) { try { $packageID = ($appName -split '—')[-1].Trim() } catch { } }
    if ([string]::IsNullOrWhiteSpace($version))   { Update-Status "Version could not be determined."; return }
    if ([string]::IsNullOrWhiteSpace($packageID)) { Update-Status "Cannot upload: failed to resolve PackageId."; return }
    $folder = $pathBox.Text
    if (-not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
    try {
        $uploadButton.Enabled = $false
        $progressBar.Value = 0; $progressBar.Visible = $true
        Update-Status "Starting Upload..."
        Start-Sleep -Milliseconds 300
        Deploy-WtWin32App -PackageId $packageID -Version $version -RootPackageFolder $folder -ErrorAction Stop
        $progressBar.Value = 100
        Update-Status "Upload completed."
        $uploadButton.Visible = $false
        $appSearchBox.Text = ""
        $dropdown.Items.Clear()
    } finally { $uploadButton.Enabled = $true }
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
    Update-Status "Searching for updates..."
    
    # Reset UI / cache
    $updateDropdown.Items.Clear()
    $global:updateApps = @()

    # 1) Load all apps, then verify latest via winget where possible to avoid false positives/negatives
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
      Write-Log ("Resolved winget id for {0}: {1}" -f $app.Name, ($wingetId ? $wingetId : '<none>'))
      $verified = $false
      if ($wingetId) {
        try {
          $wgVersions = @(Get-WingetVersions -PackageId $wingetId)
        } catch { $wgVersions = @() }
        if ($wgVersions -and $wgVersions.Count -gt 0) {
          $wgLatest = $wgVersions[0]
          if ($wgLatest) {
            # overwrite LatestVersion with authoritative winget latest
            try { $app.LatestVersion = $wgLatest } catch {}
            if (Test-IsNewerVersion $wgLatest $app.CurrentVersion) {
              $candidates += $app
            }
            $verified = $true
          }
        }
      }
      if (-not $verified) {
        # fallback to manual comparison only when we could not verify via winget for this app
        if ($app.LatestVersion -and (Test-IsNewerVersion $app.LatestVersion $app.CurrentVersion)) { $candidates += $app }
      }
    }

    # 3) Populate dropdown and cache
    $count = 0
    foreach ($app in ($candidates | Sort-Object Name)) {
      if (-not $app -or -not $app.Name) { continue }
      [void]$updateDropdown.Items.Add($app.Name)
      $global:updateApps += $app
      $count++
    }

    if ($count -gt 0) {
      $updateDropdown.SelectedIndex = 0
      Update-Status ("Search updates completed: {0} candidate(s) found." -f $count)
    } else {
      Update-Status "No update candidates found."
    }
  } finally {
    $updateSearchButton.Enabled = $true
  }
})

# -----------------------------
# UPDATED: Update Selected flow
# -----------------------------
$updateSelectedButton.Add_Click({
    $selectedAppName = $updateDropdown.SelectedItem
    if (-not $selectedAppName) { Update-Status "No app selected."; return }
    $app = $global:updateApps | Where-Object { $_.Name -eq $selectedAppName }
    $rootPackageFolder = $pathBox.Text
    if (-not $app) { Update-Status "No valid app selected."; return }
    if (-not (Test-Path $rootPackageFolder)) { New-Item -ItemType Directory -Path $rootPackageFolder -Force | Out-Null }
    try {
        $updateSelectedButton.Enabled = $false
        $wingetId = Try-ResolveWingetIdForApp -App $app
        Write-Log ("Resolved winget id for {0}: {1}" -f $app.Name, ($wingetId ? $wingetId : '<none>'))
        if ([string]::IsNullOrWhiteSpace($wingetId)) { Update-Status ("Cannot determine PackageId for '{0}'." -f $app.Name); return }
        Update-Status "Update for $($app.Name) is running..."
        $effectiveVersion = $null
        $progressBar.Value = 0; $progressBar.Visible = $true; $progressBar.Value = 30; Start-Sleep -Milliseconds 300
        # 1) Create/refresh package using reusable fallback logic
        $desired = $app.LatestVersion
        try { if ($global:selectedPackageVersions.ContainsKey($wingetId)) { $desired = $global:selectedPackageVersions[$wingetId] } } catch {}
        $resPkg = New-WingetPackageWithFallback -PackageId $wingetId -PackageFolder $rootPackageFolder -DesiredVersion $desired -LatestVersion $app.LatestVersion -InstalledVersion $app.CurrentVersion -AllowUserRetry -ErrorAction SilentlyContinue
        if (-not $resPkg -or -not $resPkg.Succeeded) { Update-Status ("Package creation failed for {0}." -f $app.Name); return }
        $effectiveVersion = if ($resPkg.EffectiveVersion) { $resPkg.EffectiveVersion } else { $app.LatestVersion }
        # 2) Deploy with best available identifier
        $deploySplat = @{ RootPackageFolder = $rootPackageFolder; ErrorAction = 'Stop' }
        if ($app.GraphId) {
            $deploySplat.GraphId = $app.GraphId
            $deploySplat.KeepAssignments = $true
            $deploySplat.PackageId = $wingetId
            $deploySplat.Version   = $effectiveVersion
            Update-Status "Deploying update by GraphId (+ PackageId/Version) ($($app.GraphId))..."
        } else {
            $deploySplat.PackageId = $wingetId
            $deploySplat.Version   = $effectiveVersion
            Update-Status "Deploying by PackageId ($wingetId) version $($effectiveVersion)..."
        }
        Deploy-WtWin32App @deploySplat
        $progressBar.Value = 100
        Update-Status "Update completed."
        try { $updateSearchButton.PerformClick() } catch {}
    } catch {
        Update-Status ("Deploy failed: {0}" -f $_.Exception.Message)
    } finally {
        $updateSelectedButton.Enabled = $true
    }
})


# -------------------------
# UPDATED: Update All flow
# -------------------------
$updateAllButton.Add_Click({
    $rootPackageFolder = $pathBox.Text
    if (-not (Test-Path $rootPackageFolder)) { New-Item -ItemType Directory -Path $rootPackageFolder -Force | Out-Null }
    # Build candidate list to show in confirmation dialog
    try { $updatedApps = @(Get-WtWin32Apps -Update $true -Superseded $false) } catch { Write-Verbose "Get-WtWin32Apps update threw: $($_)"; $updatedApps = @() }
    $updatedApps = @(( $updatedApps | Where-Object { $_.LatestVersion -and $_.CurrentVersion -and (Test-IsNewerVersion $_.LatestVersion $_.CurrentVersion) } | Sort-Object Name ))
    if (-not $updatedApps -or $updatedApps.Count -eq 0) { Update-Status "No update candidates found."; return }
    $appNames = ($updatedApps | Select-Object -ExpandProperty Name) -join "`r`n"
    $confirm = [System.Windows.Forms.MessageBox]::Show("The following apps will be updated:`r`n$appNames", "Confirm", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { Update-Status "Mass update canceled."; return }
    try {
        $updateAllButton.Enabled = $false
        Update-Status "Starting mass update..."
        $progressBar.Value = 0; $progressBar.Visible = $true; $progressBar.Value = 10
        foreach ($app in @($updatedApps)) {
            try {
                $wingetId = Try-ResolveWingetIdForApp -App $app
                Write-Log ("Resolved winget id for {0}: {1}" -f $app.Name, ($wingetId ? $wingetId : '<none>'))
                if ([string]::IsNullOrWhiteSpace($wingetId)) { Update-Status ("Update skipped for {0}: missing PackageId/GraphId" -f $app.Name); continue }
                Update-Status ("Update: {0}" -f $app.Name)
                Start-Sleep -Milliseconds 200
                # 1) Create/refresh package using reusable fallback logic
                $desired = $app.LatestVersion
                try { if ($global:selectedPackageVersions.ContainsKey($wingetId)) { $desired = $global:selectedPackageVersions[$wingetId] } } catch {}
                $resPkg = New-WingetPackageWithFallback -PackageId $wingetId -PackageFolder $rootPackageFolder -DesiredVersion $desired -LatestVersion $app.LatestVersion -InstalledVersion $app.CurrentVersion -AllowUserRetry -ErrorAction SilentlyContinue
                if (-not $resPkg -or -not $resPkg.Succeeded) { Update-Status ("Package creation failed for {0}. Skipping." -f $app.Name); continue }
                $effectiveVersion = if ($resPkg.EffectiveVersion) { $resPkg.EffectiveVersion } else { $app.LatestVersion }
                # 2) Deploy appropriately
                $deploySplat = @{ RootPackageFolder = $rootPackageFolder; ErrorAction = 'Stop' }
                if ($app.GraphId) {
                    $deploySplat.GraphId = $app.GraphId
                    $deploySplat.KeepAssignments = $true
                    $deploySplat.PackageId = $wingetId
                    $deploySplat.Version   = $effectiveVersion
                } else {
                    $deploySplat.PackageId = $wingetId
                    $deploySplat.Version   = $effectiveVersion
                }
                Deploy-WtWin32App @deploySplat
            } catch {
                Update-Status ("Update failed for {0}: {1}" -f $app.Name, $_.Exception.Message)
            }
        }
        $progressBar.Value = 100
        Update-Status "All Updates Completed"
        try { $updateSearchButton.PerformClick() } catch {}
    } finally {
        $updateAllButton.Enabled = $true
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
  $bw.Add_DoWork({ try { Disconnect-WtWinTuner } catch {} })
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
'@
Invoke-Expression $__WinTunerMain
