#WinTuner GUI by Manuel Höfler
# --- PowerShell version gate (runs on PS5 without parsing the main body) ---
try { $psMajor = $PSVersionTable.PSVersion.Major } catch { $psMajor = 0 }
if ($psMajor -lt 7) {
    try { Add-Type -AssemblyName PresentationFramework -ErrorAction Stop } catch {}
    if ([System.AppDomain]::CurrentDomain.GetAssemblies().Location -match 'PresentationFramework') {
        [void][System.Windows.MessageBox]::Show(
            "This script requires PowerShell 7 or higher. Please upgrade your PowerShell version to continue.",
            "PowerShell Version Error",
            "OK",
            "Error"
        )
    } else {
        Write-Host "This script requires PowerShell 7 or higher. Please upgrade your PowerShell version to continue." -ForegroundColor Red
    }
    return
}


$__WinTunerMain = @'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# Version comparison helper: returns $true if Latest > Current
function Test-IsNewerVersion {
    param([string]$Latest, [string]$Current)
    if (-not $Latest -or -not $Current) { return $false }
    try {
        $vL = [version]$Latest
        $vC = [version]$Current
        return $vL -gt $vC
    } catch {
        # Fallback: compare numeric segments
        $numsL = ($Latest -split '[^0-9]+' | Where-Object { $_ -ne '' }) | ForEach-Object { [int]$_ }
        $numsC = ($Current -split '[^0-9]+' | Where-Object { $_ -ne '' }) | ForEach-Object { [int]$_ }
        $len = [Math]::Max(($numsL | Measure-Object).Count, ($numsC | Measure-Object).Count)
        for ($i=0; $i -lt $len; $i++) {
            $a = if ($i -lt $numsL.Count) { $numsL[$i] } else { 0 }
            $b = if ($i -lt $numsC.Count) { $numsC[$i] } else { 0 }
            if ($a -gt $b) { return $true }
            if ($a -lt $b) { return $false }
        }
        # Same numeric skeleton; consider different strings as newer to avoid false negatives
        return $Latest -ne $Current
    }
}
# Dark mode theme colors
$global:darkTheme = @{
    BackColor = [System.Drawing.Color]::FromArgb(32, 32, 32)
    ForeColor = [System.Drawing.Color]::FromArgb(255, 255, 255)
    ButtonBackColor = [System.Drawing.Color]::FromArgb(64, 64, 64)
    ButtonForeColor = [System.Drawing.Color]::FromArgb(255, 255, 255)
    TextBoxBackColor = [System.Drawing.Color]::FromArgb(48, 48, 48)
    TextBoxForeColor = [System.Drawing.Color]::FromArgb(255, 255, 255)
    TabBackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
    TabForeColor = [System.Drawing.Color]::FromArgb(255, 255, 255)
}
# Light mode theme colors
$global:lightTheme = @{
    BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    ForeColor = [System.Drawing.Color]::FromArgb(0, 0, 0)
    ButtonBackColor = [System.Drawing.Color]::FromArgb(225, 225, 225)
    ButtonForeColor = [System.Drawing.Color]::FromArgb(0, 0, 0)
    TextBoxBackColor = [System.Drawing.Color]::FromArgb(255, 255, 255)
    TextBoxForeColor = [System.Drawing.Color]::FromArgb(0, 0, 0)
    TabBackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    TabForeColor = [System.Drawing.Color]::FromArgb(0, 0, 0)
}
$global:isDarkMode = $true
$global:currentTheme = $global:darkTheme
# Function to apply theme to all controls
function Apply-Theme {
    param([System.Windows.Forms.Control]$control, [hashtable]$theme)
    # Apply theme to the control itself
    if ($control -is [System.Windows.Forms.Form]) {
        $control.BackColor = $theme.BackColor
        $control.ForeColor = $theme.ForeColor
    }
    elseif ($control -is [System.Windows.Forms.Button]) {
        $control.BackColor = $theme.ButtonBackColor
        $control.ForeColor = $theme.ButtonForeColor
        $control.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $control.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
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
        # ProgressBar styling is limited in WinForms
        $control.BackColor = $theme.TextBoxBackColor
    }
    else {
        $control.BackColor = $theme.BackColor
        $control.ForeColor = $theme.ForeColor
    }
    # Recursively apply theme to child controls
    foreach ($childControl in @($control.Controls)) {
        Apply-Theme -control $childControl -theme $theme
    }
}
# Function to toggle theme
function Toggle-Theme {
    $global:isDarkMode = -not $global:isDarkMode
    $global:currentTheme = if ($global:isDarkMode) { $global:darkTheme } else { $global:lightTheme }
    Apply-Theme -control $form -theme $global:currentTheme
    # Update toggle button text
    $themeToggleButton.Text = if ($global:isDarkMode) { "Light Mode" } else { "Dark Mode" }
    # Force refresh
    $form.Refresh()
}
# Logging function
function Write-Log {
    param([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path "WinTuner_GUI.log" -Value "$timestamp - $message"
    $outputBox.AppendText("$timestamp - $message`r`n")
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
    $upnRegex = '^(?=.{3,256}$)(?![.])(?!.*[.]{2})[A-Za-z0-9._%+\-]+@(?:[A-Za-z0-9\-]+\.)+[A-Za-z]{2,}$'
    return ($UserName -match $upnRegex)
}
# Helper: check if WinTuner is connected (simple smoke test)
function Test-WtConnected {
    try {
        $null = Get-WtWin32Apps -Update:$false -Superseded:$false -ErrorAction Stop | Select-Object -First 1 | Out-Null
        return $true
    } catch {
        return $false
    }
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
        $tabControl.Visible = $false
        $logoutButton.Visible = $false
    }
}
$global:isConnected = $false
# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = "WinTuner GUI"
$form.Size = New-Object System.Drawing.Size(900, 700)
$form.Padding = '5,5,5,5'
# Theme toggle button (top right)
$themeToggleButton = New-Object System.Windows.Forms.Button
$themeToggleButton.Text = "Dark Mode"
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
# ENTER im Username-Feld -> "Login"
if ($usernameBox -ne $null) {
    $usernameBox.Add_KeyDown({
        param($sender, $e)
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            if ($loginButton -and $loginButton.Enabled) {
                $loginButton.PerformClick()
            } else {
                [System.Windows.Forms.MessageBox]::Show("Please enter a valid M365 UPN, z.B. name@firma.de")
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
        $usernameError.Text = "Please enter a valid M365 UPN, z.B. name@firma.de"
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
# Fortschrittsbalken
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
# TabControl
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(10, 60)
$tabControl.Size = New-Object System.Drawing.Size(760, 420)
$tabControl.Visible = $false
$form.Controls.Add($tabControl)
# Tab: Paket erstellen
$tabCreate = New-Object System.Windows.Forms.TabPage
$tabCreate.Text = "WinGet Apps"
$tabControl.TabPages.Add($tabCreate)
$appSearchLabel = New-Object System.Windows.Forms.Label
$appSearchLabel.Text = "Appsearch:"
$appSearchLabel.Location = New-Object System.Drawing.Point(10,20)
$appSearchLabel.AutoSize = $true
$tabCreate.Controls.Add($appSearchLabel)
$appSearchBox = New-Object System.Windows.Forms.TextBox
# ENTER im App-Suchfeld -> "Search" (auslösen des Such-Buttons)
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
$dropdown.Width = 650
$tabCreate.Controls.Add($dropdown)
$pathLabel = New-Object System.Windows.Forms.Label
$pathLabel.Text = "Filepath:"
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
    if ($folderBrowser.ShowDialog() -eq "OK") {
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
# Label über "Updates suchen"
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
$tabUpdate.Controls.Add($updateSearchButton)
$updateDropdown = New-Object System.Windows.Forms.ComboBox
$updateDropdown.Location = New-Object System.Drawing.Point(100,90)
$updateDropdown.Width = 650
$tabUpdate.Controls.Add($updateDropdown)
$updateSelectedButton = New-Object System.Windows.Forms.Button
$updateSelectedButton.Text = "Update selected apps"
$updateSelectedButton.Location = New-Object System.Drawing.Point(100,130)
$updateSelectedButton.Width = 250
$tabUpdate.Controls.Add($updateSelectedButton)
$updateAllButton = New-Object System.Windows.Forms.Button
$updateAllButton.Text = "Update all apps"
$updateAllButton.Location = New-Object System.Drawing.Point(420,130)
$updateAllButton.Width = 250
$tabUpdate.Controls.Add($updateAllButton)
# Button: Abgelöste Apps suchen
#$supersededLabel = New-Object System.Windows.Forms.Label
#$supersededLabel.Text = "Hier kann man abgelöste Apps entfernen"
#$supersededLabel.AutoSize = $true
#$supersededLabel.Location = New-Object System.Drawing.Point(100, 200)
#$tabUpdate.Controls.Add($supersededLabel)
# Label über "Search Superseded Apps"
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
$tabUpdate.Controls.Add($supersededSearchButton)
# Dropdown: Abgelöste Apps (Name + Version)
$supersededDropdown = New-Object System.Windows.Forms.ComboBox
$supersededDropdown.Location = New-Object System.Drawing.Point(100,240)
$supersededDropdown.Width = 650
$supersededDropdown.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$tabUpdate.Controls.Add($supersededDropdown)
# Button: Ausgewählte App löschen
$deleteSelectedAppButton = New-Object System.Windows.Forms.Button
$deleteSelectedAppButton.Text = "Ausgewählte App löschen"
$deleteSelectedAppButton.Location = New-Object System.Drawing.Point(100,280)
$deleteSelectedAppButton.Width = 250
$tabUpdate.Controls.Add($deleteSelectedAppButton)
$removeOldAppsButton = New-Object System.Windows.Forms.Button
$removeOldAppsButton.Text = "Delete all Superseded Apps"
$removeOldAppsButton.Location = New-Object System.Drawing.Point(360,280)
$removeOldAppsButton.Width = 250
$tabUpdate.Controls.Add($removeOldAppsButton)
# Hashtable zur Speicherung von AppName → {PackageID, Version}
$global:packageMap = @{}
# Modulprüfung
Update-Status "Checking WinTuner Modul..."
if (Get-Module -ListAvailable -Name WinTuner) {
    Update-Status "Module found, searching for updates..."
    Update-Module -Name WinTuner
} else {
    Update-Status "Module not found, install..."
    Install-Module -Name WinTuner -Force
}
Import-Module WinTuner
Update-Status "Module imported."
# Login button
$loginButton = New-Object System.Windows.Forms.Button
$loginButton.Text = "Login to Tenant"
$loginButton.Location = New-Object System.Drawing.Point(570,20)
$loginButton.Width = 180
$form.Controls.Add($loginButton)
# initialize login button enabled state based on username validation
if (Test-ValidM365UserName -UserName $usernameBox.Text) { $loginButton.Enabled = $true } else { $loginButton.Enabled = $false }
$loginButton.Add_Click({
    if (-not (Test-ValidM365UserName -UserName $usernameBox.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a valid M365 UPN.","Invalid Username")
        return
    }
    try {
        Update-Status "Connecting to tenant..."
        $global:isConnected = $false
        $null = Connect-WtWinTuner -Username $usernameBox.Text -ErrorAction Stop
        if (-not (Test-WtConnected)) {
            throw "Authentication error or failed."
        }
        $global:isConnected = $true
        Update-Status "Login success."
        Set-ConnectedUIState -Connected $true
    } catch {
        Update-Status ("Login cancled/failed: {0}" -f $_.Exception.Message)
        Set-ConnectedUIState -Connected $false
    }
})
$searchButton.Add_Click({
    if ($appSearchBox.Text -eq "") {
        [System.Windows.Forms.MessageBox]::Show("App Search can´t be empty.")
        return
    }
    Update-Status "Searching..."
    $results = Search-WtWinGetPackage -SearchQuery $appSearchBox.Text
    $dropdown.Items.Clear()
    $global:packageMap.Clear()
foreach ($result in @($results)) {
		$displayText = "$($result.Name) — $($result.PackageID)"
		$dropdown.Items.Add($displayText)
        $global:packageMap[$displayText] = @{
            PackageID = $result.PackageID
            Version = $result.Version
        }
    }
    if ($dropdown.Items.Count -gt 0) {
        $dropdown.SelectedIndex = 0
    }
    Update-Status "Search completed."
})
$createButton.Add_Click({
    $appName = $dropdown.SelectedItem
    $packageID = $global:packageMap[$appName].PackageID
    $folder = $pathBox.Text
    $filePath = Join-Path $folder "$packageID.wtpackage"
    if (Test-Path $filePath) {
        Update-Status "Paket existiert bereits."
        return
    }
    Update-Status "Creating package..."
    $progressBar.Value = 0
    $progressBar.Visible = $true
    $progressBar.Value = 30
    Start-Sleep -Milliseconds 500
    New-WtWingetPackage -PackageId $packageID -PackageFolder $folder
    $progressBar.Value = 100
    Update-Status "Success creating package."
    $uploadButton.Visible = $true
})
$uploadButton.Add_Click({
    $appName = $dropdown.SelectedItem
    $packageID = $global:packageMap[$appName].PackageID
    $version = $global:packageMap[$appName].Version
    $folder = $pathBox.Text
    if (-not $version) {
        Update-Status "Version could not be determined."
        return
    }
    $progressBar.Value = 0
    $progressBar.Visible = $true
    Update-Status "Starting Upload..."
    Start-Sleep -Milliseconds 500
    Deploy-WtWin32App -PackageId $packageID -Version $version -RootPackageFolder $folder
    $progressBar.Value = 100
    Update-Status "Upload completed."
    $uploadButton.Visible = $false
    $appSearchBox.Text = ""
    $dropdown.Items.Clear()
})
$updateSearchButton.Add_Click({
    Update-Status "Search for updates..."
    try { $tempApps = @(Get-WtWin32Apps -Update $true -Superseded $false) } catch { Write-Verbose "Get-WtWin32Apps update threw: $($_)"; $tempApps = @() }
$updateDropdown.Items.Clear()
    $global:updateApps = @()
    $updateCandidates = @(($tempApps | Where-Object { $_.LatestVersion -and $_.CurrentVersion -and (Test-IsNewerVersion $_.LatestVersion $_.CurrentVersion) }) | Sort-Object Name)
    foreach ($app in @($updateCandidates)) {
        $updateDropdown.Items.Add($app.Name)
        $global:updateApps += $app
    }
    if ($updateDropdown.Items.Count -gt 0) {
        $updateDropdown.SelectedIndex = 0
    }
    Update-Status "Search updates completed"
})
$updateSelectedButton.Add_Click({
    $selectedAppName = $updateDropdown.SelectedItem
    $app = $global:updateApps | Where-Object { $_.Name -eq $selectedAppName }
    $rootPackageFolder = $pathBox.Text
    if (-not $app) {
        Update-Status "No valid app selected."
        return
    }
    Update-Status "Update for $($app.Name) is running..."
    $progressBar.Value = 0
    $progressBar.Visible = $true
    $progressBar.Value = 30
    Start-Sleep -Milliseconds 500
    New-WtWingetPackage -PackageId $app.PackageId -PackageFolder $rootPackageFolder -Version $app.LatestVersion |
        Deploy-WtWin32App -GraphId $app.GraphId -KeepAssignments
    $progressBar.Value = 100
    Update-Status "Update completed."
})
$updateAllButton.Add_Click({
    $rootPackageFolder = $pathBox.Text
    Update-Status "Starting mass update..."
    $progressBar.Value = 0
    $progressBar.Visible = $true
    $progressBar.Value = 10
    try { $updatedApps = @(Get-WtWin32Apps -Update $true -Superseded $false) } catch { Write-Verbose "Get-WtWin32Apps update threw: $($_)"; $updatedApps = @() }
$updatedApps = @(($updatedApps | Where-Object { $_.LatestVersion -and $_.CurrentVersion -and (Test-IsNewerVersion $_.LatestVersion $_.CurrentVersion) }) | Sort-Object Name)
    foreach ($app in @($updatedApps)) {
        Update-Status "Update: $($app.Name)"
        Start-Sleep -Milliseconds 300
        New-WtWingetPackage -PackageId $app.PackageId -PackageFolder $rootPackageFolder -Version $app.LatestVersion |
            Deploy-WtWin32App -GraphId $app.GraphId -KeepAssignments
    }
    $progressBar.Value = 100
    Update-Status "All Updates Completed"
})
$removeOldAppsButton.Add_Click({
    $oldApps = Get-WtWin32Apps -Superseded $true
    if ($oldApps.Count -eq 0) {
        Update-Status "No Superseded Apps Found"
        return
    }
    $appNames = ($oldApps | Select-Object -ExpandProperty Name) -join "`r`n"
    $result = [System.Windows.Forms.MessageBox]::Show("The following outdated apps will be removed:`r`n$appNames", "Bestätigung", "YesNo")
    if ($result -eq "Yes") {
        $progressBar.Value = 0
        $progressBar.Visible = $true
        foreach ($app in @($oldApps)) {
            Remove-WtWin32App -AppId $app.GraphId
            Update-Status "Entfernt: $($app.Name)"
        }
        $progressBar.Value = 100
        Update-Status "Deleted all superseded Apps..."
    } else {
        Update-Status "Removal aborted."
    }
})
# Handler: Abgelöste Apps suchen
$supersededSearchButton.Add_Click({
    try {
        Update-Status "Search for superseded apps..."
        $global:supersededApps = Get-WtWin32Apps -Superseded $true
        $supersededDropdown.Items.Clear()
        foreach ($app in @($global:supersededApps)) {
            $name     = $app.Name
            $version  = $app.CurrentVersion
            $display  = "$name — $version"   # Bindestrich statt nur Leerzeichen
            [void]$supersededDropdown.Items.Add($display)
        }
        if ($supersededDropdown.Items.Count -gt 0) { 
            $supersededDropdown.SelectedIndex = 0 
        }
        Update-Status ("Search completed: {0} superseded Apps found." -f $supersededDropdown.Items.Count)
    } catch {
        Update-Status ("Error while search: {0}" -f $_.Exception.Message)
    }
})
# Handler: Ausgewählte App löschen
$deleteSelectedAppButton.Add_Click({
    if (-not $global:supersededApps -or $supersededDropdown.SelectedIndex -lt 0) {
        Update-Status "Please first select a superseded app from the dropdown."
        return
    }
    $app = $global:supersededApps[$supersededDropdown.SelectedIndex]
    $result = [System.Windows.Forms.MessageBox]::Show("Delete App '" + $app.Name + "'?", "Bestätigung", "YesNo")
    if ($result -eq "Yes") {
        try {
            Remove-WtWin32App -AppId $app.GraphId
            Update-Status ("Deleted: {0}" -f $app.Name)
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
    Update-Status "Logout success."
    Set-ConnectedUIState -Connected $false
})
$form.Add_FormClosing({
    Disconnect-WtWinTuner
    Write-Log "Logout while closing."
})
# Apply initial theme (Dark by default)
Apply-Theme -control $form -theme $global:currentTheme
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::Run($form)

# Safe logger for closing context
function Write-FileLog {
    param([string]$message)
    try {
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        Add-Content -Path 'WinTuner_GUI.log' -Value "$timestamp - $message"
    } catch { }
}

# Re-Entryschutz
$global:_closingInProgress = $false

$form.Add_FormClosing({
    param($sender, [System.Windows.Forms.FormClosingEventArgs]$e)

    if ($global:_closingInProgress) { return }
    if (-not $global:isConnected)  { return }

    $e.Cancel = $true
    $global:_closingInProgress = $true

    try {
        $form.Enabled = $false
        if ($statusLabel) { $statusLabel.Text = "Closing... signing out from tenant" }
    } catch { }

    Write-FileLog 'Shutdown: starting tenant disconnect (BackgroundWorker).'

    # 5s Timeout on UI thread
    $script:_wtCloseTimer = New-Object System.Windows.Forms.Timer
    $script:_wtCloseTimer.Interval = 5000
    $script:_wtCloseTimer.Add_Tick({
        $script:_wtCloseTimer.Stop()
        try { Write-FileLog 'Shutdown: disconnect timeout after 5s. Forcing close.' } catch { }
        try { $form.Close() } catch { }
    })
    $script:_wtCloseTimer.Start()

    # Disconnect in background
    $bw = New-Object System.ComponentModel.BackgroundWorker
    $bw.WorkerSupportsCancellation = $false
    $bw.Add_DoWork({ try { Disconnect-WtWinTuner } catch { } })
    $bw.Add_RunWorkerCompleted({
        try { Write-FileLog 'Shutdown: disconnect finished. Closing form.' } catch { }
        try { if ($script:_wtCloseTimer) { $script:_wtCloseTimer.Stop() } } catch { }
        try { $form.Close() } catch { }
    })
    $bw.RunWorkerAsync()
})

'@

Invoke-Expression $__WinTunerMain
