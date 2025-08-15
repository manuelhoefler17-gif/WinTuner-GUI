#WinTuner GUI by Manuel Höfler

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

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
    foreach ($childControl in $control.Controls) {
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
        $usernameError.Text = "Bitte gültigen M365 UPN eingeben, z.B. name@firma.de"
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
$tabCreate.Text = "Paket erstellen"
$tabControl.TabPages.Add($tabCreate)

$appSearchLabel = New-Object System.Windows.Forms.Label
$appSearchLabel.Text = "Appsuche:"
$appSearchLabel.Location = New-Object System.Drawing.Point(10,20)
$appSearchLabel.AutoSize = $true
$tabCreate.Controls.Add($appSearchLabel)

$appSearchBox = New-Object System.Windows.Forms.TextBox
$appSearchBox.Location = New-Object System.Drawing.Point(100,20)
$appSearchBox.Width = 450
$appSearchBox.BorderStyle = 'FixedSingle'
$tabCreate.Controls.Add($appSearchBox)

$searchButton = New-Object System.Windows.Forms.Button
$searchButton.Text = "Suchen"
$searchButton.Location = New-Object System.Drawing.Point(570,20)
$searchButton.Width = 180
$tabCreate.Controls.Add($searchButton)

$dropdown = New-Object System.Windows.Forms.ComboBox
$dropdown.Location = New-Object System.Drawing.Point(100,60)
$dropdown.Width = 650
$tabCreate.Controls.Add($dropdown)

$pathLabel = New-Object System.Windows.Forms.Label
$pathLabel.Text = "Speicherpfad:"
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
$browseButton.Text = "Durchsuchen..."
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
$createButton.Text = "Paket erstellen"
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
$updateHeaderLabel.Text = "Vohanden Apps aktualisieren"
$updateHeaderLabel.Location = New-Object System.Drawing.Point(100,20)
$updateHeaderLabel.AutoSize = $true
$updateHeaderLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)

$tabUpdate.Controls.Add($updateHeaderLabel)
$updateSearchButton = New-Object System.Windows.Forms.Button
$updateSearchButton.Text = "Updates suchen"
$updateSearchButton.Location = New-Object System.Drawing.Point(100,50)
$updateSearchButton.Width = 180
$tabUpdate.Controls.Add($updateSearchButton)

$updateDropdown = New-Object System.Windows.Forms.ComboBox
$updateDropdown.Location = New-Object System.Drawing.Point(100,90)
$updateDropdown.Width = 650
$tabUpdate.Controls.Add($updateDropdown)

$updateSelectedButton = New-Object System.Windows.Forms.Button
$updateSelectedButton.Text = "Ausgewählte App updaten"
$updateSelectedButton.Location = New-Object System.Drawing.Point(100,130)
$updateSelectedButton.Width = 250
$tabUpdate.Controls.Add($updateSelectedButton)

$updateAllButton = New-Object System.Windows.Forms.Button
$updateAllButton.Text = "Alle Apps updaten"
$updateAllButton.Location = New-Object System.Drawing.Point(420,130)
$updateAllButton.Width = 250
$tabUpdate.Controls.Add($updateAllButton)
# Button: Abgelöste Apps suchen
#$supersededLabel = New-Object System.Windows.Forms.Label
#$supersededLabel.Text = "Hier kann man abgelöste Apps entfernen"
#$supersededLabel.AutoSize = $true
#$supersededLabel.Location = New-Object System.Drawing.Point(100, 200)
#$tabUpdate.Controls.Add($supersededLabel)


# Label über "Abgelöste Apps suchen"
$supersededHeaderLabel = New-Object System.Windows.Forms.Label
$supersededHeaderLabel.Text = "Hier kann man abgelöste Apps entfernen"
$supersededHeaderLabel.Location = New-Object System.Drawing.Point(100,170)
$supersededHeaderLabel.AutoSize = $true
$supersededHeaderLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$tabUpdate.Controls.Add($supersededHeaderLabel)
$supersededSearchButton = New-Object System.Windows.Forms.Button
$supersededSearchButton.Text = "Abgelöste Apps suchen"
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
$removeOldAppsButton.Text = "Abgelöste Apps entfernen"
$removeOldAppsButton.Location = New-Object System.Drawing.Point(360,280)
$removeOldAppsButton.Width = 250
$tabUpdate.Controls.Add($removeOldAppsButton)

# Hashtable zur Speicherung von AppName → {PackageID, Version}
$global:packageMap = @{}

# Modulprüfung
Update-Status "Prüfe WinTuner Modul..."
if (Get-Module -ListAvailable -Name WinTuner) {
    Update-Status "Modul gefunden, prüfe auf Updates..."
    Update-Module -Name WinTuner
} else {
    Update-Status "Modul nicht gefunden, installiere..."
    Install-Module -Name WinTuner -Force
}
Import-Module WinTuner
Update-Status "Modul importiert."

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
        [System.Windows.Forms.MessageBox]::Show("Bitte geben Sie einen gültigen M365-Benutzernamen (UPN) ein.","Ungültiger Benutzername")
        return
    }
    try {
        Update-Status "Verbinde mit Tenant..."
        $global:isConnected = $false
        $null = Connect-WtWinTuner -Username $usernameBox.Text -ErrorAction Stop
        if (-not (Test-WtConnected)) {
            throw "Authentifizierung abgebrochen oder fehlgeschlagen."
        }
        $global:isConnected = $true
        Update-Status "Login erfolgreich."
        Set-ConnectedUIState -Connected $true
    } catch {
        Update-Status ("Login abgebrochen/fehlgeschlagen: {0}" -f $_.Exception.Message)
        Set-ConnectedUIState -Connected $false
    }
})
$searchButton.Add_Click({
    if ($appSearchBox.Text -eq "") {
        [System.Windows.Forms.MessageBox]::Show("Appsuche ist ein Pflichtfeld.")
        return
    }
    Update-Status "Suche läuft..."
    $results = Search-WtWinGetPackage -SearchQuery $appSearchBox.Text
    $dropdown.Items.Clear()
    $global:packageMap.Clear()

foreach ($result in $results) {
        
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
    Update-Status "Suche abgeschlossen."
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

    Update-Status "Paket wird erstellt..."
    $progressBar.Value = 0
    $progressBar.Visible = $true
    $progressBar.Value = 30
    Start-Sleep -Milliseconds 500
    New-WtWingetPackage -PackageId $packageID -PackageFolder $folder
    $progressBar.Value = 100
    Update-Status "Paket erfolgreich erstellt."
    $uploadButton.Visible = $true
})

$uploadButton.Add_Click({
    $appName = $dropdown.SelectedItem
    $packageID = $global:packageMap[$appName].PackageID
    $version = $global:packageMap[$appName].Version
    $folder = $pathBox.Text

    if (-not $version) {
        Update-Status "Version konnte nicht ermittelt werden."
        return
    }

    $progressBar.Value = 0
    $progressBar.Visible = $true
    Update-Status "Starte Upload..."
    Start-Sleep -Milliseconds 500
    Deploy-WtWin32App -PackageId $packageID -Version $version -RootPackageFolder $folder
    $progressBar.Value = 100
    Update-Status "Upload abgeschlossen."

    $uploadButton.Visible = $false
    $appSearchBox.Text = ""
    $dropdown.Items.Clear()
})

$updateSearchButton.Add_Click({
    Update-Status "Suche nach Updates..."
    $tempApps = Get-WtWin32Apps -Update $true -Superseded $false
    $updateDropdown.Items.Clear()
    $global:updateApps = @()

    foreach ($app in $tempApps) {
        $updateDropdown.Items.Add($app.Name)
        $global:updateApps += $app
    }

    if ($updateDropdown.Items.Count -gt 0) {
        $updateDropdown.SelectedIndex = 0
    }

    Update-Status "Updatesuche abgeschlossen."
})

$updateSelectedButton.Add_Click({
    $selectedAppName = $updateDropdown.SelectedItem
    $app = $global:updateApps | Where-Object { $_.Name -eq $selectedAppName }
    $rootPackageFolder = $pathBox.Text

    if (-not $app) {
        Update-Status "Keine gültige App ausgewählt."
        return
    }

    Update-Status "Update für $($app.Name) wird durchgeführt..."
    $progressBar.Value = 0
    $progressBar.Visible = $true
    $progressBar.Value = 30
    Start-Sleep -Milliseconds 500
    New-WtWingetPackage -PackageId $app.PackageId -PackageFolder $rootPackageFolder -Version $app.LatestVersion |
        Deploy-WtWin32App -GraphId $app.GraphId -KeepAssignments
    $progressBar.Value = 100
    Update-Status "Update abgeschlossen."
})

$updateAllButton.Add_Click({
    $rootPackageFolder = $pathBox.Text
    Update-Status "Starte Massen-Update..."
    $progressBar.Value = 0
    $progressBar.Visible = $true
    $progressBar.Value = 10
    $updatedApps = Get-WtWin32Apps -Update $true -Superseded $false

    foreach ($app in $updatedApps) {
        Update-Status "Update: $($app.Name)"
        Start-Sleep -Milliseconds 300
        New-WtWingetPackage -PackageId $app.PackageId -PackageFolder $rootPackageFolder -Version $app.LatestVersion |
            Deploy-WtWin32App -GraphId $app.GraphId -KeepAssignments
    }

    $progressBar.Value = 100
    Update-Status "Alle Updates abgeschlossen."
})

$removeOldAppsButton.Add_Click({
    $oldApps = Get-WtWin32Apps -Superseded $true
    if ($oldApps.Count -eq 0) {
        Update-Status "Keine abgelösten Apps gefunden."
        return
    }

    $appNames = ($oldApps | Select-Object -ExpandProperty Name) -join "`r`n"
    $result = [System.Windows.Forms.MessageBox]::Show("Folgende veraltete Apps werden entfernt:`r`n$appNames", "Bestätigung", "YesNo")

    if ($result -eq "Yes") {
        $progressBar.Value = 0
        $progressBar.Visible = $true
        foreach ($app in $oldApps) {
            Remove-WtWin32App -AppId $app.GraphId
            Update-Status "Entfernt: $($app.Name)"
        }
        $progressBar.Value = 100
        Update-Status "Alle abgelösten Apps wurden entfernt."
    } else {
        Update-Status "Entfernen abgebrochen."
    }
})



# Handler: Abgelöste Apps suchen
$supersededSearchButton.Add_Click({
    try {
        Update-Status "Suche nach abgelösten Apps..."
        $global:supersededApps = Get-WtWin32Apps -Superseded $true
        $supersededDropdown.Items.Clear()

        foreach ($app in $global:supersededApps) {
            $name     = $app.Name
            $version  = $app.CurrentVersion
            $display  = "$name — $version"   # Bindestrich statt nur Leerzeichen

            [void]$supersededDropdown.Items.Add($display)
        }

        if ($supersededDropdown.Items.Count -gt 0) { 
            $supersededDropdown.SelectedIndex = 0 
        }
        Update-Status ("Suche abgeschlossen: {0} abgelöste Apps gefunden." -f $supersededDropdown.Items.Count)
    } catch {
        Update-Status ("Fehler bei der Suche: {0}" -f $_.Exception.Message)
    }
})

# Handler: Ausgewählte App löschen
$deleteSelectedAppButton.Add_Click({
    if (-not $global:supersededApps -or $supersededDropdown.SelectedIndex -lt 0) {
        Update-Status "Bitte zuerst eine abgelöste App im Dropdown auswählen."
        return
    }
    $app = $global:supersededApps[$supersededDropdown.SelectedIndex]
    $result = [System.Windows.Forms.MessageBox]::Show("Soll die App '" + $app.Name + "' gelöscht werden?", "Bestätigung", "YesNo")
    if ($result -eq "Yes") {
        try {
            Remove-WtWin32App -AppId $app.GraphId
            Update-Status ("Gelöscht: {0}" -f $app.Name)
        } catch {
            Update-Status ("Fehler beim Löschen: {0}" -f $_.Exception.Message)
        }
    } else {
        Update-Status "Löschen abgebrochen."
    }
})

$logoutButton.Add_Click({
    Disconnect-WtWinTuner
    $global:isConnected = $false
    Update-Status "Logout erfolgreich."
    Set-ConnectedUIState -Connected $false
})
$form.Add_FormClosing({
    Disconnect-WtWinTuner
    Write-Log "Logout beim Schließen durchgeführt."
})

# Apply initial theme (Dark by default)
Apply-Theme -control $form -theme $global:currentTheme

[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::Run($form)
