# Updated WinTuner_GUI.ps1

# Previous Code Block...
$deleteSelectedAppButton.Text = "Delete selected app"

# Replace message strings
MessageBox::Show("Confirmation")

# Update status messages
Update-Status ("Removed: {0}" -f $app.Name)

# Update All button handler
$appName = $app.Name; $appCurrentVersion = $app.CurrentVersion; $appLatestVersion = $app.LatestVersion; $appGraphId = $app.GraphId; $appPackageId = Try-ResolveWingetIdForApp -App $app; $result = Update-SingleApp -AppName $appName -CurrentVersion $appCurrentVersion -LatestVersion $appLatestVersion -GraphId $appGraphId -PackageIdentifier $appPackageId -RootPackageFolder $rootPackageFolder

# Rest of the file...