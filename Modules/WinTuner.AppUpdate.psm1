Set-StrictMode -Version Latest

function Get-WinTunerAppUpdateProperty {
    param(
        [Parameter(Mandatory)][object]$InputObject,
        [Parameter(Mandatory)][string]$Name
    )

    $property = $InputObject.PSObject.Properties[$Name]
    if ($property) { return [string]$property.Value }
    return ''
}

function Get-WinTunerAppUpdateObjectValue {
    param(
        [AllowNull()][object]$InputObject,
        [Parameter(Mandatory)][string]$Name
    )

    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [System.Collections.IDictionary]) {
        if ($InputObject.Contains($Name)) { return $InputObject[$Name] }
        return $null
    }

    $property = $InputObject.PSObject.Properties[$Name]
    if ($property) { return $property.Value }
    return $null
}

function ConvertTo-WinTunerAppUpdateHashtable {
    param([AllowNull()][object]$InputObject)

    if ($null -eq $InputObject) { return [ordered]@{} }
    return ($InputObject | ConvertTo-Json -Depth 30 | ConvertFrom-Json -AsHashtable -Depth 30)
}

function Set-WinTunerAvailableAssignmentAutoUpdate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$GraphId,
        [scriptblock]$InvokeGraphRequest = {
            param([string]$Method, [string]$Uri, [AllowNull()][string]$Body)
            $parameters = @{ Method = $Method; Uri = $Uri; ErrorAction = 'Stop' }
            if (-not [string]::IsNullOrWhiteSpace($Body)) {
                $parameters.Body = $Body
                $parameters.ContentType = 'application/json'
            }
            Invoke-WinTunerGraphRequest @parameters
        }
    )

    $escapedGraphId = [uri]::EscapeDataString($GraphId)
    $nextLink = "https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps/$escapedGraphId/assignments?`$top=999"
    $assignments = [System.Collections.Generic.List[object]]::new()
    while (-not [string]::IsNullOrWhiteSpace($nextLink)) {
        $page = & $InvokeGraphRequest 'GET' $nextLink $null
        foreach ($assignment in @(Get-WinTunerAppUpdateObjectValue -InputObject $page -Name 'value')) {
            if ($assignment) { $assignments.Add($assignment) }
        }
        $nextLink = [string](Get-WinTunerAppUpdateObjectValue -InputObject $page -Name '@odata.nextLink')
    }

    $availableCount = 0
    $enabledCount = 0
    $alreadyEnabledCount = 0
    foreach ($assignment in $assignments) {
        if (-not [string]::Equals([string](Get-WinTunerAppUpdateObjectValue -InputObject $assignment -Name 'intent'), 'available', [System.StringComparison]::OrdinalIgnoreCase)) {
            continue
        }
        $availableCount++

        $settings = ConvertTo-WinTunerAppUpdateHashtable -InputObject (Get-WinTunerAppUpdateObjectValue -InputObject $assignment -Name 'settings')
        $existingAutoUpdate = Get-WinTunerAppUpdateObjectValue -InputObject $settings -Name 'autoUpdateSettings'
        $existingState = [string](Get-WinTunerAppUpdateObjectValue -InputObject $existingAutoUpdate -Name 'autoUpdateSupersededAppsState')
        if ([string]::Equals($existingState, 'enabled', [System.StringComparison]::OrdinalIgnoreCase)) {
            $alreadyEnabledCount++
            continue
        }

        if (-not $settings.Contains('@odata.type')) {
            $settings['@odata.type'] = '#microsoft.graph.win32LobAppAssignmentSettings'
        }
        $settings['autoUpdateSettings'] = [ordered]@{
            '@odata.type' = '#microsoft.graph.win32LobAppAutoUpdateSettings'
            autoUpdateSupersededAppsState = 'enabled'
        }

        $assignmentId = [string](Get-WinTunerAppUpdateObjectValue -InputObject $assignment -Name 'id')
        if ([string]::IsNullOrWhiteSpace($assignmentId)) {
            throw 'Microsoft Graph returned an Available assignment without an ID.'
        }
        $assignmentUri = "https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps/$escapedGraphId/assignments/$([uri]::EscapeDataString($assignmentId))"
        $body = @{ settings = $settings } | ConvertTo-Json -Depth 30 -Compress
        $null = & $InvokeGraphRequest 'PATCH' $assignmentUri $body
        $enabledCount++
    }

    $summary = if ($availableCount -eq 0) {
        'No Available assignments found; AutoUpdate was not required.'
    }
    elseif ($enabledCount -eq 0) {
        "AutoUpdate was already enabled for all $alreadyEnabledCount Available assignment(s)."
    }
    else {
        "Enabled AutoUpdate for $enabledCount Available assignment(s); already enabled for $alreadyEnabledCount."
    }
    return [pscustomobject]@{
        AvailableCount      = $availableCount
        EnabledCount        = $enabledCount
        AlreadyEnabledCount = $alreadyEnabledCount
        Summary             = $summary
    }
}

function ConvertTo-WinTunerCanonicalPlanValue {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [Collections.IDictionary] -or $Value -is [pscustomobject]) {
        $result = [ordered]@{}
        $keys = if ($Value -is [Collections.IDictionary]) { @($Value.Keys) } else { @($Value.PSObject.Properties.Name) }
        foreach ($key in @($keys | Sort-Object)) {
            $result[$key] = ConvertTo-WinTunerCanonicalPlanValue (Get-WinTunerAppUpdateObjectValue $Value $key)
        }
        return $result
    }
    if ($Value -is [Collections.IEnumerable] -and $Value -isnot [string]) {
        $items = [Collections.Generic.List[object]]::new()
        foreach ($item in $Value) { $items.Add((ConvertTo-WinTunerCanonicalPlanValue $item)) }
        return ,$items.ToArray()
    }
    return $Value
}

function New-WinTunerUpdatePlan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object[]]$Apps,
        [Parameter(Mandatory)][guid]$TenantId,
        [Parameter(Mandatory)][string]$RootPackageFolder,
        [scriptblock]$GetApp = {
            param($GraphId)
            # displayVersion is exposed by the Intune beta Win32 model, not v1.0.
            Invoke-WinTunerGraphRequest -Uri "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$([uri]::EscapeDataString($GraphId))"
        },
        [scriptblock]$GetAssignments = {
            param($GraphId)
            $uri = "https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps/$([uri]::EscapeDataString($GraphId))/assignments"
            $seen = [Collections.Generic.HashSet[string]]::new()
            while (-not [string]::IsNullOrWhiteSpace($uri)) {
                if (-not $seen.Add($uri) -or $seen.Count -gt 1000) { throw 'Invalid assignment pagination.' }
                $target = [uri]$uri
                if ($target.Scheme -ne 'https' -or $target.Host -ne 'graph.microsoft.com' -or $target.Port -ne 443) { throw 'Invalid Graph pagination target.' }
                $page = Invoke-WinTunerGraphRequest -Uri $uri
                foreach ($entry in @(Get-WinTunerAppUpdateObjectValue $page 'value')) { if ($entry) { $entry } }
                $uri = [string](Get-WinTunerAppUpdateObjectValue $page '@odata.nextLink')
            }
        }
    )
    if ($TenantId -eq [guid]::Empty) { throw 'A verified tenant ID is required for deployment preview.' }
    $items = [Collections.Generic.List[object]]::new()
    foreach ($app in $Apps) {
        foreach ($field in 'Name','GraphId','PackageId','CurrentVersion','LatestVersion') {
            if ([string]::IsNullOrWhiteSpace((Get-WinTunerAppUpdateProperty $app $field))) { throw "Deployment preview requires $field for every app. Run a fresh update scan." }
        }
        $current = & $GetApp ([string]$app.GraphId)
        if ([string](Get-WinTunerAppUpdateObjectValue $current 'id') -ne [string]$app.GraphId) {
            throw "The Intune app identity for '$($app.Name)' could not be verified. Run a fresh update scan."
        }
        $liveVersion = [string](Get-WinTunerAppUpdateObjectValue $current 'displayVersion')
        if ([string]::IsNullOrWhiteSpace($liveVersion)) {
            throw "Cannot read the current Intune version of '$($app.Name)': displayVersion is missing. No upload was started."
        }
        if ($liveVersion -cne [string]$app.CurrentVersion) {
            throw "The current Intune version of '$($app.Name)' changed. Run a fresh update scan."
        }
        $assignments = @(& $GetAssignments ([string]$app.GraphId))
        $canonical = ConvertTo-WinTunerCanonicalPlanValue @($assignments | Sort-Object { [string](Get-WinTunerAppUpdateObjectValue $_ 'id') })
        $json = ConvertTo-Json -InputObject $canonical -Depth 40 -Compress
        $hash = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes($json)))
        $items.Add([pscustomobject]@{
            Name = [string]$app.Name; GraphId = [string]$app.GraphId; PackageId = [string]$app.PackageId
            CurrentVersion = [string]$app.CurrentVersion; LatestVersion = [string]$app.LatestVersion
            AssignmentHash = $hash; Assignments = @($assignments)
            RuleDescription = [string](Get-WinTunerAppUpdateObjectValue $app 'RuleDescription')
        })
    }
    [pscustomobject]@{ TenantId = $TenantId.ToString('D'); RootPackageFolder = $RootPackageFolder; Items = @($items); CreatedAtUtc = [datetime]::UtcNow.ToString('o') }
}

function Format-WinTunerUpdatePlan {
    param([Parameter(Mandatory)][object]$Plan)
    $lines = [Collections.Generic.List[string]]::new()
    $lines.Add("Tenant: $($Plan.TenantId)")
    $lines.Add("Package folder: $($Plan.RootPackageFolder)")
    foreach ($item in @($Plan.Items)) {
        $lines.Add('')
        $lines.Add("$($item.Name): $($item.CurrentVersion) -> $($item.LatestVersion) [$($item.PackageId)]")
        $ruleDescription = [string](Get-WinTunerAppUpdateObjectValue $item 'RuleDescription')
        if ($ruleDescription) { $lines.Add("  Rule: $ruleDescription") }
        if (@($item.Assignments).Count -eq 0) { $lines.Add('  No assignments. The new app will not be assigned.'); continue }
        foreach ($assignment in @($item.Assignments)) {
            $intent = [string](Get-WinTunerAppUpdateObjectValue $assignment 'intent')
            $target = Get-WinTunerAppUpdateObjectValue $assignment 'target'
            $targetType = [string](Get-WinTunerAppUpdateObjectValue $target '@odata.type')
            $groupId = [string](Get-WinTunerAppUpdateObjectValue $target 'groupId')
            $targetText = if ($targetType -match 'exclusionGroup') { "Excluded group: $groupId" } elseif ($groupId) { "Group: $groupId" } elseif ($targetType -match 'allLicensedUsers') { 'All users' } elseif ($targetType -match 'allDevices') { 'All devices' } else { $targetType }
            $filterId = [string](Get-WinTunerAppUpdateObjectValue $target 'deviceAndAppManagementAssignmentFilterId')
            if ($filterId) { $targetText += " | Filter: $filterId ($([string](Get-WinTunerAppUpdateObjectValue $target 'deviceAndAppManagementAssignmentFilterType')))" }
            $auto = if ($intent -eq 'available') { ' | AutoUpdate: enabled on new app' } else { '' }
            $lines.Add("  $intent | $targetText$auto")
        }
    }
    $lines.Add('')
    $skippedProperty = $Plan.PSObject.Properties['Skipped']
    if ($skippedProperty -and @($skippedProperty.Value).Count -gt 0) {
        $lines.Add('Skipped by update rules:')
        foreach ($entry in @($skippedProperty.Value)) { $lines.Add("  $($entry.Name): $($entry.Description)") }
    }
    $lines.Add('Assignments will be preserved. If the Intune version, assignments or update rules change before upload, the affected update is blocked for a new review.')
    $lines -join [Environment]::NewLine
}

function New-WinTunerAppUpdateItemResult {
    param(
        [bool]$Succeeded,
        [string]$Name,
        [AllowNull()][string]$CurrentVersion,
        [AllowNull()][string]$TargetVersion,
        [AllowNull()][string]$EffectiveVersion,
        [AllowNull()][string]$GraphId,
        [AllowNull()][string]$PackageId,
        [string]$ReasonCode,
        [string]$Message,
        [AllowNull()][string]$AutoUpdateSummary
    )

    [pscustomobject]@{
        Succeeded        = $Succeeded
        Name             = $Name
        CurrentVersion   = $CurrentVersion
        TargetVersion    = $TargetVersion
        EffectiveVersion = $EffectiveVersion
        GraphId          = $GraphId
        PackageId        = $PackageId
        ReasonCode       = $ReasonCode
        Message          = $Message
        AutoUpdateSummary = $AutoUpdateSummary
    }
}

function Invoke-WinTunerAppUpdateItem {
    param(
        [Parameter(Mandatory)][object]$App,
        [Parameter(Mandatory)][string]$RootPackageFolder,
        [Parameter(Mandatory)][scriptblock]$BuildPackage,
        [Parameter(Mandatory)][scriptblock]$ValidateArtifact,
        [Parameter(Mandatory)][scriptblock]$DeployPackage,
        [Parameter(Mandatory)][scriptblock]$ConfigureAvailableAssignmentAutoUpdate,
        [Parameter(Mandatory)][scriptblock]$ValidatePlan,
        [Parameter(Mandatory)][scriptblock]$ReportRetry
    )

    $name = Get-WinTunerAppUpdateProperty -InputObject $App -Name 'Name'
    $currentVersion = Get-WinTunerAppUpdateProperty -InputObject $App -Name 'CurrentVersion'
    $targetVersion = Get-WinTunerAppUpdateProperty -InputObject $App -Name 'LatestVersion'
    $graphId = Get-WinTunerAppUpdateProperty -InputObject $App -Name 'GraphId'
    $packageId = Get-WinTunerAppUpdateProperty -InputObject $App -Name 'PackageId'

    if ([string]::IsNullOrWhiteSpace($packageId)) {
        return New-WinTunerAppUpdateItemResult `
            -Succeeded $false `
            -Name $name `
            -CurrentVersion $currentVersion `
            -TargetVersion $targetVersion `
            -GraphId $graphId `
            -ReasonCode 'MissingPackageId' `
            -Message "Cannot determine PackageId for '$name'."
    }

    if ([string]::IsNullOrWhiteSpace($targetVersion)) {
        return New-WinTunerAppUpdateItemResult `
            -Succeeded $false `
            -Name $name `
            -CurrentVersion $currentVersion `
            -GraphId $graphId `
            -PackageId $packageId `
            -ReasonCode 'MissingTargetVersion' `
            -Message "Cannot determine the target version for '$name'."
    }

    try {
        $buildResult = & $BuildPackage $packageId $RootPackageFolder $targetVersion $currentVersion $ReportRetry
    }
    catch {
        return New-WinTunerAppUpdateItemResult `
            -Succeeded $false `
            -Name $name `
            -CurrentVersion $currentVersion `
            -TargetVersion $targetVersion `
            -GraphId $graphId `
            -PackageId $packageId `
            -ReasonCode 'PackageBuildFailed' `
            -Message "Package creation failed for ${name}: $($_.Exception.Message)"
    }

    if (-not $buildResult -or -not [bool]$buildResult.Succeeded) {
        $detail = if ($buildResult -and -not [string]::IsNullOrWhiteSpace([string]$buildResult.ErrorMessage)) {
            [string]$buildResult.ErrorMessage
        }
        else {
            'The package builder returned no successful result.'
        }

        return New-WinTunerAppUpdateItemResult `
            -Succeeded $false `
            -Name $name `
            -CurrentVersion $currentVersion `
            -TargetVersion $targetVersion `
            -GraphId $graphId `
            -PackageId $packageId `
            -ReasonCode 'PackageBuildFailed' `
            -Message "Package creation failed for ${name}: $detail"
    }

    $effectiveVersion = if ([string]::IsNullOrWhiteSpace([string]$buildResult.EffectiveVersion)) {
        $targetVersion
    }
    else {
        [string]$buildResult.EffectiveVersion
    }

    try {
        $validation = & $ValidateArtifact $RootPackageFolder $packageId $effectiveVersion
    }
    catch {
        return New-WinTunerAppUpdateItemResult `
            -Succeeded $false `
            -Name $name `
            -CurrentVersion $currentVersion `
            -TargetVersion $targetVersion `
            -EffectiveVersion $effectiveVersion `
            -GraphId $graphId `
            -PackageId $packageId `
            -ReasonCode 'ValidationError' `
            -Message "Package validation failed for ${name}: $($_.Exception.Message)"
    }

    if (-not $validation -or -not [bool]$validation.IsValid) {
        $validationReason = if ($validation -and -not [string]::IsNullOrWhiteSpace([string]$validation.Reason)) {
            [string]$validation.Reason
        }
        else {
            'The exact package artifact is invalid.'
        }
        $validationCode = if ($validation -and -not [string]::IsNullOrWhiteSpace([string]$validation.ReasonCode)) {
            [string]$validation.ReasonCode
        }
        else {
            'InvalidArtifact'
        }

        return New-WinTunerAppUpdateItemResult `
            -Succeeded $false `
            -Name $name `
            -CurrentVersion $currentVersion `
            -TargetVersion $targetVersion `
            -EffectiveVersion $effectiveVersion `
            -GraphId $graphId `
            -PackageId $packageId `
            -ReasonCode $validationCode `
            -Message "Package validation failed for $name ($validationCode): $validationReason"
    }

    try {
        $null = & $ValidatePlan $App $effectiveVersion
    } catch {
        return New-WinTunerAppUpdateItemResult -Succeeded $false -Name $name -CurrentVersion $currentVersion -TargetVersion $targetVersion -EffectiveVersion $effectiveVersion -GraphId $graphId -PackageId $packageId -ReasonCode 'PlanChanged' -Message "Update blocked for ${name}: $($_.Exception.Message)"
    }

    try {
        $deployment = & $DeployPackage $App $packageId $effectiveVersion $RootPackageFolder
    }
    catch {
        return New-WinTunerAppUpdateItemResult `
            -Succeeded $false `
            -Name $name `
            -CurrentVersion $currentVersion `
            -TargetVersion $targetVersion `
            -EffectiveVersion $effectiveVersion `
            -GraphId $graphId `
            -PackageId $packageId `
            -ReasonCode 'DeploymentFailed' `
            -Message "Update failed for ${name}: $($_.Exception.Message)"
    }

    $autoUpdateSummary = ''
    if (-not [string]::IsNullOrWhiteSpace($graphId)) {
        $deployedGraphId = [string](Get-WinTunerAppUpdateObjectValue -InputObject $deployment -Name 'Id')
        if ([string]::IsNullOrWhiteSpace($deployedGraphId)) {
            $autoUpdateSummary = 'Warning: AutoUpdate could not be configured because the new app ID was not returned.'
        }
        else {
            try {
                $autoUpdateResult = & $ConfigureAvailableAssignmentAutoUpdate $deployedGraphId
                $autoUpdateSummary = [string](Get-WinTunerAppUpdateObjectValue -InputObject $autoUpdateResult -Name 'Summary')
            }
            catch {
                $autoUpdateSummary = "Warning: Available-assignment AutoUpdate could not be configured: $($_.Exception.Message)"
            }
        }
    }

    return New-WinTunerAppUpdateItemResult `
        -Succeeded $true `
        -Name $name `
        -CurrentVersion $currentVersion `
        -TargetVersion $targetVersion `
        -EffectiveVersion $effectiveVersion `
        -GraphId $graphId `
        -PackageId $packageId `
        -ReasonCode 'Updated' `
        -Message "Update completed successfully for $name." `
        -AutoUpdateSummary $autoUpdateSummary
}

function Invoke-WinTunerAppUpdateBatch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateNotNull()][object[]]$Apps,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$RootPackageFolder,
        [scriptblock]$BuildPackage = {
            param($Id, $Folder, $Target, $Installed, $RetryReporter)
            Invoke-WinTunerPackageBuild `
                -PackageId $Id `
                -PackageFolder $Folder `
                -DesiredVersion $Target `
                -LatestVersion $Target `
                -InstalledVersion $Installed `
                -ReportRetry $RetryReporter
        },
        [scriptblock]$ValidateArtifact = {
            param($Root, $Id, $Version)
            Test-WinTunerPackageArtifact `
                -RootPackageFolder $Root `
                -PackageId $Id `
                -Version $Version
        },
        [scriptblock]$DeployPackage = {
            param($App, $Id, $Version, $Root)
            if (-not [string]::IsNullOrWhiteSpace([string]$App.GraphId)) {
                Deploy-WtWin32App `
                    -GraphId ([string]$App.GraphId) `
                    -KeepAssignments `
                    -PackageId $Id `
                    -Version $Version `
                    -RootPackageFolder $Root `
                    -ErrorAction Stop
            }
            else {
                Deploy-WtWin32App `
                    -PackageId $Id `
                    -Version $Version `
                    -RootPackageFolder $Root `
                    -ErrorAction Stop
            }
        },
        [scriptblock]$ConfigureAvailableAssignmentAutoUpdate = {
            param($DeployedGraphId)
            Set-WinTunerAvailableAssignmentAutoUpdate -GraphId $DeployedGraphId
        },
        [scriptblock]$ValidatePlan = { param($App, $Version) },
        [scriptblock]$RecordResult = { param($Result) },
        [scriptblock]$ReportProgress = { param($ProgressInfo) }
    )

    $items = @($Apps)
    $results = [System.Collections.Generic.List[object]]::new()
    $processed = 0

    foreach ($app in $items) {
        $processed++
        try {
            & $ReportProgress ([pscustomobject]@{
                Stage     = 'Updating'
                Processed = $processed
                Total     = $items.Count
                AppName   = Get-WinTunerAppUpdateProperty -InputObject $app -Name 'Name'
            })
        }
        catch {}

        $itemResult = Invoke-WinTunerAppUpdateItem `
            -App $app `
            -RootPackageFolder $RootPackageFolder `
            -BuildPackage $BuildPackage `
            -ValidateArtifact $ValidateArtifact `
            -DeployPackage $DeployPackage `
            -ConfigureAvailableAssignmentAutoUpdate $ConfigureAvailableAssignmentAutoUpdate `
            -ValidatePlan $ValidatePlan `
            -ReportRetry {
                param($retryInfo)
                $null = & $ReportProgress ([pscustomobject]@{
                    Stage = 'Retry'; Processed = $processed; Total = $items.Count
                    AppName = [string]$app.Name
                    DelaySeconds = $retryInfo.DelaySeconds; Attempt = $retryInfo.Attempt; MaxAttempts = $retryInfo.MaxAttempts
                })
            }.GetNewClosure()
        try { $null = & $RecordResult $itemResult }
        catch { $itemResult | Add-Member -NotePropertyName HistoryWarning -NotePropertyValue 'The local history entry could not be saved; the deployment result is unchanged.' }
        $results.Add($itemResult)
    }

    $successCount = @($results | Where-Object Succeeded).Count
    [pscustomobject]@{
        SuccessCount = $successCount
        FailureCount = $results.Count - $successCount
        Results      = @($results)
    }
}

Export-ModuleMember -Function Invoke-WinTunerAppUpdateBatch,New-WinTunerUpdatePlan,Format-WinTunerUpdatePlan
