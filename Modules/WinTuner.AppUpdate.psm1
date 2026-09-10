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
        [string]$Message
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
    }
}

function Invoke-WinTunerAppUpdateItem {
    param(
        [Parameter(Mandatory)][object]$App,
        [Parameter(Mandatory)][string]$RootPackageFolder,
        [Parameter(Mandatory)][scriptblock]$BuildPackage,
        [Parameter(Mandatory)][scriptblock]$ValidateArtifact,
        [Parameter(Mandatory)][scriptblock]$DeployPackage
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
        $buildResult = & $BuildPackage $packageId $RootPackageFolder $targetVersion $currentVersion
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
        $null = & $DeployPackage $App $packageId $effectiveVersion $RootPackageFolder
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

    return New-WinTunerAppUpdateItemResult `
        -Succeeded $true `
        -Name $name `
        -CurrentVersion $currentVersion `
        -TargetVersion $targetVersion `
        -EffectiveVersion $effectiveVersion `
        -GraphId $graphId `
        -PackageId $packageId `
        -ReasonCode 'Updated' `
        -Message "Update completed successfully for $name."
}

function Invoke-WinTunerAppUpdateBatch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateNotNull()][object[]]$Apps,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$RootPackageFolder,
        [scriptblock]$BuildPackage = {
            param($Id, $Folder, $Target, $Installed)
            Invoke-WinTunerPackageBuild `
                -PackageId $Id `
                -PackageFolder $Folder `
                -DesiredVersion $Target `
                -LatestVersion $Target `
                -InstalledVersion $Installed
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
            -DeployPackage $DeployPackage
        $results.Add($itemResult)
    }

    $successCount = @($results | Where-Object Succeeded).Count
    [pscustomobject]@{
        SuccessCount = $successCount
        FailureCount = $results.Count - $successCount
        Results      = @($results)
    }
}

Export-ModuleMember -Function Invoke-WinTunerAppUpdateBatch
