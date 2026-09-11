Set-StrictMode -Version Latest

function Get-WinTunerDiscoveryDeploymentProperty {
    param([Parameter(Mandatory)][object]$InputObject, [Parameter(Mandatory)][string]$Name)
    $property = $InputObject.PSObject.Properties[$Name]
    if ($property) { return [string]$property.Value }
    return ''
}

function Invoke-WinTunerDiscoveryDeployment {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateNotNull()][object[]]$Apps,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$RootPackageFolder,
        [scriptblock]$BuildPackage = {
            param($Id, $Folder, $Version)
            Invoke-WinTunerPackageBuild -PackageId $Id -PackageFolder $Folder -DesiredVersion $Version -LatestVersion $Version
        },
        [scriptblock]$ValidateArtifact = {
            param($Root, $Id, $Version)
            Test-WinTunerPackageArtifact -RootPackageFolder $Root -PackageId $Id -Version $Version
        },
        [scriptblock]$DeployPackage = {
            param($Id, $Version, $Root)
            Deploy-WtWin32App -PackageId $Id -Version $Version -RootPackageFolder $Root -ErrorAction Stop
        },
        [scriptblock]$ReportProgress = { param($ProgressInfo) }
    )

    $items = @($Apps)
    $results = [System.Collections.Generic.List[object]]::new()
    $processed = 0

    foreach ($app in $items) {
        $processed++
        $wingetProperty = $app.PSObject.Properties['WingetApp']
        $wingetApp = if ($wingetProperty) { $wingetProperty.Value } else { $null }
        $name = if ($wingetApp) { Get-WinTunerDiscoveryDeploymentProperty -InputObject $wingetApp -Name 'Name' } else { '' }
        $packageId = if ($wingetApp) { Get-WinTunerDiscoveryDeploymentProperty -InputObject $wingetApp -Name 'PackageID' } else { '' }
        $targetVersion = if ($wingetApp) { Get-WinTunerDiscoveryDeploymentProperty -InputObject $wingetApp -Name 'Version' } else { '' }
        if ([string]::IsNullOrWhiteSpace($name)) { $name = Get-WinTunerDiscoveryDeploymentProperty -InputObject $app -Name 'DisplayName' }

        try {
            & $ReportProgress ([pscustomobject]@{
                Stage = 'Deploying'
                Processed = $processed
                Total = $items.Count
                AppName = $name
            })
        } catch {}

        $reasonCode = ''
        $message = ''
        $effectiveVersion = $targetVersion

        if ([string]::IsNullOrWhiteSpace($packageId)) {
            $reasonCode = 'MissingPackageId'
            $message = "Cannot deploy '$name' because its PackageId is missing."
        } elseif ([string]::IsNullOrWhiteSpace($targetVersion)) {
            $reasonCode = 'MissingTargetVersion'
            $message = "Cannot deploy '$name' because its target version is missing."
        }

        if ([string]::IsNullOrWhiteSpace($reasonCode)) {
            try {
                $buildResult = & $BuildPackage $packageId $RootPackageFolder $targetVersion
            } catch {
                $reasonCode = 'PackageBuildFailed'
                $message = "Package creation failed for '$name': $($_.Exception.Message)"
            }
        }

        if ([string]::IsNullOrWhiteSpace($reasonCode)) {
            if (-not $buildResult -or -not [bool]$buildResult.Succeeded) {
                $detail = if ($buildResult -and $buildResult.ErrorMessage) { [string]$buildResult.ErrorMessage } else { 'The package builder returned no successful result.' }
                $reasonCode = 'PackageBuildFailed'
                $message = "Package creation failed for '$name': $detail"
            } elseif (-not [string]::IsNullOrWhiteSpace([string]$buildResult.EffectiveVersion)) {
                $effectiveVersion = [string]$buildResult.EffectiveVersion
            }
        }

        if ([string]::IsNullOrWhiteSpace($reasonCode)) {
            try {
                $validation = & $ValidateArtifact $RootPackageFolder $packageId $effectiveVersion
            } catch {
                $reasonCode = 'ValidationError'
                $message = "Package validation failed for '$name': $($_.Exception.Message)"
            }
        }

        if ([string]::IsNullOrWhiteSpace($reasonCode) -and (-not $validation -or -not [bool]$validation.IsValid)) {
            $reasonCode = if ($validation -and $validation.ReasonCode) { [string]$validation.ReasonCode } else { 'InvalidArtifact' }
            $detail = if ($validation -and $validation.Reason) { [string]$validation.Reason } else { 'The exact package artifact is invalid.' }
            $message = "Package validation failed for '$name' ($reasonCode): $detail"
        }

        if ([string]::IsNullOrWhiteSpace($reasonCode)) {
            try {
                $null = & $DeployPackage $packageId $effectiveVersion $RootPackageFolder
            } catch {
                $reasonCode = 'DeploymentFailed'
                $message = "Deployment failed for '$name': $($_.Exception.Message)"
            }
        }
        $succeeded = [string]::IsNullOrWhiteSpace($reasonCode)
        $results.Add([pscustomobject]@{
            Succeeded = $succeeded
            Name = $name
            PackageId = $packageId
            TargetVersion = $targetVersion
            EffectiveVersion = $effectiveVersion
            ReasonCode = if ($succeeded) { 'Deployed' } else { $reasonCode }
            Message = if ($succeeded) { "Deployment completed successfully for '$name'." } else { $message }
        })
    }

    $successCount = @($results | Where-Object Succeeded).Count
    [pscustomobject]@{
        SuccessCount = $successCount
        FailureCount = $results.Count - $successCount
        Results = @($results)
    }
}

Export-ModuleMember -Function Invoke-WinTunerDiscoveryDeployment
