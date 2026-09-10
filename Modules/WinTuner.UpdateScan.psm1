Set-StrictMode -Version Latest

function Invoke-WinTunerUpdateScan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [scriptblock]$GetApps,

        [Parameter(Mandatory)]
        [scriptblock]$ResolvePackageId,

        [Parameter(Mandatory)]
        [scriptblock]$GetVersions,

        [Parameter(Mandatory)]
        [scriptblock]$IsNewerVersion,

        [scriptblock]$ShouldCancel = { $false },

        [scriptblock]$ReportProgress = { param($ProgressInfo) }
    )

    $result = [ordered]@{
        Canceled          = $false
        TotalApps         = 0
        CheckedApps       = 0
        CandidateCount    = 0
        LookupFailureCount = 0
        Candidates        = @()
        ErrorMessage      = $null
    }

    try {
        & $ReportProgress ([pscustomobject]@{
            Stage     = 'Loading'
            Processed = 0
            Total     = 0
            AppName   = $null
        })

        $allApps = @(& $GetApps)
        $appsToCheck = @(
            $allApps | Where-Object {
                if (-not $_) { return $false }
                $currentVersionProperty = $_.PSObject.Properties['CurrentVersion']
                return $currentVersionProperty -and -not [string]::IsNullOrWhiteSpace([string]$currentVersionProperty.Value)
            }
        )
        $result.TotalApps = $appsToCheck.Count

        if (& $ShouldCancel) {
            $result.Canceled = $true
            return [pscustomobject]$result
        }

        $candidates = [System.Collections.Generic.List[object]]::new()

        foreach ($app in $appsToCheck) {
            if (& $ShouldCancel) {
                $result.Canceled = $true
                break
            }

            $result.CheckedApps++
            $nameProperty = $app.PSObject.Properties['Name']
            $currentVersionProperty = $app.PSObject.Properties['CurrentVersion']
            $latestVersionProperty = $app.PSObject.Properties['LatestVersion']
            $appName = if ($nameProperty) { [string]$nameProperty.Value } else { '' }
            $currentVersion = [string]$currentVersionProperty.Value

            & $ReportProgress ([pscustomobject]@{
                Stage     = 'Checking'
                Processed = $result.CheckedApps
                Total     = $result.TotalApps
                AppName   = $appName
            })

            $packageId = $null
            $latestVersion = $null
            $verified = $false
            try {
                $packageId = & $ResolvePackageId $app
            }
            catch {
                $result.LookupFailureCount++
            }

            if (-not [string]::IsNullOrWhiteSpace([string]$packageId)) {
                try {
                    $versions = @(& $GetVersions ([string]$packageId))
                    if ($versions.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace([string]$versions[0])) {
                        $latestVersion = [string]$versions[0]
                        $verified = $true
                    }
                }
                catch {
                    $result.LookupFailureCount++
                }
            }

            if (-not $verified -and $latestVersionProperty -and -not [string]::IsNullOrWhiteSpace([string]$latestVersionProperty.Value)) {
                $latestVersion = [string]$latestVersionProperty.Value
            }

            if (
                -not [string]::IsNullOrWhiteSpace($latestVersion) -and
                (& $IsNewerVersion $latestVersion $currentVersion)
            ) {
                $graphIdProperty = $app.PSObject.Properties['GraphId']
                $graphId = if ($graphIdProperty) { [string]$graphIdProperty.Value } else { $null }

                $candidates.Add([pscustomobject]@{
                    Name           = $appName
                    CurrentVersion = $currentVersion
                    LatestVersion  = $latestVersion
                    GraphId        = $graphId
                    PackageId      = [string]$packageId
                    Checked        = $false
                })
            }
        }

        $result.Candidates = @($candidates)
        $result.CandidateCount = $candidates.Count

        & $ReportProgress ([pscustomobject]@{
            Stage     = if ($result.Canceled) { 'Canceled' } else { 'Completed' }
            Processed = $result.CheckedApps
            Total     = $result.TotalApps
            AppName   = $null
        })
    }
    catch {
        $result.ErrorMessage = $_.Exception.Message
    }

    return [pscustomobject]$result
}

Export-ModuleMember -Function Invoke-WinTunerUpdateScan
