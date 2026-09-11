[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^[^@]+@[^@]+$')]
    [string]$UserPrincipalName,

    [switch]$ForceFreshDiscovery,

    [switch]$ValidateDiscoveryCache
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$repositoryRoot = Split-Path -Parent $PSScriptRoot

if ($ForceFreshDiscovery -and $ValidateDiscoveryCache) {
    throw 'Use either ForceFreshDiscovery or ValidateDiscoveryCache, not both.'
}

Import-Module WinTuner -ErrorAction Stop
Import-Module (Join-Path (Join-Path $repositoryRoot 'Modules') 'WinTuner.Connection.psm1') -Force -ErrorAction Stop
Import-Module (Join-Path (Join-Path $repositoryRoot 'Modules') 'WinTuner.Intune.psm1') -Force -ErrorAction Stop

$managedState = [pscustomobject]@{ Apps = @() }
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

try {
    Write-Host "Connecting WinTuner as $UserPrincipalName..."
    $null = Connect-WtWinTuner -Username $UserPrincipalName -ErrorAction Stop

    $verification = Invoke-WinTunerConnectionVerification -GetApps {
        $managedState.Apps = @(Get-WtWin32Apps -Update:$false -Superseded:$false -ErrorAction Stop)
        $managedState.Apps
    }.GetNewClosure() -MaxAttempts 4 -RetryDelayMilliseconds 500

    if (-not $verification.Succeeded) {
        throw "Tenant verification failed after $($verification.Attempts) attempt(s): $($verification.ErrorMessage)"
    }

    Write-Host 'Connecting Microsoft Graph...'
    $graphContext = Connect-WinTunerGraph -UserPrincipalName $UserPrincipalName

    $detectedParameters = @{
        PageSize = 500
        MaxPages = 1000
    }

    $freshDetectedResult = $null
    $cachedDetectedResult = $null
    if ($ValidateDiscoveryCache) {
        Write-Host 'Clearing the local detected-app cache before the two-pass validation...'
        Clear-WinTunerDetectedAppsCache

        Write-Host 'Reading fresh detected applications from Microsoft Graph...'
        $freshDetectedResult = Get-WinTunerDetectedApps @detectedParameters -ForceRefresh

        if ($freshDetectedResult.FromCache) {
            throw 'Discovery cache validation failed: the first Graph read was expected to be fresh.'
        }

        if ($freshDetectedResult.LimitReached) {
            throw 'Discovery cache validation failed: the fresh Graph read reached the configured pagination limit and was not cached.'
        }

        Write-Host 'Reading detected applications again to verify persistent cache reuse...'
        $cachedDetectedResult = Get-WinTunerDetectedApps @detectedParameters

        if (-not $cachedDetectedResult.FromCache) {
            throw 'Discovery cache validation failed: the second Graph read did not use the persistent cache.'
        }

        if (@($freshDetectedResult.Apps).Count -ne @($cachedDetectedResult.Apps).Count) {
            throw "Discovery cache validation failed: fresh and cached counts differ ($(@($freshDetectedResult.Apps).Count) vs $(@($cachedDetectedResult.Apps).Count))."
        }

        $detectedResult = $cachedDetectedResult
    } else {
        if ($ForceFreshDiscovery) {
            $detectedParameters.ForceRefresh = $true
        }

        Write-Host 'Reading detected applications from Microsoft Graph...'
        $detectedResult = Get-WinTunerDetectedApps @detectedParameters
    }

    $stopwatch.Stop()
    [pscustomobject]@{
        Succeeded = $true
        UserPrincipalName = $UserPrincipalName
        TenantId = [string]$graphContext.TenantId
        ManagedWin32Apps = @($managedState.Apps).Count
        DetectedApps = @($detectedResult.Apps).Count
        GraphSource = if ($detectedResult.FromCache) { 'Cached' } else { 'Fresh' }
        GraphPages = [int]$detectedResult.PageCount
        GraphLimitReached = [bool]$detectedResult.LimitReached
        GraphDataAgeMinutes = [math]::Round([double]$detectedResult.AgeMinutes, 1)
        GraphRuns = if ($ValidateDiscoveryCache) { 2 } else { 1 }
        FreshDetectedApps = if ($freshDetectedResult) { @($freshDetectedResult.Apps).Count } else { $null }
        CachedDetectedApps = if ($cachedDetectedResult) { @($cachedDetectedResult.Apps).Count } else { $null }
        CacheReuseVerified = [bool]$ValidateDiscoveryCache
        DurationSeconds = [math]::Round($stopwatch.Elapsed.TotalSeconds, 1)
        TenantWritesPerformed = 0
    }
} finally {
    try { Disconnect-WinTunerGraph } catch {}
    try { Disconnect-WtWinTuner -ErrorAction SilentlyContinue } catch {}
}
