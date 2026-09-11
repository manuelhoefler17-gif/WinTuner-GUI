[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^[^@]+@[^@]+$')]
    [string]$UserPrincipalName,

    [switch]$ForceFreshDiscovery
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$repositoryRoot = Split-Path -Parent $PSScriptRoot

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
    if ($ForceFreshDiscovery) {
        $detectedParameters.ForceRefresh = $true
    }

    Write-Host 'Reading detected applications from Microsoft Graph...'
    $detectedResult = Get-WinTunerDetectedApps @detectedParameters

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
        DurationSeconds = [math]::Round($stopwatch.Elapsed.TotalSeconds, 1)
        TenantWritesPerformed = 0
    }
} finally {
    try { Disconnect-WinTunerGraph } catch {}
    try { Disconnect-WtWinTuner -ErrorAction SilentlyContinue } catch {}
}