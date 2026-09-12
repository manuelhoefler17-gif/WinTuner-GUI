[CmdletBinding()]
param(
    [ValidateSet('Interactive', 'ClientSecret', 'Certificate')]
    [string]$AuthenticationMode = 'Interactive',

    [AllowNull()]
    [string]$UserPrincipalName,

    [AllowNull()]
    [string]$TenantId,

    [AllowNull()]
    [string]$ClientId,

    [AllowNull()]
    [securestring]$ClientSecret,

    [AllowNull()]
    [string]$CertificateThumbprint,

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
Import-Module (Join-Path (Join-Path $repositoryRoot 'Modules') 'WinTuner.Authentication.psm1') -Force -ErrorAction Stop
Import-Module (Join-Path (Join-Path $repositoryRoot 'Modules') 'WinTuner.Connection.psm1') -Force -ErrorAction Stop
Import-Module (Join-Path (Join-Path $repositoryRoot 'Modules') 'WinTuner.Intune.psm1') -Force -ErrorAction Stop

$managedState = [pscustomobject]@{ Apps = @() }
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$plainClientSecret = $null
$connectionParameters = $null
$tokenCacheReset = $null
$permissionPreflight = $null

try {
    if ($AuthenticationMode -eq 'ClientSecret') {
        if (-not $ClientSecret) {
            throw 'ClientSecret mode requires -ClientSecret as a SecureString. Use Read-Host -AsSecureString.'
        }
        $secretPointer = [IntPtr]::Zero
        try {
            $secretPointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($ClientSecret)
            $plainClientSecret = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($secretPointer)
        } finally {
            if ($secretPointer -ne [IntPtr]::Zero) {
                [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($secretPointer)
            }
        }
    }

    $configuration = Test-WinTunerAuthenticationConfiguration -Mode $AuthenticationMode -UserPrincipalName $UserPrincipalName -TenantId $TenantId -ClientId $ClientId -ClientSecret $plainClientSecret -CertificateThumbprint $CertificateThumbprint -RequireSecret:($AuthenticationMode -eq 'ClientSecret')
    if (-not $configuration.IsValid) {
        throw $configuration.Reason
    }
    if ($AuthenticationMode -eq 'Certificate') {
        $certificateValidation = Test-WinTunerCertificateAvailable -Thumbprint $configuration.CertificateThumbprint
        if (-not $certificateValidation.IsAvailable) {
            throw $certificateValidation.Reason
        }
    }

    $identity = if ($AuthenticationMode -eq 'Interactive') { $configuration.IdentityLabel } else { "$AuthenticationMode app $($configuration.ClientId)" }
    Write-Host "Connecting WinTuner using $identity..."
    $connectionParameters = New-WinTunerModuleConnectionParameters -Mode $AuthenticationMode -UserPrincipalName $UserPrincipalName -TenantId $configuration.TenantId -ClientId $configuration.ClientId -ClientSecret $plainClientSecret -CertificateThumbprint $configuration.CertificateThumbprint
    if ($AuthenticationMode -in @('ClientSecret', 'Certificate')) {
        try { Disconnect-WtWinTuner -ErrorAction SilentlyContinue } catch {}
        $tokenCacheReset = Clear-WinTunerAppOnlyTokenCache
    }
    $null = Connect-WtWinTuner @connectionParameters

    Write-Host 'Connecting Microsoft Graph...'
    switch ($AuthenticationMode) {
        'Interactive' {
            $graphContext = Connect-WinTunerGraph -AuthenticationMode Interactive -UserPrincipalName $UserPrincipalName
        }
        'ClientSecret' {
            $graphContext = Connect-WinTunerGraph -AuthenticationMode ClientSecret -TenantId $configuration.TenantId -ClientId $configuration.ClientId -ClientSecret $plainClientSecret
        }
        'Certificate' {
            $graphContext = Connect-WinTunerGraph -AuthenticationMode Certificate -TenantId $configuration.TenantId -ClientId $configuration.ClientId -CertificateThumbprint $configuration.CertificateThumbprint
        }
    }

    if ($AuthenticationMode -in @('ClientSecret', 'Certificate')) {
        Write-Host 'Checking both required Microsoft Graph Application permissions...'
        $permissionPreflight = Invoke-WinTunerGraphPermissionPreflight -GetManagedApps {
            Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps?$top=1' -ErrorAction Stop
        } -GetDetectedApps {
            Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/deviceManagement/detectedApps?$top=1' -ErrorAction Stop
        }
        if (-not $permissionPreflight.Succeeded) {
            throw $permissionPreflight.ErrorMessage
        }
    }

    $verification = Invoke-WinTunerConnectionVerification -GetApps {
        $managedState.Apps = @(Get-WtWin32Apps -Update:$false -Superseded:$false -ErrorAction Stop)
        $managedState.Apps
    }.GetNewClosure() -MaxAttempts 4 -RetryDelayMilliseconds 500

    if (-not $verification.Succeeded) {
        throw "Tenant verification failed after $($verification.Attempts) attempt(s): $($verification.ErrorMessage)"
    }

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
        AuthenticationMode = $AuthenticationMode
        UserPrincipalName = if ($AuthenticationMode -eq 'Interactive') { $UserPrincipalName } else { '' }
        ClientId = if ($AuthenticationMode -eq 'Interactive') { '' } else { $configuration.ClientId }
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
        PermissionPreflightPassed = [bool]($permissionPreflight -and $permissionPreflight.Succeeded)
        WinTunerAppOnlyCacheCleared = [bool]($tokenCacheReset -and $tokenCacheReset.Cleared)
        DurationSeconds = [math]::Round($stopwatch.Elapsed.TotalSeconds, 1)
        TenantWritesPerformed = 0
    }
} finally {
    if ($connectionParameters -and $connectionParameters.ContainsKey('ClientSecret')) {
        $connectionParameters['ClientSecret'] = $null
        [void]$connectionParameters.Remove('ClientSecret')
    }
    $plainClientSecret = $null
    try { Disconnect-WinTunerGraph } catch {}
    try { Disconnect-WtWinTuner -ErrorAction SilentlyContinue } catch {}
}
