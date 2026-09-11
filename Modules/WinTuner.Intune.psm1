Set-StrictMode -Version Latest

function Connect-WinTunerGraph {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$UserPrincipalName,

        [string[]]$Scopes = @(
            'DeviceManagementApps.ReadWrite.All',
            'DeviceManagementManagedDevices.Read.All',
            'Directory.Read.All'
        )
    )

    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        throw "Microsoft.Graph.Authentication module not found. Install Microsoft.Graph first."
    }

    if ($UserPrincipalName -notmatch '^[^@\s]+@[^@\s]+$') {
        throw "Invalid user principal name: '$UserPrincipalName'."
    }

    $context = Get-MgContext -ErrorAction SilentlyContinue
    $needsAuth = $false

    if (-not $context) {
        $needsAuth = $true
    }
    else {
        foreach ($scope in $Scopes) {
            if ($context.Scopes -notcontains $scope) {
                $needsAuth = $true
                break
            }
        }

        if ($context.Account -ne $UserPrincipalName) {
            $needsAuth = $true
        }
    }

    if ($needsAuth) {
        if ($context) {
            try {
                Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            }
            catch {
            }
        }

        $tenantDomain = $UserPrincipalName.Split('@')[1]

        $null = Connect-MgGraph `
            -TenantId $tenantDomain `
            -Scopes $Scopes `
            -ContextScope Process `
            -NoWelcome `
            -ErrorAction Stop *>&1

        $context = Get-MgContext -ErrorAction Stop
    }

    return $context
}


$script:detectedAppsCachePath = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'WinTuner_DetectedAppsCache.json'

function Get-WinTunerDetectedAppsCache {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TenantId,

        [ValidateRange(1,1440)]
        [int]$CacheTtlMinutes = 360
    )

    if (-not (Test-Path $script:detectedAppsCachePath)) {
        return $null
    }

    try {
        $raw = Get-Content `
            -Path $script:detectedAppsCachePath `
            -Raw `
            -Encoding utf8 `
            -ErrorAction Stop

        if ([string]::IsNullOrWhiteSpace($raw)) {
            return $null
        }

        $cache = $raw | ConvertFrom-Json -ErrorAction Stop

        if ([string]$cache.TenantId -ne $TenantId) {
            return $null
        }

        $timestamp = [datetime]::Parse(
            [string]$cache.Timestamp,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::RoundtripKind
        )

        $ageMinutes = (
            [datetime]::UtcNow -
            $timestamp.ToUniversalTime()
        ).TotalMinutes

        if ($ageMinutes -ge $CacheTtlMinutes) {
            return $null
        }

        return [pscustomobject]@{
            Apps      = @($cache.Apps)
            PageCount = [int]$cache.PageCount
            Timestamp = $timestamp
            AgeMinutes = [math]::Max(0, $ageMinutes)
        }
    }
    catch {
        Write-Verbose "Could not read detected apps cache: $($_.Exception.Message)"
        return $null
    }
}

function Save-WinTunerDetectedAppsCache {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TenantId,

        [Parameter(Mandatory)]
        [object[]]$Apps,

        [Parameter(Mandatory)]
        [int]$PageCount
    )

    try {
        $cacheObject = [ordered]@{
            Version   = 1
            TenantId  = $TenantId
            Timestamp = [datetime]::UtcNow.ToString('o')
            PageCount = $PageCount
            Apps      = @($Apps)
        }

        $tempPath = "$($script:detectedAppsCachePath).tmp"

        $cacheObject |
            ConvertTo-Json -Depth 12 |
            Set-Content `
                -Path $tempPath `
                -Encoding utf8 `
                -ErrorAction Stop

        Move-Item `
            -Path $tempPath `
            -Destination $script:detectedAppsCachePath `
            -Force `
            -ErrorAction Stop
    }
    catch {
        Write-Verbose "Could not save detected apps cache: $($_.Exception.Message)"
    }
}

function Clear-WinTunerDetectedAppsCache {
    [CmdletBinding()]
    param()

    Remove-Item `
        -Path $script:detectedAppsCachePath `
        -Force `
        -ErrorAction SilentlyContinue
}

function Get-WinTunerDetectedApps {
    [CmdletBinding()]
    param(
        [ValidateRange(1,1000)]
        [int]$PageSize = 500,

        [ValidateRange(1,1000)]
        [int]$MaxPages = 100,

        [ValidateRange(0,10000)]
        [int]$PageDelayMs = 750,

        [ValidateRange(1,10)]
        [int]$MaxThrottleRetries = 5,

        [ValidateRange(1,1440)]
        [int]$CacheTtlMinutes = 360,

        [switch]$ForceRefresh
    )

    if (-not (Get-Command Invoke-MgRestMethod -ErrorAction SilentlyContinue)) {
        throw "Invoke-MgRestMethod is not available. Connect/import Microsoft.Graph.Authentication first."
    }

    $context = Get-MgContext -ErrorAction SilentlyContinue
    $tenantId = $null

    if ($context -and $context.TenantId) {
        $tenantId = [string]$context.TenantId
    }

    if (-not $ForceRefresh -and $tenantId) {
        $cachedResult = Get-WinTunerDetectedAppsCache `
            -TenantId $tenantId `
            -CacheTtlMinutes $CacheTtlMinutes

        if ($cachedResult) {
            Write-Verbose "Using cached Intune detected apps ($($cachedResult.Apps.Count) apps)."

            return [pscustomobject]@{
                Apps         = @($cachedResult.Apps)
                PageCount    = $cachedResult.PageCount
                LimitReached = $false
                FromCache    = $true
                RetrievedAt = $cachedResult.Timestamp
                AgeMinutes = [math]::Round([double]$cachedResult.AgeMinutes, 1)
            }
        }
    }

    $uri = "https://graph.microsoft.com/beta/deviceManagement/detectedApps?`$top=$PageSize&`$orderby=deviceCount desc"

    $apps = [System.Collections.Generic.List[object]]::new()
    $pageCount = 0
    $limitReached = $false

    do {
        $response = $null
        $retry = 0

        while ($null -eq $response) {
            try {
                $response = Invoke-MgRestMethod `
                    -Uri $uri `
                    -Method GET `
                    -ErrorAction Stop `
                    2>$null 3>$null 4>$null 5>$null 6>$null
            }
            catch {
                $message = $_.Exception.Message

                if ($message -notmatch 'TooManyRequests|429') {
                    throw
                }

                $retry++

                if ($retry -gt $MaxThrottleRetries) {
                    throw "Intune throttling persisted after $MaxThrottleRetries retries. Last error: $message"
                }

                $delaySeconds = [math]::Min(
                    60,
                    [math]::Pow(2, $retry + 1)
                )

                Write-Verbose "Intune throttling detected. Waiting $delaySeconds seconds before retry $retry/$MaxThrottleRetries."
                Start-Sleep -Seconds $delaySeconds
            }
        }

        if ($response.value) {
            $apps.AddRange([object[]]$response.value)
        }

        $nextLink = $null

        if ($response -is [System.Collections.IDictionary]) {
            if ($response.Contains('@odata.nextLink')) {
                $nextLink = [string]$response['@odata.nextLink']
            }
        }
        else {
            $nextLinkProperty = $response.PSObject.Properties['@odata.nextLink']

            if ($nextLinkProperty) {
                $nextLink = [string]$nextLinkProperty.Value
            }
        }

        $uri = $nextLink
        $pageCount++

        if ($pageCount -ge $MaxPages -and $uri) {
            $limitReached = $true
            break
        }

        if ($uri -and $PageDelayMs -gt 0) {
            Start-Sleep -Milliseconds $PageDelayMs
        }

    } while ($uri)

    $retrievedAt = [datetime]::UtcNow
    if (-not $limitReached -and $tenantId) {
        Save-WinTunerDetectedAppsCache `
            -TenantId $tenantId `
            -Apps @($apps) `
            -PageCount $pageCount
    }

    return [pscustomobject]@{
        Apps         = $apps
        PageCount    = $pageCount
        LimitReached = $limitReached
        FromCache    = $false
        RetrievedAt = $retrievedAt
        AgeMinutes = 0
    }
}

function Disconnect-WinTunerGraph {
    [CmdletBinding()]
    param()

    try {
        if (Get-MgContext -ErrorAction SilentlyContinue) {
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        }
    }
    catch {
    }
}


Export-ModuleMember -Function `
    Connect-WinTunerGraph, `
    Get-WinTunerDetectedApps, `
    Clear-WinTunerDetectedAppsCache, `
    Disconnect-WinTunerGraph
