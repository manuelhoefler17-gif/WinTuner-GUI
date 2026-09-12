Set-StrictMode -Version Latest

function Get-WinTunerGraphContextValidation {
    [CmdletBinding()]
    param(
        [ValidateSet('Interactive', 'ClientSecret', 'Certificate')]
        [string]$AuthenticationMode = 'Interactive',
        [AllowNull()][string]$UserPrincipalName,
        [AllowNull()][string]$TenantId,
        [AllowNull()][string]$ClientId,
        [AllowNull()][object]$Context
    )

    if (-not $Context) {
        $Context = Get-MgContext -ErrorAction SilentlyContinue
    }
    if (-not $Context) {
        return [pscustomobject]@{ IsValid = $false; Reason = 'No Microsoft Graph process context is available.'; Context = $null }
    }

    if ($AuthenticationMode -eq 'Interactive') {
        if ([string]::IsNullOrWhiteSpace($UserPrincipalName)) {
            return [pscustomobject]@{ IsValid = $false; Reason = 'The expected user principal name is empty.'; Context = $Context }
        }
        if (-not [string]::Equals([string]$Context.Account, [string]$UserPrincipalName, [System.StringComparison]::OrdinalIgnoreCase)) {
            return [pscustomobject]@{ IsValid = $false; Reason = 'The Microsoft Graph account does not match the requested user.'; Context = $Context }
        }
    } else {
        if ([string]::IsNullOrWhiteSpace($ClientId)) {
            return [pscustomobject]@{ IsValid = $false; Reason = 'The expected Entra client ID is empty.'; Context = $Context }
        }
        if (-not [string]::Equals([string]$Context.ClientId, [string]$ClientId, [System.StringComparison]::OrdinalIgnoreCase)) {
            return [pscustomobject]@{ IsValid = $false; Reason = 'The Microsoft Graph client ID does not match the configured Entra application.'; Context = $Context }
        }

        $expectedTenantGuid = [guid]::Empty
        if ([guid]::TryParse([string]$TenantId, [ref]$expectedTenantGuid)) {
            if (-not [string]::Equals([string]$Context.TenantId, $expectedTenantGuid.ToString(), [System.StringComparison]::OrdinalIgnoreCase)) {
                return [pscustomobject]@{ IsValid = $false; Reason = 'The Microsoft Graph tenant does not match the configured tenant ID.'; Context = $Context }
            }
        }

        $authTypeProperty = $Context.PSObject.Properties['AuthType']
        if (-not $authTypeProperty -or [string]$authTypeProperty.Value -ne 'AppOnly') {
            return [pscustomobject]@{ IsValid = $false; Reason = 'The Microsoft Graph context is not an app-only context.'; Context = $Context }
        }
    }

    return [pscustomobject]@{ IsValid = $true; Reason = ''; Context = $Context }
}

function Confirm-WinTunerGraphContext {
    [CmdletBinding()]
    param(
        [ValidateSet('Interactive', 'ClientSecret', 'Certificate')]
        [string]$AuthenticationMode = 'Interactive',
        [AllowNull()][string]$UserPrincipalName,
        [AllowNull()][string]$TenantId,
        [AllowNull()][string]$ClientId
    )

    $validation = Get-WinTunerGraphContextValidation -AuthenticationMode $AuthenticationMode -UserPrincipalName $UserPrincipalName -TenantId $TenantId -ClientId $ClientId
    if (-not $validation.IsValid) {
        throw $validation.Reason
    }
    return $validation.Context
}

function Connect-WinTunerGraph {
    [CmdletBinding()]
    param(
        [ValidateSet('Interactive', 'ClientSecret', 'Certificate')]
        [string]$AuthenticationMode = 'Interactive',
        [AllowNull()][string]$UserPrincipalName,
        [AllowNull()][string]$TenantId,
        [AllowNull()][string]$ClientId,
        [AllowNull()][string]$ClientSecret,
        [AllowNull()][string]$CertificateThumbprint,
        [string[]]$Scopes = @(
            'DeviceManagementApps.ReadWrite.All',
            'DeviceManagementManagedDevices.Read.All',
            'Directory.Read.All'
        )
    )

    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        throw 'Microsoft.Graph.Authentication module not found. Install Microsoft.Graph first.'
    }

    if ($AuthenticationMode -eq 'Interactive') {
        if ($UserPrincipalName -notmatch '^[^@\s]+@[^@\s]+$') {
            throw "Invalid user principal name: '$UserPrincipalName'."
        }
        $TenantId = $UserPrincipalName.Split('@')[1]
    } else {
        if ([string]::IsNullOrWhiteSpace($TenantId)) {
            throw 'Tenant ID is required for app-only Microsoft Graph authentication.'
        }
        $parsedClientId = [guid]::Empty
        if (-not [guid]::TryParse([string]$ClientId, [ref]$parsedClientId)) {
            throw 'A valid Entra application (client) ID is required for app-only Microsoft Graph authentication.'
        }
        if ($AuthenticationMode -eq 'ClientSecret' -and [string]::IsNullOrWhiteSpace($ClientSecret)) {
            throw 'Client secret is required for client-secret Microsoft Graph authentication.'
        }
        if ($AuthenticationMode -eq 'Certificate' -and [string]::IsNullOrWhiteSpace($CertificateThumbprint)) {
            throw 'Certificate thumbprint is required for certificate Microsoft Graph authentication.'
        }
    }

    $context = Get-MgContext -ErrorAction SilentlyContinue
    if ($context) {
        $existingValidation = Get-WinTunerGraphContextValidation -AuthenticationMode $AuthenticationMode -UserPrincipalName $UserPrincipalName -TenantId $TenantId -ClientId $ClientId -Context $context
        if ($existingValidation.IsValid) {
            if ($AuthenticationMode -ne 'Interactive' -or @($Scopes | Where-Object { $context.Scopes -notcontains $_ }).Count -eq 0) {
                return $context
            }
        }

        try {
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        } catch {}
    }

    switch ($AuthenticationMode) {
        'Interactive' {
            $null = Connect-MgGraph -TenantId $TenantId -Scopes $Scopes -ContextScope Process -NoWelcome -ErrorAction Stop *>&1
        }
        'ClientSecret' {
            $secureSecret = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
            $credential = [System.Management.Automation.PSCredential]::new($ClientId, $secureSecret)
            try {
                $null = Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $credential -ContextScope Process -NoWelcome -ErrorAction Stop *>&1
            } finally {
                $credential = $null
                $secureSecret = $null
                $ClientSecret = $null
            }
        }
        'Certificate' {
            $normalizedThumbprint = ($CertificateThumbprint -replace '[^0-9A-Fa-f]', '').ToUpperInvariant()
            $null = Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $normalizedThumbprint -ContextScope Process -NoWelcome -ErrorAction Stop *>&1
        }
    }

    $context = Get-MgContext -ErrorAction Stop
    $validation = Get-WinTunerGraphContextValidation -AuthenticationMode $AuthenticationMode -UserPrincipalName $UserPrincipalName -TenantId $TenantId -ClientId $ClientId -Context $context
    if (-not $validation.IsValid) {
        throw "Microsoft Graph authentication returned an unexpected context: $($validation.Reason)"
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
    Get-WinTunerGraphContextValidation, `
    Confirm-WinTunerGraphContext, `
    Get-WinTunerDetectedApps, `
    Clear-WinTunerDetectedAppsCache, `
    Disconnect-WinTunerGraph
