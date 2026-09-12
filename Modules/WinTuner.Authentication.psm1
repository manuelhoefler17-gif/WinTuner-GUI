Set-StrictMode -Version Latest

$script:WinTunerClientSecretEntropy = [System.Text.Encoding]::UTF8.GetBytes('WinTunerGUI/EntraClientSecret/v1')

function Protect-WinTunerClientSecretForCurrentUser {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ClientSecret)

    if ([string]::IsNullOrWhiteSpace($ClientSecret)) {
        throw 'The client secret cannot be empty.'
    }

    $plainBytes = $null
    $protectedBytes = $null
    try {
        $plainBytes = [System.Text.Encoding]::UTF8.GetBytes($ClientSecret)
        $protectedBytes = [System.Security.Cryptography.ProtectedData]::Protect(
            $plainBytes,
            $script:WinTunerClientSecretEntropy,
            [System.Security.Cryptography.DataProtectionScope]::CurrentUser
        )
        return [Convert]::ToBase64String($protectedBytes)
    } finally {
        if ($plainBytes) { [Array]::Clear($plainBytes, 0, $plainBytes.Length) }
        if ($protectedBytes) { [Array]::Clear($protectedBytes, 0, $protectedBytes.Length) }
    }
}

function Unprotect-WinTunerClientSecretForCurrentUser {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ProtectedClientSecret)

    if ([string]::IsNullOrWhiteSpace($ProtectedClientSecret)) {
        throw 'No protected client secret is stored.'
    }

    $protectedBytes = $null
    $plainBytes = $null
    try {
        $protectedBytes = [Convert]::FromBase64String($ProtectedClientSecret)
        $plainBytes = [System.Security.Cryptography.ProtectedData]::Unprotect(
            $protectedBytes,
            $script:WinTunerClientSecretEntropy,
            [System.Security.Cryptography.DataProtectionScope]::CurrentUser
        )
        return [System.Text.Encoding]::UTF8.GetString($plainBytes)
    } catch {
        throw 'The saved client secret cannot be decrypted for this Windows user. Open Authentication settings and save it again.'
    } finally {
        if ($protectedBytes) { [Array]::Clear($protectedBytes, 0, $protectedBytes.Length) }
        if ($plainBytes) { [Array]::Clear($plainBytes, 0, $plainBytes.Length) }
    }
}

function ConvertTo-WinTunerCertificateThumbprint {
    [CmdletBinding()]
    param([AllowNull()][string]$Thumbprint)

    if ([string]::IsNullOrWhiteSpace($Thumbprint)) {
        return ''
    }

    return ($Thumbprint -replace '[^0-9A-Fa-f]', '').ToUpperInvariant()
}

function Test-WinTunerAuthenticationConfiguration {
    [CmdletBinding()]
    param(
        [ValidateSet('Interactive', 'ClientSecret', 'Certificate')]
        [string]$Mode = 'Interactive',

        [AllowNull()][string]$UserPrincipalName,
        [AllowNull()][string]$TenantId,
        [AllowNull()][string]$ClientId,
        [AllowNull()][string]$ClientSecret,
        [AllowNull()][string]$CertificateThumbprint,
        [switch]$RequireSecret
    )

    $reason = ''
    $identityLabel = ''
    $normalizedThumbprint = ConvertTo-WinTunerCertificateThumbprint -Thumbprint $CertificateThumbprint
    $parsedClientId = [guid]::Empty
    $clientIdIsValid = [guid]::TryParse([string]$ClientId, [ref]$parsedClientId)

    switch ($Mode) {
        'Interactive' {
            $upn = [string]$UserPrincipalName
            if ($upn -notmatch '^[^@\s]+@[^@\s]+$') {
                $reason = 'Enter a valid Microsoft 365 user principal name.'
            } else {
                $identityLabel = $upn
            }
        }
        'ClientSecret' {
            if ([string]::IsNullOrWhiteSpace($TenantId)) {
                $reason = 'Enter the tenant ID or verified tenant domain.'
            } elseif ([string]::IsNullOrWhiteSpace($ClientId) -or -not $clientIdIsValid) {
                $reason = 'Enter a valid Entra application (client) ID.'
            } elseif ($RequireSecret -and [string]::IsNullOrWhiteSpace($ClientSecret)) {
                $reason = 'Enter the client secret for this login.'
            } else {
                $identityLabel = "App-only client secret: $ClientId"
            }
        }
        'Certificate' {
            if ([string]::IsNullOrWhiteSpace($TenantId)) {
                $reason = 'Enter the tenant ID or verified tenant domain.'
            } elseif ([string]::IsNullOrWhiteSpace($ClientId) -or -not $clientIdIsValid) {
                $reason = 'Enter a valid Entra application (client) ID.'
            } elseif ($normalizedThumbprint -notmatch '^[0-9A-F]{40,64}$') {
                $reason = 'Enter a valid certificate thumbprint (40 to 64 hexadecimal characters).'
            } else {
                $identityLabel = "App-only certificate: $ClientId"
            }
        }
    }

    return [pscustomobject]@{
        IsValid = [string]::IsNullOrWhiteSpace($reason)
        Reason = $reason
        Mode = $Mode
        IdentityLabel = $identityLabel
        TenantId = ([string]$TenantId).Trim()
        ClientId = ([string]$ClientId).Trim()
        CertificateThumbprint = $normalizedThumbprint
    }
}

function New-WinTunerModuleConnectionParameters {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateSet('Interactive', 'ClientSecret', 'Certificate')][string]$Mode,
        [AllowNull()][string]$UserPrincipalName,
        [AllowNull()][string]$TenantId,
        [AllowNull()][string]$ClientId,
        [AllowNull()][string]$ClientSecret,
        [AllowNull()][string]$CertificateThumbprint
    )

    $validation = Test-WinTunerAuthenticationConfiguration -Mode $Mode -UserPrincipalName $UserPrincipalName -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret -CertificateThumbprint $CertificateThumbprint -RequireSecret:($Mode -eq 'ClientSecret')
    if (-not $validation.IsValid) {
        throw $validation.Reason
    }

    switch ($Mode) {
        'Interactive' {
            return @{ Username = [string]$UserPrincipalName; ErrorAction = 'Stop' }
        }
        'ClientSecret' {
            return @{
                TenantId = $validation.TenantId
                ClientId = $validation.ClientId
                ClientSecret = [string]$ClientSecret
                ErrorAction = 'Stop'
            }
        }
        'Certificate' {
            return @{
                TenantId = $validation.TenantId
                ClientId = $validation.ClientId
                ClientCertificateThumbprint = $validation.CertificateThumbprint
                ErrorAction = 'Stop'
            }
        }
    }
}

function Protect-WinTunerAuthenticationErrorMessage {
    [CmdletBinding()]
    param(
        [AllowNull()][string]$Message,
        [AllowNull()][string]$Secret
    )

    if ([string]::IsNullOrEmpty($Message) -or [string]::IsNullOrEmpty($Secret)) {
        return [string]$Message
    }

    return $Message.Replace($Secret, '[REDACTED]')
}

function Get-WinTunerAppOnlyTokenCachePath {
    [CmdletBinding()]
    param([AllowNull()][string]$LocalApplicationDataPath)

    $root = if ([string]::IsNullOrWhiteSpace($LocalApplicationDataPath)) {
        [Environment]::GetFolderPath('LocalApplicationData')
    } else {
        $LocalApplicationDataPath
    }
    if ([string]::IsNullOrWhiteSpace($root)) {
        throw 'The current user local application data path is unavailable.'
    }

    return [System.IO.Path]::Combine(
        [System.IO.Path]::GetFullPath($root),
        '.IdentityService',
        'WinTuner-PowerShell-CC.nocae'
    )
}

function Clear-WinTunerAppOnlyTokenCache {
    [CmdletBinding()]
    param([AllowNull()][string]$LocalApplicationDataPath)

    $cachePath = Get-WinTunerAppOnlyTokenCachePath -LocalApplicationDataPath $LocalApplicationDataPath
    $identityServicePath = Split-Path -Parent $cachePath
    if (Test-Path -LiteralPath $identityServicePath) {
        $identityServiceItem = Get-Item -LiteralPath $identityServicePath -Force -ErrorAction Stop
        if (-not $identityServiceItem.PSIsContainer) {
            throw 'The WinTuner IdentityService path unexpectedly points to a file.'
        }
        if (($identityServiceItem.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw 'The WinTuner IdentityService path unexpectedly points to a reparse point.'
        }
    }
    if (-not (Test-Path -LiteralPath $cachePath)) {
        return [pscustomobject]@{
            CachePath = $cachePath
            Existed = $false
            Cleared = $false
        }
    }

    $cacheItem = Get-Item -LiteralPath $cachePath -Force -ErrorAction Stop
    if ($cacheItem.PSIsContainer) {
        throw 'The WinTuner app-only token cache path unexpectedly points to a directory.'
    }
    if (($cacheItem.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw 'The WinTuner app-only token cache path unexpectedly points to a reparse point.'
    }

    Remove-Item -LiteralPath $cacheItem.FullName -Force -ErrorAction Stop
    if (Test-Path -LiteralPath $cachePath) {
        throw 'The WinTuner app-only token cache could not be cleared.'
    }

    return [pscustomobject]@{
        CachePath = $cachePath
        Existed = $true
        Cleared = $true
    }
}

function Test-WinTunerCertificateAvailable {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Thumbprint)

    $normalized = ConvertTo-WinTunerCertificateThumbprint -Thumbprint $Thumbprint
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return [pscustomobject]@{ IsAvailable = $false; Reason = 'Certificate thumbprint is empty.'; Certificate = $null }
    }

    $certificate = Get-ChildItem -LiteralPath "Cert:\CurrentUser\My\$normalized" -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $certificate) {
        return [pscustomobject]@{ IsAvailable = $false; Reason = "Certificate $normalized was not found in CurrentUser\My."; Certificate = $null }
    }
    if (-not $certificate.HasPrivateKey) {
        return [pscustomobject]@{ IsAvailable = $false; Reason = "Certificate $normalized does not have an accessible private key."; Certificate = $certificate }
    }
    if ($certificate.NotAfter -le [datetime]::Now) {
        return [pscustomobject]@{ IsAvailable = $false; Reason = "Certificate $normalized is expired."; Certificate = $certificate }
    }

    return [pscustomobject]@{ IsAvailable = $true; Reason = ''; Certificate = $certificate }
}

Export-ModuleMember -Function @(
    'Protect-WinTunerClientSecretForCurrentUser'
    'Unprotect-WinTunerClientSecretForCurrentUser'
    'ConvertTo-WinTunerCertificateThumbprint'
    'Test-WinTunerAuthenticationConfiguration'
    'New-WinTunerModuleConnectionParameters'
    'Test-WinTunerCertificateAvailable'
    'Protect-WinTunerAuthenticationErrorMessage'
    'Get-WinTunerAppOnlyTokenCachePath'
    'Clear-WinTunerAppOnlyTokenCache'
)
