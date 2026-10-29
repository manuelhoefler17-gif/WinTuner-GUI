Set-StrictMode -Version Latest
Import-Module (Join-Path $PSScriptRoot 'WinTuner.Authentication.psm1') -ErrorAction Stop

function New-WinTunerTenantProfile {
    [CmdletBinding()]
    param(
        [guid]$Id = [guid]::NewGuid(),
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][guid]$TenantId,
        [Parameter(Mandatory)][ValidateSet('Interactive','ClientSecret','Certificate')][string]$AuthenticationMode,
        [string]$UserPrincipalName = '', [string]$ClientId = '',
        [string]$CertificateThumbprint = '', [string]$EntraClientSecretProtected = '',
        [string]$DefaultPackagePath = 'C:\Temp'
    )
    $Name = $Name.Trim()
    if ($Id -eq [guid]::Empty -or $TenantId -eq [guid]::Empty) { throw 'A profile requires a non-empty profile and tenant ID.' }
    if ($Name.Length -eq 0 -or $Name.Length -gt 80 -or $Name -match '[\x00-\x1f]') { throw 'Enter a profile name of 1 to 80 characters without control characters.' }
    $validation = Test-WinTunerAuthenticationConfiguration -Mode $AuthenticationMode -UserPrincipalName $UserPrincipalName -TenantId $TenantId.ToString('D') -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
    if (-not $validation.IsValid) { throw $validation.Reason }
    if ([string]::IsNullOrWhiteSpace($DefaultPackagePath) -or -not [IO.Path]::IsPathFullyQualified($DefaultPackagePath)) { throw 'The package folder must be an absolute path.' }
    if ($AuthenticationMode -ne 'ClientSecret') { $EntraClientSecretProtected = '' }
    if ($EntraClientSecretProtected) {
        # DPAPI blobs, never a plaintext secret. Decryption belongs exclusively to login.
        try { $bytes = [Convert]::FromBase64String($EntraClientSecretProtected) } catch { throw 'Invalid protected client secret.' }
        $prefix = [byte[]](1,0,0,0,208,140,157,223,1,21,209,17,140,122,0,192,79,194,151,235)
        if ($bytes.Length -le $prefix.Length -or [Convert]::ToHexString($bytes[0..19]) -ne [Convert]::ToHexString($prefix)) { throw 'Only a Windows DPAPI-protected client secret can be saved in a profile.' }
    }
    [pscustomobject]@{
        Id = $Id.ToString('D'); Name = $Name; TenantId = $TenantId.ToString('D')
        AuthenticationMode = $AuthenticationMode
        UserPrincipalName = $(if ($AuthenticationMode -eq 'Interactive') { $UserPrincipalName.Trim() } else { '' })
        ClientId = $(if ($AuthenticationMode -ne 'Interactive') { $validation.ClientId } else { '' })
        CertificateThumbprint = $(if ($AuthenticationMode -eq 'Certificate') { $validation.CertificateThumbprint } else { '' })
        EntraClientSecretProtected = $EntraClientSecretProtected; DefaultPackagePath = $DefaultPackagePath
    }
}

function Get-WinTunerTenantProfilesPath {
    param([string]$Path)
    if ($Path) { return [IO.Path]::GetFullPath($Path) }
    Join-Path ([Environment]::GetFolderPath('ApplicationData')) 'WinTunerGUI\TenantProfiles.json'
}

function Get-WinTunerTenantProfiles {
    [CmdletBinding()]
    param([string]$Path)
    $Path = Get-WinTunerTenantProfilesPath -Path $Path
    if (-not (Test-Path -LiteralPath $Path)) { return @() }
    try {
        $document = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        if ($document.SchemaVersion -ne 1 -or $null -eq $document.Profiles) { throw 'Invalid schema.' }
        $ids = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        $names = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($profile in @($document.Profiles)) {
            $parameters = @{}
            foreach ($key in @('Id','Name','TenantId','AuthenticationMode','UserPrincipalName','ClientId','CertificateThumbprint','EntraClientSecretProtected','DefaultPackagePath')) {
                $parameters[$key] = [string]$profile.$key
            }
            $validated = New-WinTunerTenantProfile @parameters
            if (-not $ids.Add($validated.Id) -or -not $names.Add($validated.Name)) { throw 'Duplicate profile.' }
            $validated
        }
    } catch { throw 'The tenant profile file could not be read or validated. Repair it or explicitly choose manual configuration; no profile fallback was applied.' }
}

function Set-WinTunerTenantProfile {
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Profile, [switch]$Remove, [string]$Path)
    $parameters = @{}
    foreach ($key in @('Id','Name','TenantId','AuthenticationMode','UserPrincipalName','ClientId','CertificateThumbprint','EntraClientSecretProtected','DefaultPackagePath')) { $parameters[$key] = [string]$Profile.$key }
    $validated = New-WinTunerTenantProfile @parameters
    $Path = Get-WinTunerTenantProfilesPath -Path $Path
    $key = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes($Path.ToUpperInvariant())))
    $mutex = [Threading.Mutex]::new($false, "Local\WinTuner-TenantProfiles-$key")
    $acquired = $false
    $temporary = "$Path.$([guid]::NewGuid().ToString('N')).tmp"
    try {
        try { $acquired = $mutex.WaitOne(5000) } catch [Threading.AbandonedMutexException] { $acquired = $true }
        if (-not $acquired) { throw 'Another instance is saving tenant profiles. Try again.' }
        $profiles = @(Get-WinTunerTenantProfiles -Path $Path)
        $existing = @($profiles | Where-Object Id -EQ $validated.Id)
        if ($existing.Count -and $existing[0].TenantId -ne $validated.TenantId) { throw 'A saved profile cannot be moved to another tenant. Create a new profile.' }
        $remaining = @($profiles | Where-Object Id -NE $validated.Id)
        if (-not $Remove) {
            if (@($remaining | Where-Object Name -EQ $validated.Name).Count) { throw 'A profile with this name already exists.' }
            $remaining += $validated
        }
        [void][IO.Directory]::CreateDirectory((Split-Path -Parent $Path))
        $document = [ordered]@{ SchemaVersion = 1; Profiles = @($remaining) }
        [IO.File]::WriteAllText($temporary, ($document | ConvertTo-Json -Depth 6), [Text.UTF8Encoding]::new($false))
        [IO.File]::Move($temporary, $Path, $true)
    } finally {
        if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force }
        if ($acquired) { $mutex.ReleaseMutex() }
        $mutex.Dispose()
    }
}

function Assert-WinTunerTenantProfileContext {
    param([string]$ExpectedTenantId, [string]$ActualTenantId)
    if (-not $ExpectedTenantId) { return }
    $expected = [guid]::Parse($ExpectedTenantId)
    $actual = [guid]::Empty
    if ($expected -eq [guid]::Empty -or -not [guid]::TryParse($ActualTenantId, [ref]$actual) -or $expected -ne $actual) {
        throw "The signed-in tenant does not match the selected profile (expected $expected). Sign out and choose the correct tenant."
    }
}

function Get-WinTunerTenantProfileSettings {
    param([Parameter(Mandatory)][hashtable]$Settings, [AllowNull()][object]$Profile, [bool]$Connected, [bool]$IsBusy)
    if ($Connected -or $IsBusy) { throw 'Sign out and wait for all operations to finish before switching profiles.' }
    $result = $Settings.Clone()
    # Clear the entire previous identity, including its protected secret, on every switch.
    $result.ActiveTenantProfileId = ''
    $result.AuthenticationMode = 'Interactive'
    $result.EntraTenantId = ''; $result.EntraClientId = ''; $result.EntraCertificateThumbprint = ''; $result.EntraClientSecretProtected = ''; $result.LastUser = ''
    if ($null -ne $Profile) {
        $parameters = @{}
        foreach ($key in @('Id','Name','TenantId','AuthenticationMode','UserPrincipalName','ClientId','CertificateThumbprint','EntraClientSecretProtected','DefaultPackagePath')) { $parameters[$key] = [string]$Profile.$key }
        $validated = New-WinTunerTenantProfile @parameters
        $result.ActiveTenantProfileId = $validated.Id
        $result.AuthenticationMode = $validated.AuthenticationMode
        $result.EntraTenantId = $validated.TenantId
        $result.EntraClientId = $validated.ClientId
        $result.EntraCertificateThumbprint = $validated.CertificateThumbprint
        $result.EntraClientSecretProtected = $validated.EntraClientSecretProtected
        $result.LastUser = $validated.UserPrincipalName
        $result.DefaultPackagePath = $validated.DefaultPackagePath
    }
    return $result
}

Export-ModuleMember -Function New-WinTunerTenantProfile, Get-WinTunerTenantProfiles, Set-WinTunerTenantProfile, Assert-WinTunerTenantProfileContext, Get-WinTunerTenantProfileSettings
