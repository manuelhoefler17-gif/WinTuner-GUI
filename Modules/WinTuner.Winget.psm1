Set-StrictMode -Version Latest

$script:wingetVersionCache = @{}
$script:diskCache          = @{}
$script:diskCacheLoaded    = $false
$script:versionCachePath   = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'WinTuner_VersionCache.json'

function Get-VersionDiskCache {
    if (-not $script:versionCachePath) {
        return @{}
    }

    try {
        if (Test-Path $script:versionCachePath) {
            $raw = Get-Content $script:versionCachePath -Raw -Encoding utf8 -ErrorAction Stop
            $parsed = $raw | ConvertFrom-Json -ErrorAction Stop

            $ht = @{}

            foreach ($prop in $parsed.PSObject.Properties) {
                $ht[$prop.Name] = @{
                    versions  = @($prop.Value.versions)
                    timestamp = [datetime]::Parse(
                        $prop.Value.timestamp,
                        [System.Globalization.CultureInfo]::InvariantCulture,
                        [System.Globalization.DateTimeStyles]::RoundtripKind
                    )
                }
            }

            return $ht
        }
    }
    catch {
        Write-Verbose "Could not read WinGet version cache: $($_.Exception.Message)"
    }

    return @{}
}

function Save-VersionDiskCache {
    param(
        [Parameter(Mandatory)]
        [hashtable]$Cache
    )

    if (-not $script:versionCachePath) {
        return
    }

    try {
        $obj = @{}

        foreach ($key in $Cache.Keys) {
            $obj[$key] = @{
                versions  = $Cache[$key].versions
                timestamp = $Cache[$key].timestamp.ToString('o')
            }
        }

        $obj |
            ConvertTo-Json -Depth 4 |
            Set-Content -Path $script:versionCachePath -Encoding utf8 -ErrorAction Stop
    }
    catch {
        Write-Verbose "Could not save WinGet version cache: $($_.Exception.Message)"
    }
}

function Get-WingetVersions {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$PackageId
    )

    # RAM cache
    if ($script:wingetVersionCache.ContainsKey($PackageId)) {
        return $script:wingetVersionCache[$PackageId]
    }

    # Disk cache
    if (-not $script:diskCacheLoaded) {
        $script:diskCache = Get-VersionDiskCache
        $script:diskCacheLoaded = $true
    }

    if ($script:diskCache.ContainsKey($PackageId)) {
        $entry = $script:diskCache[$PackageId]
        $ageHours = ([datetime]::UtcNow - $entry.timestamp.ToUniversalTime()).TotalHours

        if (
            $ageHours -lt 6 -and
            $entry.versions -and
            $entry.versions.Count -gt 0
        ) {
            $script:wingetVersionCache[$PackageId] = @($entry.versions)
            return @($entry.versions)
        }
    }

    # Query winget
    try {
        $output = & winget show --id $PackageId --versions 2>$null
    }
    catch {
        return @()
    }

    if (-not $output) {
        return @()
    }

    $candidates = @()

    foreach ($line in @($output)) {
        $text = ($line -replace '^[\s\-•]+', '').Trim()

        if (-not $text) {
            continue
        }

        if ($text -match '^(\d+)(\.[0-9A-Za-z]+)*([\-+._][0-9A-Za-z]+)*$') {
            $candidates += $text
        }
    }

    $unique = @($candidates | Select-Object -Unique)

    $parsed = foreach ($version in $unique) {
        $numeric = $false
        $versionObject = $null

        try {
            $versionObject = [version]$version
            $numeric = $true
        }
        catch {
        }

        [pscustomobject]@{
            Text    = $version
            Parsed  = $versionObject
            Numeric = $numeric
        }
    }

    $result = @()

    if ($parsed | Where-Object Numeric) {
        $result = @(
            $parsed |
                Where-Object Numeric |
                Sort-Object Parsed -Descending |
                Select-Object -ExpandProperty Text
        )
    }
    else {
        $result = @(
            $parsed |
                Sort-Object Text -Descending |
                Select-Object -ExpandProperty Text
        )
    }

    $script:wingetVersionCache[$PackageId] = $result

    $script:diskCache[$PackageId] = @{
        versions  = $result
        timestamp = [datetime]::UtcNow
    }

    Save-VersionDiskCache -Cache $script:diskCache

    return $result
}

function Get-PreviousWingetVersion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$PackageId,

        [string]$LatestVersion
    )

    $allVersions = @(Get-WingetVersions -PackageId $PackageId)

    if (-not $allVersions -or $allVersions.Count -eq 0) {
        return $null
    }

    $candidates = @(
        $allVersions |
            Where-Object { $_ -ne $LatestVersion }
    )

    if ($candidates.Count -gt 0) {
        return $candidates[0]
    }

    return $null
}

function Clear-WingetVersionCache {
    [CmdletBinding()]
    param()

    $script:wingetVersionCache = @{}
    $script:diskCache          = @{}
    $script:diskCacheLoaded    = $false

    Remove-Item $script:versionCachePath -Force -ErrorAction SilentlyContinue
}

Export-ModuleMember -Function @(
    'Get-WingetVersions'
    'Get-PreviousWingetVersion'
    'Clear-WingetVersionCache'
)
