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


$script:discoverySearchCache       = @{}
$script:discoverySearchCacheLoaded = $false
$script:discoverySearchCacheDirty  = $false
$script:discoveryCachePath         = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'WinTuner_DiscoveryCache.json'

function Get-DiscoverySearchDiskCache {
    $cache = @{}

    try {
        if (-not (Test-Path $script:discoveryCachePath)) {
            return $cache
        }

        $raw = Get-Content `
            -Path $script:discoveryCachePath `
            -Raw `
            -Encoding utf8 `
            -ErrorAction Stop

        if ([string]::IsNullOrWhiteSpace($raw)) {
            return $cache
        }

        $parsed = $raw | ConvertFrom-Json -ErrorAction Stop

        foreach ($prop in $parsed.PSObject.Properties) {
            $results = @(
                foreach ($item in @($prop.Value.results)) {
                    if ($null -eq $item) {
                        continue
                    }

                    [pscustomobject]@{
                        Name      = [string]$item.Name
                        PackageID = [string]$item.PackageID
                    }
                }
            )

            $cache[$prop.Name] = @{
                timestamp = [datetime]::Parse(
                    [string]$prop.Value.timestamp,
                    [System.Globalization.CultureInfo]::InvariantCulture,
                    [System.Globalization.DateTimeStyles]::RoundtripKind
                )
                results = $results
            }
        }
    }
    catch {
        Write-Verbose "Could not read WinGet discovery cache: $($_.Exception.Message)"
        return @{}
    }

    return $cache
}

function Save-WinTunerDiscoveryCache {
    [CmdletBinding()]
    param()

    if (-not $script:discoverySearchCacheLoaded) {
        return
    }

    if (-not $script:discoverySearchCacheDirty) {
        return
    }

    try {
        $obj = [ordered]@{}

        foreach ($key in @($script:discoverySearchCache.Keys | Sort-Object)) {
            $entry = $script:discoverySearchCache[$key]

            $obj[$key] = @{
                timestamp = $entry.timestamp.ToString('o')
                results   = @(
                    foreach ($item in @($entry.results)) {
                        [pscustomobject]@{
                            Name      = [string]$item.Name
                            PackageID = [string]$item.PackageID
                        }
                    }
                )
            }
        }

        $obj |
            ConvertTo-Json -Depth 6 |
            Set-Content `
                -Path $script:discoveryCachePath `
                -Encoding utf8 `
                -ErrorAction Stop

        $script:discoverySearchCacheDirty = $false
    }
    catch {
        Write-Verbose "Could not save WinGet discovery cache: $($_.Exception.Message)"
    }
}

function Clear-WinTunerDiscoveryCache {
    [CmdletBinding()]
    param()

    $script:discoverySearchCache       = @{}
    $script:discoverySearchCacheLoaded = $false
    $script:discoverySearchCacheDirty  = $false

    Remove-Item `
        -Path $script:discoveryCachePath `
        -Force `
        -ErrorAction SilentlyContinue
}

function Search-WinTunerDiscoveryPackageCached {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$SearchQuery,

        [ValidateRange(1,120)]
        [int]$TimeoutSeconds = 12,

        [ValidateRange(1,168)]
        [int]$CacheTtlHours = 24,

        [scriptblock]$OnWait
    )

    if (-not $script:discoverySearchCacheLoaded) {
        $script:discoverySearchCache = Get-DiscoverySearchDiskCache
        $script:discoverySearchCacheLoaded = $true
    }

    $cacheKey = $SearchQuery.Trim().ToLowerInvariant()

    if ($script:discoverySearchCache.ContainsKey($cacheKey)) {
        $entry = $script:discoverySearchCache[$cacheKey]
        $ageHours = (
            [datetime]::UtcNow -
            $entry.timestamp.ToUniversalTime()
        ).TotalHours

        if ($ageHours -lt $CacheTtlHours) {
            return @($entry.results)
        }

        [void]$script:discoverySearchCache.Remove($cacheKey)
        $script:discoverySearchCacheDirty = $true
    }

    $rawResults = @(
        Search-WinTunerPackageWithTimeout `
            -SearchQuery $SearchQuery `
            -TimeoutSeconds $TimeoutSeconds `
            -OnWait $OnWait
    )

    $results = @(
        $rawResults |
            Where-Object {
                -not [string]::IsNullOrWhiteSpace([string]$_.Name) -and
                -not [string]::IsNullOrWhiteSpace([string]$_.PackageID)
            } |
            ForEach-Object {
                [pscustomobject]@{
                    Name      = [string]$_.Name
                    PackageID = [string]$_.PackageID
                }
            } |
            Sort-Object PackageID -Unique
    )

    $script:discoverySearchCache[$cacheKey] = @{
        timestamp = [datetime]::UtcNow
        results   = $results
    }

    $script:discoverySearchCacheDirty = $true

    return $results
}

function Search-WinTunerPackageWithTimeout {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$SearchQuery,

        [ValidateRange(1,120)]
        [int]$TimeoutSeconds = 12,

        [scriptblock]$OnWait
    )

    $modulePath = (Get-Module WinTuner -ListAvailable |
        Sort-Object Version -Descending |
        Select-Object -First 1).Path

    if (-not $modulePath) {
        throw "WinTuner module not found."
    }

    $job = Start-ThreadJob -ScriptBlock {
        param($ModulePath, $Query)

        Import-Module $ModulePath -Force -ErrorAction Stop

        @(Search-WtWinGetPackage `
            -SearchQuery $Query `
            -ErrorAction SilentlyContinue `
            2>$null 3>$null 4>$null 5>$null 6>$null)

    } -ArgumentList $modulePath, $SearchQuery

    try {
        $deadline = [datetime]::UtcNow.AddSeconds($TimeoutSeconds)

        while ($job.State -eq 'NotStarted' -or $job.State -eq 'Running') {

            if ([datetime]::UtcNow -ge $deadline) {
                Stop-Job $job -ErrorAction SilentlyContinue
                throw "WinGet search timed out after $TimeoutSeconds seconds."
            }

            if ($OnWait) {
                & $OnWait
            }

            Start-Sleep -Milliseconds 100
        }

        if ($job.State -eq 'Failed') {
            $reason = $job.ChildJobs[0].JobStateInfo.Reason
            if ($reason) {
                throw $reason
            }

            throw "WinGet search failed."
        }

        return @(Receive-Job $job -ErrorAction Stop)
    }
    finally {
        Remove-Job $job -Force -ErrorAction SilentlyContinue
    }
}

Export-ModuleMember -Function Search-WinTunerPackageWithTimeout

function Search-WinTunerDiscoveryPackagesBatchCached {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$SearchQueries,

        [ValidateRange(1,100)]
        [int]$BatchSize = 25,

        [ValidateRange(1,720)]
        [int]$QueryTimeoutSeconds = 12,

        [ValidateRange(1,720)]
        [int]$CacheTtlHours = 24,

        [scriptblock]$OnWait,

        [scriptblock]$ShouldCancel
    )

    if (-not $script:discoverySearchCacheLoaded) {
        $script:discoverySearchCache = Get-DiscoverySearchDiskCache
        $script:discoverySearchCacheLoaded = $true
    }

    $repoRoot = Split-Path $PSScriptRoot -Parent
    $workerPath = Join-Path $repoRoot 'Workers\WinTuner.DiscoveryWorker.ps1'

    if (-not (Test-Path $workerPath)) {
        throw "Discovery worker not found: $workerPath"
    }

    $pwshPath = (Get-Command pwsh -ErrorAction Stop).Source

    if ([string]::IsNullOrWhiteSpace($pwshPath)) {
        throw 'PowerShell 7 executable not found.'
    }

    $uniqueQueries = [System.Collections.Generic.List[string]]::new()
    $seenQueries = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )

    foreach ($queryItem in $SearchQueries) {
        $query = [string]$queryItem

        if ([string]::IsNullOrWhiteSpace($query)) {
            continue
        }

        $query = $query.Trim()

        if ($seenQueries.Add($query)) {
            $uniqueQueries.Add($query)
        }
    }

    $resultList = [System.Collections.Generic.List[object]]::new()
    $misses = [System.Collections.Generic.List[string]]::new()

    foreach ($query in $uniqueQueries) {
        $cacheKey = $query.ToLowerInvariant()

        if ($script:discoverySearchCache.ContainsKey($cacheKey)) {
            $entry = $script:discoverySearchCache[$cacheKey]

            try {
                $timestamp = [datetime]$entry.timestamp
                $age = [datetime]::UtcNow - $timestamp.ToUniversalTime()

                if ($age.TotalHours -lt $CacheTtlHours) {
                    $cachedResults = @(
                        foreach ($item in @($entry.results)) {
                            [pscustomobject]@{
                                Name      = [string]$item.Name
                                PackageID = [string]$item.PackageID
                            }
                        }
                    )

                    $resultList.Add(
                        [pscustomobject]@{
                            Query     = $query
                            Success   = $true
                            FromCache = $true
                            Results   = $cachedResults
                            Error     = $null
                        }
                    )

                    continue
                }
            }
            catch {
            }

            [void]$script:discoverySearchCache.Remove($cacheKey)
            $script:discoverySearchCacheDirty = $true
        }

        $misses.Add($query)
    }

    $workerCount = 0
    $workerQueryCount = 0
    $canceled = $false

    for (
        $offset = 0;
        $offset -lt $misses.Count;
        $offset += $BatchSize
    ) {
        if ($ShouldCancel -and (& $ShouldCancel)) {
            $canceled = $true
            break
        }

        $lastIndex = [math]::Min(
            $offset + $BatchSize - 1,
            $misses.Count - 1
        )

        $batch = @(
            $misses[$offset..$lastIndex]
        )

        $workerCount++
        $workerQueryCount += $batch.Count

        $token = [guid]::NewGuid().ToString('N')
        $inputPath = Join-Path $env:TEMP "WinTuner_Discovery_Input_$token.json"
        $outputPath = Join-Path $env:TEMP "WinTuner_Discovery_Output_$token.json"

        $process = $null

        try {
            $batch |
                ConvertTo-Json |
                Set-Content `
                    -Path $inputPath `
                    -Encoding utf8 `
                    -NoNewline

            Remove-Item `
                -Path $outputPath `
                -Force `
                -ErrorAction SilentlyContinue

            $psi = [System.Diagnostics.ProcessStartInfo]::new()
            $psi.FileName = $pwshPath
            $psi.UseShellExecute = $false
            $psi.CreateNoWindow = $true
            $psi.RedirectStandardOutput = $true
            $psi.RedirectStandardError = $true

            [void]$psi.ArgumentList.Add('-NoProfile')
            [void]$psi.ArgumentList.Add('-ExecutionPolicy')
            [void]$psi.ArgumentList.Add('Bypass')
            [void]$psi.ArgumentList.Add('-File')
            [void]$psi.ArgumentList.Add($workerPath)
            [void]$psi.ArgumentList.Add('-InputPath')
            [void]$psi.ArgumentList.Add($inputPath)
            [void]$psi.ArgumentList.Add('-OutputPath')
            [void]$psi.ArgumentList.Add($outputPath)

            $process = [System.Diagnostics.Process]::new()
            $process.StartInfo = $psi

            if (-not $process.Start()) {
                throw "Discovery worker $workerCount could not be started."
            }

            $batchDeadline = [datetime]::UtcNow.AddSeconds(
                ($batch.Count * $QueryTimeoutSeconds) + 30
            )

            while (-not $process.HasExited) {
                if ($ShouldCancel -and (& $ShouldCancel)) {
                    try {
                        $process.Kill($true)
                    }
                    catch {
                    }

                    $canceled = $true
                    break
                }

                if ([datetime]::UtcNow -ge $batchDeadline) {
                    try {
                        $process.Kill($true)
                    }
                    catch {
                    }

                    throw "Discovery worker $workerCount exceeded batch timeout."
                }

                if ($OnWait) {
                    & $OnWait
                }

                Start-Sleep -Milliseconds 100
            }

            if ($canceled) {
                break
            }

            $stdout = $process.StandardOutput.ReadToEnd()
            $stderr = $process.StandardError.ReadToEnd()

            if ($process.ExitCode -ne 0) {
                throw "Discovery worker $workerCount failed with exit code $($process.ExitCode). STDERR: $stderr STDOUT: $stdout"
            }

            if (-not (Test-Path $outputPath)) {
                throw "Discovery worker $workerCount produced no output file."
            }

            $workerResults = @(
                Get-Content `
                    -Path $outputPath `
                    -Raw `
                    -Encoding utf8 |
                    ConvertFrom-Json
            )

            $returnedQueries = [System.Collections.Generic.HashSet[string]]::new(
                [System.StringComparer]::OrdinalIgnoreCase
            )

            foreach ($workerResult in $workerResults) {
                $query = [string]$workerResult.Query

                if ([string]::IsNullOrWhiteSpace($query)) {
                    continue
                }

                [void]$returnedQueries.Add($query)

                if (-not [bool]$workerResult.Success) {
                    $resultList.Add(
                        [pscustomobject]@{
                            Query     = $query
                            Success   = $false
                            FromCache = $false
                            Results   = @()
                            Error     = [string]$workerResult.Error
                        }
                    )

                    continue
                }

                $cleanResults = @(
                    foreach ($item in @($workerResult.Results)) {
                        if ($null -eq $item) {
                            continue
                        }

                        $packageId = [string]$item.PackageID

                        if ([string]::IsNullOrWhiteSpace($packageId)) {
                            continue
                        }

                        [pscustomobject]@{
                            Name      = [string]$item.Name
                            PackageID = $packageId
                        }
                    }
                )

                $cacheKey = $query.Trim().ToLowerInvariant()

                $script:discoverySearchCache[$cacheKey] = @{
                    timestamp = [datetime]::UtcNow
                    results   = $cleanResults
                }

                $script:discoverySearchCacheDirty = $true

                $resultList.Add(
                    [pscustomobject]@{
                        Query     = $query
                        Success   = $true
                        FromCache = $false
                        Results   = $cleanResults
                        Error     = $null
                    }
                )
            }

            foreach ($query in $batch) {
                if (-not $returnedQueries.Contains($query)) {
                    $resultList.Add(
                        [pscustomobject]@{
                            Query     = $query
                            Success   = $false
                            FromCache = $false
                            Results   = @()
                            Error     = 'Worker returned no result for this query.'
                        }
                    )
                }
            }

            Save-WinTunerDiscoveryCache
        }
        finally {
            if ($process -and -not $process.HasExited) {
                try {
                    $process.Kill($true)
                }
                catch {
                }
            }

            if ($process) {
                $process.Dispose()
            }

            Remove-Item `
                -Path $inputPath, $outputPath `
                -Force `
                -ErrorAction SilentlyContinue
        }
    }

    [pscustomobject]@{
        Results       = @($resultList)
        Canceled      = $canceled
        TotalQueries  = $uniqueQueries.Count
        CacheHits     = @(
            $resultList |
                Where-Object FromCache
        ).Count
        WorkerQueries = $workerQueryCount
        WorkerCount   = $workerCount
    }
}
Export-ModuleMember -Function @(
    'Get-WingetVersions'
    'Get-PreviousWingetVersion'
    'Clear-WingetVersionCache'
    'Search-WinTunerDiscoveryPackageCached',
    'Search-WinTunerDiscoveryPackagesBatchCached'
    'Save-WinTunerDiscoveryCache'
    'Clear-WinTunerDiscoveryCache'
)
