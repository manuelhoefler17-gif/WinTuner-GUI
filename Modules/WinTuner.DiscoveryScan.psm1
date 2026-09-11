Set-StrictMode -Version Latest

function Get-WinTunerDiscoveryValue {
    param([AllowNull()][object]$InputObject, [Parameter(Mandatory)][string]$Name)
    if ($null -eq $InputObject) { return $null }
    $property = $InputObject.PSObject.Properties[$Name]
    if ($property) { return $property.Value }
    return $null
}

function Test-WinTunerDiscoveryCandidate {
    param([AllowNull()][string]$DisplayName)
    if ([string]::IsNullOrWhiteSpace($DisplayName)) { return $false }
    $name = $DisplayName.Trim()
    if ($name.Length -lt 3) { return $false }
    if ($name -match '^[a-z0-9]+(\.[a-z0-9_]+){2,}$') { return $false }
    if ($name -match '(?i)^com\.') { return $false }
    if ($name -match '(?i)\b(apn|provisioner|sim toolkit|sim card|carrier services|system ui|one ui home|setup wizard)\b') { return $false }
    return $true
}

function Get-WinTunerDiscoverySimilarity {
    param([AllowNull()][string]$First, [AllowNull()][string]$Second)
    if ([string]::IsNullOrWhiteSpace($First) -or [string]::IsNullOrWhiteSpace($Second)) { return 0 }
    $cleanFirst = $First.ToLowerInvariant() -replace '[^\w\s]', ' '
    $cleanSecond = $Second.ToLowerInvariant() -replace '[^\w\s]', ' '
    $firstWords = @($cleanFirst -split '\s+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $secondWords = @($cleanSecond -split '\s+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($firstWords.Count -eq 0 -or $secondWords.Count -eq 0) { return 0 }
    $matches = 0
    foreach ($word in $firstWords) {
        if ($secondWords -contains $word) { $matches++ }
    }
    [math]::Round(($matches / [math]::Min($firstWords.Count, $secondWords.Count)) * 100)
}

function New-WinTunerDiscoveryScanResult {
    param(
        [bool]$Canceled,
        [AllowNull()][string]$ErrorMessage,
        [object[]]$Apps,
        [int]$DetectedCount = 0,
        [int]$FilteredCount = 0,
        [int]$NormalizedCount = 0,
        [int]$MatchedRawCount = 0,
        [int]$SkippedNonCandidateCount = 0,
        [bool]$GraphFromCache = $false,
        [int]$GraphPageCount = 0,
        [bool]$GraphLimitReached = $false,
        [int]$TotalQueries = 0,
        [int]$CacheHits = 0,
        [int]$WorkerQueries = 0,
        [int]$WorkerCount = 0
    )
    [pscustomobject]@{
        Canceled = $Canceled
        ErrorMessage = $ErrorMessage
        Apps = @($Apps)
        DetectedCount = $DetectedCount
        FilteredCount = $FilteredCount
        NormalizedCount = $NormalizedCount
        MatchedRawCount = $MatchedRawCount
        UniquePackageCount = @($Apps).Count
        SkippedNonCandidateCount = $SkippedNonCandidateCount
        GraphFromCache = $GraphFromCache
        GraphPageCount = $GraphPageCount
        GraphLimitReached = $GraphLimitReached
        TotalQueries = $TotalQueries
        CacheHits = $CacheHits
        WorkerQueries = $WorkerQueries
        WorkerCount = $WorkerCount
    }
}

function Invoke-WinTunerDiscoveryScan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][scriptblock]$ConnectGraph,
        [Parameter(Mandatory)][scriptblock]$GetExistingApps,
        [Parameter(Mandatory)][scriptblock]$ResolvePackageId,
        [Parameter(Mandatory)][scriptblock]$GetDetectedApps,
        [Parameter(Mandatory)][scriptblock]$SearchPackages,
        [scriptblock]$ShouldCancel = { $false },
        [scriptblock]$ReportProgress = { param($ProgressInfo) },
        [bool]$SkipLowValueCandidates = $false
    )

    $detectedCount = 0
    $filteredCount = 0
    $normalizedCount = 0
    $matchedRawCount = 0
    $skippedCount = 0
    $graphFromCache = $false
    $graphPageCount = 0
    $graphLimitReached = $false
    $totalQueries = 0
    $cacheHits = 0
    $workerQueries = 0
    $workerCount = 0

    try {
        if (& $ShouldCancel) { return New-WinTunerDiscoveryScanResult -Canceled $true -Apps @() }

        try { & $ReportProgress ([pscustomobject]@{ Stage = 'Connecting'; Processed = 0; Total = 1; AppName = '' }) } catch {}
        $null = & $ConnectGraph
        if (& $ShouldCancel) { return New-WinTunerDiscoveryScanResult -Canceled $true -Apps @() }

        try { & $ReportProgress ([pscustomobject]@{ Stage = 'LoadingExisting'; Processed = 0; Total = 1; AppName = '' }) } catch {}
        $existingPackageIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($existingApp in @(& $GetExistingApps)) {
            if (& $ShouldCancel) { return New-WinTunerDiscoveryScanResult -Canceled $true -Apps @() }
            $packageId = [string](& $ResolvePackageId $existingApp)
            if (-not [string]::IsNullOrWhiteSpace($packageId)) { $null = $existingPackageIds.Add($packageId) }
        }

        try { & $ReportProgress ([pscustomobject]@{ Stage = 'FetchingDetected'; Processed = 0; Total = 1; AppName = '' }) } catch {}
        $detectedResult = & $GetDetectedApps
        $detectedApps = @(Get-WinTunerDiscoveryValue -InputObject $detectedResult -Name 'Apps')
        $detectedCount = $detectedApps.Count
        $graphFromCache = [bool](Get-WinTunerDiscoveryValue -InputObject $detectedResult -Name 'FromCache')
        $graphPageCount = [int](Get-WinTunerDiscoveryValue -InputObject $detectedResult -Name 'PageCount')
        $graphLimitReached = [bool](Get-WinTunerDiscoveryValue -InputObject $detectedResult -Name 'LimitReached')

        if (& $ShouldCancel) {
            return New-WinTunerDiscoveryScanResult -Canceled $true -Apps @() -DetectedCount $detectedCount -GraphFromCache $graphFromCache -GraphPageCount $graphPageCount -GraphLimitReached $graphLimitReached
        }

        $filteredApps = @($detectedApps | Where-Object {
            [string](Get-WinTunerDiscoveryValue -InputObject $_ -Name 'publisher') -notmatch '(?i)Intel|HP|Dell|Lenovo|AMD|NVIDIA|Realtek|Synaptics|VMware'
        })
        $filteredCount = $filteredApps.Count
        $normalizedApps = [System.Collections.Generic.List[object]]::new()
        $uniqueQueries = [System.Collections.Generic.List[string]]::new()
        $querySet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

        foreach ($app in $filteredApps) {
            if (& $ShouldCancel) {
                return New-WinTunerDiscoveryScanResult -Canceled $true -Apps @() -DetectedCount $detectedCount -FilteredCount $filteredCount -NormalizedCount $normalizedApps.Count -SkippedNonCandidateCount $skippedCount -GraphFromCache $graphFromCache -GraphPageCount $graphPageCount -GraphLimitReached $graphLimitReached
            }
            $searchName = [string](Get-WinTunerDiscoveryValue -InputObject $app -Name 'displayName')
            $searchName = ($searchName -replace '\s*\([^)]*\)', '' -replace '\s+[\d\.]+', '').Trim()
            if ([string]::IsNullOrWhiteSpace($searchName)) { continue }
            if (-not (Test-WinTunerDiscoveryCandidate -DisplayName $searchName) -and $SkipLowValueCandidates) {
                $skippedCount++
                continue
            }
            $normalizedApps.Add([pscustomobject]@{ App = $app; SearchName = $searchName })
            if ($querySet.Add($searchName)) { $uniqueQueries.Add($searchName) }
        }
        $normalizedCount = $normalizedApps.Count
        if ($uniqueQueries.Count -eq 0) {
            return New-WinTunerDiscoveryScanResult -Canceled $false -Apps @() -DetectedCount $detectedCount -FilteredCount $filteredCount -NormalizedCount $normalizedCount -SkippedNonCandidateCount $skippedCount -GraphFromCache $graphFromCache -GraphPageCount $graphPageCount -GraphLimitReached $graphLimitReached
        }

        try { & $ReportProgress ([pscustomobject]@{ Stage = 'Searching'; Processed = 0; Total = $uniqueQueries.Count; AppName = '' }) } catch {}
        $batchResult = & $SearchPackages @($uniqueQueries) $ShouldCancel
        $totalQueries = [int](Get-WinTunerDiscoveryValue -InputObject $batchResult -Name 'TotalQueries')
        $cacheHits = [int](Get-WinTunerDiscoveryValue -InputObject $batchResult -Name 'CacheHits')
        $workerQueries = [int](Get-WinTunerDiscoveryValue -InputObject $batchResult -Name 'WorkerQueries')
        $workerCount = [int](Get-WinTunerDiscoveryValue -InputObject $batchResult -Name 'WorkerCount')
        if ([bool](Get-WinTunerDiscoveryValue -InputObject $batchResult -Name 'Canceled') -or (& $ShouldCancel)) {
            return New-WinTunerDiscoveryScanResult -Canceled $true -Apps @() -DetectedCount $detectedCount -FilteredCount $filteredCount -NormalizedCount $normalizedCount -SkippedNonCandidateCount $skippedCount -GraphFromCache $graphFromCache -GraphPageCount $graphPageCount -GraphLimitReached $graphLimitReached -TotalQueries $totalQueries -CacheHits $cacheHits -WorkerQueries $workerQueries -WorkerCount $workerCount
        }

        $searchResultCache = @{}
        foreach ($searchResult in @(Get-WinTunerDiscoveryValue -InputObject $batchResult -Name 'Results')) {
            $query = [string](Get-WinTunerDiscoveryValue -InputObject $searchResult -Name 'Query')
            if ([string]::IsNullOrWhiteSpace($query)) { continue }
            $success = [bool](Get-WinTunerDiscoveryValue -InputObject $searchResult -Name 'Success')
            $searchResultCache[$query] = if ($success) { @(Get-WinTunerDiscoveryValue -InputObject $searchResult -Name 'Results') } else { @() }
        }

        $discoveredByPackageId = @{}
        $results = [System.Collections.Generic.List[object]]::new()
        $processed = 0
        foreach ($entry in $normalizedApps) {
            if (& $ShouldCancel) {
                return New-WinTunerDiscoveryScanResult -Canceled $true -Apps @() -DetectedCount $detectedCount -FilteredCount $filteredCount -NormalizedCount $normalizedCount -MatchedRawCount $matchedRawCount -SkippedNonCandidateCount $skippedCount -GraphFromCache $graphFromCache -GraphPageCount $graphPageCount -GraphLimitReached $graphLimitReached -TotalQueries $totalQueries -CacheHits $cacheHits -WorkerQueries $workerQueries -WorkerCount $workerCount
            }

            $processed++
            $app = $entry.App
            try { & $ReportProgress ([pscustomobject]@{ Stage = 'Matching'; Processed = $processed; Total = $normalizedApps.Count; AppName = [string](Get-WinTunerDiscoveryValue -InputObject $app -Name 'displayName') }) } catch {}
            $wingetResults = if ($searchResultCache.ContainsKey($entry.SearchName)) { @($searchResultCache[$entry.SearchName]) } else { @() }
            $bestMatch = $null
            $highestScore = 0
            foreach ($wingetApp in $wingetResults) {
                $score = Get-WinTunerDiscoverySimilarity -First ([string](Get-WinTunerDiscoveryValue -InputObject $app -Name 'displayName')) -Second ([string](Get-WinTunerDiscoveryValue -InputObject $wingetApp -Name 'Name'))
                if ($score -gt $highestScore) {
                    $highestScore = $score
                    $bestMatch = $wingetApp
                }
            }
            if (-not $bestMatch -or $highestScore -lt 50) { continue }

            $matchedRawCount++
            $packageId = [string](Get-WinTunerDiscoveryValue -InputObject $bestMatch -Name 'PackageID')
            if ([string]::IsNullOrWhiteSpace($packageId) -or $existingPackageIds.Contains($packageId)) { continue }

            $deviceCount = [int](Get-WinTunerDiscoveryValue -InputObject $app -Name 'deviceCount')
            $publisher = [string](Get-WinTunerDiscoveryValue -InputObject $app -Name 'publisher')
            $wingetName = [string](Get-WinTunerDiscoveryValue -InputObject $bestMatch -Name 'Name')
            if ($discoveredByPackageId.ContainsKey($packageId)) {
                $existingEntry = $discoveredByPackageId[$packageId]
                $existingEntry.DeviceCount += $deviceCount
                $existingEntry.MatchScore = [math]::Max([double]$existingEntry.MatchScore, [double]$highestScore)
                $existingEntry.DisplayText = "[$($existingEntry.DeviceCount) PCs] $($existingEntry.DisplayName) ($($existingEntry.Publisher))  -->  Winget: $($existingEntry.WingetApp.Name) [$($existingEntry.WingetApp.PackageID)] | Match: $([math]::Round($existingEntry.MatchScore))%"
            } else {
                $item = [pscustomobject]@{
                    DisplayName = $wingetName
                    Publisher = $publisher
                    DeviceCount = $deviceCount
                    WingetApp = $bestMatch
                    MatchScore = [double]$highestScore
                    Checked = $false
                    DisplayText = "[$deviceCount PCs] $wingetName ($publisher)  -->  Winget: $wingetName [$packageId] | Match: $([math]::Round($highestScore))%"
                }
                $results.Add($item)
                $discoveredByPackageId[$packageId] = $item
            }
        }

        return New-WinTunerDiscoveryScanResult -Canceled $false -Apps @($results) -DetectedCount $detectedCount -FilteredCount $filteredCount -NormalizedCount $normalizedCount -MatchedRawCount $matchedRawCount -SkippedNonCandidateCount $skippedCount -GraphFromCache $graphFromCache -GraphPageCount $graphPageCount -GraphLimitReached $graphLimitReached -TotalQueries $totalQueries -CacheHits $cacheHits -WorkerQueries $workerQueries -WorkerCount $workerCount
    } catch {
        return New-WinTunerDiscoveryScanResult -Canceled $false -ErrorMessage $_.Exception.Message -Apps @() -DetectedCount $detectedCount -FilteredCount $filteredCount -NormalizedCount $normalizedCount -MatchedRawCount $matchedRawCount -SkippedNonCandidateCount $skippedCount -GraphFromCache $graphFromCache -GraphPageCount $graphPageCount -GraphLimitReached $graphLimitReached -TotalQueries $totalQueries -CacheHits $cacheHits -WorkerQueries $workerQueries -WorkerCount $workerCount
    }
}

Export-ModuleMember -Function Invoke-WinTunerDiscoveryScan
