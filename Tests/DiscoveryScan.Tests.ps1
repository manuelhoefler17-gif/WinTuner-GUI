BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\Modules\WinTuner.DiscoveryScan.psm1') -Force
}

Describe 'Invoke-WinTunerDiscoveryScan' {
    It 'filters, matches, deduplicates, and preserves cache metrics' {
        $detected = @(
            [pscustomobject]@{ displayName = 'Contoso Tool 2.0'; publisher = 'Contoso'; deviceCount = 3 }
            [pscustomobject]@{ displayName = 'Contoso Tool (x64)'; publisher = 'Contoso'; deviceCount = 2 }
            [pscustomobject]@{ displayName = 'Managed App'; publisher = 'Fabrikam'; deviceCount = 4 }
            [pscustomobject]@{ displayName = 'Intel Driver'; publisher = 'Intel'; deviceCount = 99 }
        )
        $result = Invoke-WinTunerDiscoveryScan -ConnectGraph { $true } -GetExistingApps {
            @([pscustomobject]@{ PackageId = 'Managed.App' })
        } -ResolvePackageId {
            param($App)
            $App.PackageId
        } -GetDetectedApps {
            [pscustomobject]@{ Apps = $detected; FromCache = $true; PageCount = 2; LimitReached = $false }
        }.GetNewClosure() -SearchPackages {
            param($Queries)
            [pscustomobject]@{
                Canceled = $false
                TotalQueries = @($Queries).Count
                CacheHits = 2
                WorkerQueries = 1
                WorkerCount = 1
                Results = @(
                    [pscustomobject]@{ Query = 'Contoso Tool'; Success = $true; Results = @([pscustomobject]@{ Name = 'Contoso Tool'; PackageID = 'Contoso.Tool'; Version = '2.0' }) }
                    [pscustomobject]@{ Query = 'Managed App'; Success = $true; Results = @([pscustomobject]@{ Name = 'Managed App'; PackageID = 'Managed.App'; Version = '1.0' }) }
                )
            }
        }

        $result.ErrorMessage | Should -BeNullOrEmpty
        $result.Canceled | Should -BeFalse
        $result.DetectedCount | Should -Be 4
        $result.FilteredCount | Should -Be 3
        $result.NormalizedCount | Should -Be 3
        $result.MatchedRawCount | Should -Be 3
        $result.UniquePackageCount | Should -Be 1
        $result.Apps[0].WingetApp.PackageID | Should -Be 'Contoso.Tool'
        $result.Apps[0].DeviceCount | Should -Be 5
        $result.GraphFromCache | Should -BeTrue
        $result.CacheHits | Should -Be 2
        $result.WorkerQueries | Should -Be 1
    }

    It 'discards partial results when the WinGet phase is canceled' {
        $result = Invoke-WinTunerDiscoveryScan -ConnectGraph { $true } -GetExistingApps { @() } -ResolvePackageId { '' } -GetDetectedApps {
            [pscustomobject]@{ Apps = @([pscustomobject]@{ displayName = 'Contoso Tool'; publisher = 'Contoso'; deviceCount = 1 }); FromCache = $false }
        } -SearchPackages {
            [pscustomobject]@{ Canceled = $true; TotalQueries = 1; CacheHits = 0; WorkerQueries = 1; WorkerCount = 1; Results = @() }
        }

        $result.Canceled | Should -BeTrue
        $result.Apps | Should -HaveCount 0
    }

    It 'returns a structured error and no candidates when Graph retrieval fails' {
        $result = Invoke-WinTunerDiscoveryScan -ConnectGraph { $true } -GetExistingApps { @() } -ResolvePackageId { '' } -GetDetectedApps {
            throw 'Graph paging failed'
        } -SearchPackages { throw 'must not run' }

        $result.Canceled | Should -BeFalse
        $result.ErrorMessage | Should -Be 'Graph paging failed'
        $result.Apps | Should -HaveCount 0
    }

    It 'can skip low-value mobile identifiers before WinGet search' {
        $queryState = [pscustomobject]@{ Called = $false }
        $result = Invoke-WinTunerDiscoveryScan -ConnectGraph { $true } -GetExistingApps { @() } -ResolvePackageId { '' } -GetDetectedApps {
            [pscustomobject]@{ Apps = @([pscustomobject]@{ displayName = 'com.vendor.mobile'; publisher = 'Vendor'; deviceCount = 1 }) }
        } -SkipLowValueCandidates $true -SearchPackages {
            param($Queries)
            $queryState.Called = $true
            [pscustomobject]@{ Canceled = $false; TotalQueries = 0; CacheHits = 0; WorkerQueries = 0; WorkerCount = 0; Results = @() }
        }.GetNewClosure()

        $result.SkippedNonCandidateCount | Should -Be 1
        $result.NormalizedCount | Should -Be 0
        $queryState.Called | Should -BeFalse
    }

    It 'normalizes fresh Graph dictionaries as well as cached objects' {
        $result = Invoke-WinTunerDiscoveryScan -ConnectGraph { $true } -GetExistingApps { @() } -ResolvePackageId { '' } -GetDetectedApps {
            [ordered]@{
                Apps = @([ordered]@{ displayName = 'Contoso Tool'; publisher = 'Contoso'; deviceCount = 2 })
                FromCache = $false
                PageCount = 1
                LimitReached = $false
            }
        } -SearchPackages {
            param($Queries)
            [pscustomobject]@{
                Canceled = $false
                TotalQueries = @($Queries).Count
                CacheHits = 0
                WorkerQueries = 1
                WorkerCount = 1
                Results = @([ordered]@{
                    Query = 'Contoso Tool'
                    Success = $true
                    Results = @([ordered]@{ Name = 'Contoso Tool'; PackageID = 'Contoso.Tool'; Version = '1.0' })
                })
            }
        }

        $result.ErrorMessage | Should -BeNullOrEmpty
        $result.NormalizedCount | Should -Be 1
        $result.Apps | Should -HaveCount 1
        $result.Apps[0].WingetApp.PackageID | Should -Be 'Contoso.Tool'
        $result.GraphFromCache | Should -BeFalse
    }

    It 'retries a transient existing-app enumeration failure' {
        $state = [pscustomobject]@{ ExistingCalls = 0 }
        $result = Invoke-WinTunerDiscoveryScan -ConnectGraph { $true } -GetExistingApps {
            $state.ExistingCalls++
            if ($state.ExistingCalls -eq 1) { throw 'Collection was modified; enumeration operation may not execute.' }
            @([pscustomobject]@{ PackageId = 'Contoso.Tool' })
        }.GetNewClosure() -ResolvePackageId {
            param($App)
            $App.PackageId
        } -GetDetectedApps {
            [pscustomobject]@{ Apps = @([pscustomobject]@{ displayName = 'Contoso Tool'; publisher = 'Contoso'; deviceCount = 1 }) }
        } -SearchPackages {
            [pscustomobject]@{
                Canceled = $false
                TotalQueries = 1
                CacheHits = 0
                WorkerQueries = 1
                WorkerCount = 1
                Results = @([pscustomobject]@{
                    Query = 'Contoso Tool'
                    Success = $true
                    Results = @([pscustomobject]@{ Name = 'Contoso Tool'; PackageID = 'Contoso.Tool'; Version = '1.0' })
                })
            }
        } -ExistingAppsRetryDelayMilliseconds 0

        $state.ExistingCalls | Should -Be 2
        $result.ErrorMessage | Should -BeNullOrEmpty
        $result.Apps | Should -HaveCount 0
    }
}
