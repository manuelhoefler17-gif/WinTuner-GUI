BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\Modules\WinTuner.Winget.psm1') -Force
}

Describe 'Discovery WinGet result metadata' {
    It 'preserves the target version returned by the isolated worker' {
        $moduleRoot = Join-Path $TestDrive 'Modules'
        $fakeModule = Join-Path $moduleRoot 'WinTuner'
        $null = New-Item -ItemType Directory -Path $fakeModule -Force
        @(
            'function Search-WtWinGetPackage {'
            '    [CmdletBinding()]'
            '    param([string]$SearchQuery)'
            '    [pscustomobject]@{'
            '        Name = $SearchQuery'
            "        PackageID = 'Vendor.Tool'"
            "        Version = '4.2.1'"
            '    }'
            '}'
            'Export-ModuleMember -Function Search-WtWinGetPackage'
        ) | Set-Content -LiteralPath (Join-Path $fakeModule 'WinTuner.psm1') -Encoding utf8

        $inputPath = Join-Path $TestDrive 'queries.json'
        $outputPath = Join-Path $TestDrive 'results.json'
        '"Vendor Tool"' | Set-Content -LiteralPath $inputPath -Encoding utf8

        $originalModulePath = $env:PSModulePath
        try {
            $env:PSModulePath = $moduleRoot
            & (Join-Path $PSScriptRoot '..\Workers\WinTuner.DiscoveryWorker.ps1') -InputPath $inputPath -OutputPath $outputPath
        }
        finally {
            $env:PSModulePath = $originalModulePath
        }

        $result = Get-Content -LiteralPath $outputPath -Raw | ConvertFrom-Json
        $result.Results | Should -HaveCount 1
        $result.Results[0].PackageID | Should -Be 'Vendor.Tool'
        $result.Results[0].Version | Should -Be '4.2.1'
    }

    It 'returns the cached target version to the Discovery scan' {
        InModuleScope WinTuner.Winget {
            $script:discoverySearchCacheLoaded = $true
            $script:discoverySearchCacheDirty = $false
            $script:discoverySearchCache = @{
                'vendor tool' = @{
                    timestamp = [datetime]::UtcNow
                    results = @(
                        [pscustomobject]@{
                            Name = 'Vendor Tool'
                            PackageID = 'Vendor.Tool'
                            Version = '4.2.1'
                        }
                    )
                }
            }

            $result = Search-WinTunerDiscoveryPackagesBatchCached -SearchQueries @('Vendor Tool')

            $result.CacheHits | Should -Be 1
            $result.Results[0].Results[0].Version | Should -Be '4.2.1'
        }
    }

    It 'refreshes legacy cache entries that lack a target version' {
        InModuleScope WinTuner.Winget {
            $script:discoverySearchCacheLoaded = $true
            $script:discoverySearchCacheDirty = $false
            $script:discoverySearchCache = @{
                'legacy tool' = @{
                    timestamp = [datetime]::UtcNow
                    results = @(
                        [pscustomobject]@{
                            Name = 'Legacy Tool'
                            PackageID = 'Vendor.Legacy'
                        }
                    )
                }
            }
            Mock Search-WinTunerPackageWithTimeout {
                @(
                    [pscustomobject]@{
                        Name = 'Legacy Tool'
                        PackageID = 'Vendor.Legacy'
                        Version = '5.0'
                    }
                )
            }

            $result = Search-WinTunerDiscoveryPackageCached -SearchQuery 'Legacy Tool'

            $result.Version | Should -Be '5.0'
            Should -Invoke Search-WinTunerPackageWithTimeout -Times 1 -Exactly
        }
    }}