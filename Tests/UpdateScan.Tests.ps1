BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\Modules\WinTuner.UpdateScan.psm1'
    Import-Module $modulePath -Force
}

Describe 'Invoke-WinTunerUpdateScan' {
    It 'returns verified and fallback candidates with stable deployment fields' {
        $apps = @(
            [pscustomobject]@{ Name = 'Verified'; CurrentVersion = '1.0'; LatestVersion = '1.1'; GraphId = 'graph-1'; PackageId = 'Vendor.Verified' },
            [pscustomobject]@{ Name = 'Fallback'; CurrentVersion = '2.0'; LatestVersion = '2.5'; GraphId = 'graph-2'; PackageId = 'Vendor.Fallback' },
            [pscustomobject]@{ Name = 'Current'; CurrentVersion = '3.0'; LatestVersion = '3.0'; GraphId = 'graph-3'; PackageId = 'Vendor.Current' }
        )

        $result = Invoke-WinTunerUpdateScan `
            -GetApps { $apps }.GetNewClosure() `
            -ResolvePackageId { param($app) $app.PackageId } `
            -GetVersions { param($packageId) if ($packageId -eq 'Vendor.Verified') { '1.2' } elseif ($packageId -eq 'Vendor.Fallback') { throw 'lookup failed' } } `
            -IsNewerVersion { param($latest, $current) [version]$latest -gt [version]$current }

        $result.Canceled | Should -BeFalse
        $result.TotalApps | Should -Be 3
        $result.CheckedApps | Should -Be 3
        $result.CandidateCount | Should -Be 2
        $result.LookupFailureCount | Should -Be 1
        @($result.Candidates.Name) | Should -Be @('Verified', 'Fallback')
        $result.Candidates[0].PackageId | Should -Be 'Vendor.Verified'
        $result.Candidates[0].LatestVersion | Should -Be '1.2'
        $result.Candidates[0].Checked | Should -BeFalse
    }

    It 'cancels before the first lookup and returns no partial candidates' {
        $apps = @([pscustomobject]@{ Name = 'One'; CurrentVersion = '1.0'; LatestVersion = '2.0'; PackageId = 'Vendor.One' })
        $result = Invoke-WinTunerUpdateScan `
            -GetApps { $apps }.GetNewClosure() `
            -ResolvePackageId { param($app) $app.PackageId } `
            -GetVersions { param($packageId) '2.0' } `
            -IsNewerVersion { param($latest, $current) $true } `
            -ShouldCancel { $true }

        $result.Canceled | Should -BeTrue
        $result.CheckedApps | Should -Be 0
        $result.CandidateCount | Should -Be 0
    }

    It 'stops between app lookups when cancellation is requested' {
        $apps = @(
            [pscustomobject]@{ Name = 'One'; CurrentVersion = '1.0'; LatestVersion = '2.0'; PackageId = 'Vendor.One' },
            [pscustomobject]@{ Name = 'Two'; CurrentVersion = '1.0'; LatestVersion = '2.0'; PackageId = 'Vendor.Two' }
        )
        $cancelState = [pscustomobject]@{ Checks = 0 }

        $result = Invoke-WinTunerUpdateScan `
            -GetApps { $apps }.GetNewClosure() `
            -ResolvePackageId { param($app) $app.PackageId } `
            -GetVersions { param($packageId) '2.0' } `
            -IsNewerVersion { param($latest, $current) $true } `
            -ShouldCancel { [void]($cancelState.Checks++); $cancelState.Checks -ge 3 }.GetNewClosure()

        $result.Canceled | Should -BeTrue
        $result.CheckedApps | Should -Be 1
        $result.CandidateCount | Should -Be 1
    }

    It 'uses the tenant version when package resolution fails' {
        $apps = @([pscustomobject]@{ Name = 'Fallback'; CurrentVersion = '1.0'; LatestVersion = '1.5' })

        $result = Invoke-WinTunerUpdateScan `
            -GetApps { $apps }.GetNewClosure() `
            -ResolvePackageId { param($app) throw 'search failed' } `
            -GetVersions { param($packageId) @() } `
            -IsNewerVersion { param($latest, $current) [version]$latest -gt [version]$current }

        $result.ErrorMessage | Should -BeNullOrEmpty
        $result.LookupFailureCount | Should -Be 1
        $result.CandidateCount | Should -Be 1
        $result.Candidates[0].LatestVersion | Should -Be '1.5'
    }

    It 'reports a load failure without throwing' {
        $result = Invoke-WinTunerUpdateScan `
            -GetApps { throw 'tenant unavailable' } `
            -ResolvePackageId { param($app) $null } `
            -GetVersions { param($packageId) @() } `
            -IsNewerVersion { param($latest, $current) $false }

        $result.ErrorMessage | Should -Be 'tenant unavailable'
        $result.CandidateCount | Should -Be 0
    }

    It 'ignores apps that do not have an installed version' {
        $apps = @([pscustomobject]@{ Name = 'Unknown'; CurrentVersion = ''; LatestVersion = '2.0'; PackageId = 'Vendor.Unknown' })

        $result = Invoke-WinTunerUpdateScan `
            -GetApps { $apps }.GetNewClosure() `
            -ResolvePackageId { param($app) $app.PackageId } `
            -GetVersions { param($packageId) '2.0' } `
            -IsNewerVersion { param($latest, $current) $true }

        $result.TotalApps | Should -Be 0
        $result.CheckedApps | Should -Be 0
    }
}
