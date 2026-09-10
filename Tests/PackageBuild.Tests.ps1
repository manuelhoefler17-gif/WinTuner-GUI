BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\Modules\WinTuner.PackageBuild.psm1'
    Import-Module $modulePath -Force
}

Describe 'Invoke-WinTunerPackageBuild' {
    It 'builds the requested version directly' {
        $calls = [System.Collections.Generic.List[string]]::new()
        $creator = {
            param($Id, $Folder, $Version)
            $calls.Add("$Id|$Folder|$Version")
        }.GetNewClosure()

        $result = Invoke-WinTunerPackageBuild `
            -PackageId 'Vendor.App' `
            -PackageFolder $TestDrive `
            -DesiredVersion '2.0.0' `
            -LatestVersion '2.1.0' `
            -CreatePackage $creator

        $result.Succeeded | Should -BeTrue
        $result.EffectiveVersion | Should -Be '2.0.0'
        $result.ChoiceRequired | Should -BeNullOrEmpty
        $calls | Should -HaveCount 1
        $calls[0] | Should -Be "Vendor.App|$TestDrive|2.0.0"
    }

    It 'falls back to the previous version after a not-found response' {
        $calls = [System.Collections.Generic.List[string]]::new()
        $creator = {
            param($Id, $Folder, $Version)
            $calls.Add($Version)
            if ($Version -eq '3.0.0') { throw '404 Not Found' }
        }.GetNewClosure()

        $result = Invoke-WinTunerPackageBuild `
            -PackageId 'Vendor.App' `
            -PackageFolder $TestDrive `
            -LatestVersion '3.0.0' `
            -CreatePackage $creator `
            -GetPreviousVersion { param($Id, $Version) '2.9.0' }

        $result.Succeeded | Should -BeTrue
        $result.EffectiveVersion | Should -Be '2.9.0'
        $calls | Should -HaveCount 2
        $calls[0] | Should -Be '3.0.0'
        $calls[1] | Should -Be '2.9.0'
    }

    It 'returns a UI choice request after a hash mismatch' {
        $result = Invoke-WinTunerPackageBuild `
            -PackageId 'Vendor.App' `
            -PackageFolder $TestDrive `
            -LatestVersion '4.0.0' `
            -CreatePackage { throw 'Hash mismatch detected' }

        $result.Succeeded | Should -BeFalse
        $result.ChoiceRequired | Should -Be 'HashMismatch'
        $result.ErrorMessage | Should -Match 'Hash mismatch'
    }

    It 'does not request another choice when retrying the same version fails' {
        $result = Invoke-WinTunerPackageBuild `
            -PackageId 'Vendor.App' `
            -PackageFolder $TestDrive `
            -LatestVersion '4.0.0' `
            -Mode RetrySame `
            -CreatePackage { throw 'Hash mismatch detected again' }

        $result.Succeeded | Should -BeFalse
        $result.ChoiceRequired | Should -BeNullOrEmpty
        $result.ErrorMessage | Should -Match 'Hash mismatch'
    }

    It 'blocks an automatic fallback that is not newer than the installed version' {
        $calls = [System.Collections.Generic.List[string]]::new()
        $creator = {
            param($Id, $Folder, $Version)
            $calls.Add($Version)
            throw '404 Not Found'
        }.GetNewClosure()

        $result = Invoke-WinTunerPackageBuild `
            -PackageId 'Vendor.App' `
            -PackageFolder $TestDrive `
            -LatestVersion '5.0.0' `
            -InstalledVersion '4.9.0' `
            -CreatePackage $creator `
            -GetPreviousVersion { param($Id, $Version) '4.8.0' } `
            -IsNewerVersion { param($Candidate, $Installed) $false }

        $result.Succeeded | Should -BeFalse
        $result.ErrorMessage | Should -Match '404'
        $calls | Should -HaveCount 1
    }
}