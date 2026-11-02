BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\Modules\WinTuner.PackageUpload.psm1'
    Import-Module $modulePath -Force
}

Describe 'Invoke-WinTunerPackageUpload' {
    It 'validates the exact artifact before deploying it' {
        $events = [System.Collections.Generic.List[string]]::new()
        $validator = {
            param($Root, $Id, $Version)
            $events.Add("validate|$Root|$Id|$Version")
            [pscustomobject]@{
                IsValid      = $true
                ReasonCode   = 'Valid'
                Reason       = 'Valid package artifact.'
                IntuneWinPath = Join-Path $Root 'Vendor.App\2.0.0\Vendor.App.intunewin'
            }
        }.GetNewClosure()
        $deployer = {
            param($Id, $Version, $Root)
            $events.Add("deploy|$Root|$Id|$Version")
        }.GetNewClosure()

        Mock Test-WinTunerUploadArtifact -ModuleName WinTuner.PackageUpload -MockWith $validator
        Mock Publish-WinTunerUploadPackage -ModuleName WinTuner.PackageUpload -MockWith $deployer
        $result = Invoke-WinTunerPackageUpload `
            -PackageId 'Vendor.App' `
            -Version '2.0.0' `
            -RootPackageFolder $TestDrive

        $result.Succeeded | Should -BeTrue
        $result.ReasonCode | Should -Be 'Uploaded'
        $events | Should -HaveCount 2
        $events[0] | Should -Be "validate|$TestDrive|Vendor.App|2.0.0"
        $events[1] | Should -Be "deploy|$TestDrive|Vendor.App|2.0.0"
    }

    It 'blocks deployment when exact artifact validation fails' {
        $deployCount = 0
        $deployer = {
            param($Id, $Version, $Root)
            $deployCount++
        }.GetNewClosure()

        Mock Test-WinTunerUploadArtifact -ModuleName WinTuner.PackageUpload -MockWith {
                [pscustomobject]@{
                    IsValid    = $false
                    ReasonCode = 'PackageFileMissing'
                    Reason     = 'The metadata-referenced package file is missing.'
                }
            }
        Mock Publish-WinTunerUploadPackage -ModuleName WinTuner.PackageUpload -MockWith $deployer
        $result = Invoke-WinTunerPackageUpload `
            -PackageId 'Vendor.App' `
            -Version '2.0.0' `
            -RootPackageFolder $TestDrive

        $result.Succeeded | Should -BeFalse
        $result.ReasonCode | Should -Be 'PackageFileMissing'
        $result.ErrorMessage | Should -Match 'missing'
        $deployCount | Should -Be 0
    }

    It 'returns deployment errors without losing the validated package path' {
        $expectedPath = Join-Path $TestDrive 'Vendor.App\2.0.0\Vendor.App.intunewin'
        Mock Test-WinTunerUploadArtifact -ModuleName WinTuner.PackageUpload -MockWith {
                [pscustomobject]@{
                    IsValid       = $true
                    ReasonCode    = 'Valid'
                    Reason        = 'Valid package artifact.'
                    IntuneWinPath = $expectedPath
                }
            }.GetNewClosure()
        Mock Publish-WinTunerUploadPackage -ModuleName WinTuner.PackageUpload -MockWith { throw 'Tenant rejected the package.' }
        $result = Invoke-WinTunerPackageUpload `
            -PackageId 'Vendor.App' `
            -Version '2.0.0' `
            -RootPackageFolder $TestDrive

        $result.Succeeded | Should -BeFalse
        $result.ReasonCode | Should -Be 'DeploymentFailed'
        $result.ErrorMessage | Should -Match 'Tenant rejected'
        $result.IntuneWinPath | Should -Be $expectedPath
    }
}