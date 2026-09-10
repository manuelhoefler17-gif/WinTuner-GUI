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

        $result = Invoke-WinTunerPackageUpload `
            -PackageId 'Vendor.App' `
            -Version '2.0.0' `
            -RootPackageFolder $TestDrive `
            -ValidateArtifact $validator `
            -DeployPackage $deployer

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

        $result = Invoke-WinTunerPackageUpload `
            -PackageId 'Vendor.App' `
            -Version '2.0.0' `
            -RootPackageFolder $TestDrive `
            -ValidateArtifact {
                [pscustomobject]@{
                    IsValid    = $false
                    ReasonCode = 'PackageFileMissing'
                    Reason     = 'The metadata-referenced package file is missing.'
                }
            } `
            -DeployPackage $deployer

        $result.Succeeded | Should -BeFalse
        $result.ReasonCode | Should -Be 'PackageFileMissing'
        $result.ErrorMessage | Should -Match 'missing'
        $deployCount | Should -Be 0
    }

    It 'returns deployment errors without losing the validated package path' {
        $expectedPath = Join-Path $TestDrive 'Vendor.App\2.0.0\Vendor.App.intunewin'
        $result = Invoke-WinTunerPackageUpload `
            -PackageId 'Vendor.App' `
            -Version '2.0.0' `
            -RootPackageFolder $TestDrive `
            -ValidateArtifact {
                [pscustomobject]@{
                    IsValid       = $true
                    ReasonCode    = 'Valid'
                    Reason        = 'Valid package artifact.'
                    IntuneWinPath = $expectedPath
                }
            }.GetNewClosure() `
            -DeployPackage { throw 'Tenant rejected the package.' }

        $result.Succeeded | Should -BeFalse
        $result.ReasonCode | Should -Be 'DeploymentFailed'
        $result.ErrorMessage | Should -Match 'Tenant rejected'
        $result.IntuneWinPath | Should -Be $expectedPath
    }
}