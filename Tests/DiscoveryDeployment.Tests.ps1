BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\Modules\WinTuner.DiscoveryDeployment.psm1') -Force
}

Describe 'Invoke-WinTunerDiscoveryDeployment' {
    It 'builds, validates, and deploys an exact discovered package' {
        $calls = [System.Collections.Generic.List[string]]::new()
        $app = [pscustomobject]@{
            DisplayName = 'Contoso Tool'
            WingetApp = [pscustomobject]@{ Name = 'Contoso Tool'; PackageID = 'Contoso.Tool'; Version = '2.0' }
        }
        $result = Invoke-WinTunerDiscoveryDeployment -Apps @($app) -RootPackageFolder $TestDrive -BuildPackage {
            param($Id, $Root, $Version)
            $calls.Add("build|$Id|$Version")
            [pscustomobject]@{ Succeeded = $true; EffectiveVersion = $Version }
        }.GetNewClosure() -ValidateArtifact {
            param($Root, $Id, $Version)
            $calls.Add("validate|$Id|$Version")
            [pscustomobject]@{ IsValid = $true; IntuneWinPath = (Join-Path $Root 'package.intunewin') }
        }.GetNewClosure() -DeployPackage {
            param($Id, $Version)
            $calls.Add("deploy|$Id|$Version")
        }.GetNewClosure()

        $result.SuccessCount | Should -Be 1
        $result.Results[0].ReasonCode | Should -Be 'Deployed'
        $calls | Should -Be @('build|Contoso.Tool|2.0', 'validate|Contoso.Tool|2.0', 'deploy|Contoso.Tool|2.0')
    }

    It 'blocks deployment when exact artifact validation fails' {
        $deployCalls = 0
        $app = [pscustomobject]@{ DisplayName = 'Unsafe'; WingetApp = [pscustomobject]@{ Name = 'Unsafe'; PackageID = 'Unsafe.App'; Version = '1.0' } }
        $result = Invoke-WinTunerDiscoveryDeployment -Apps @($app) -RootPackageFolder $TestDrive -BuildPackage {
            [pscustomobject]@{ Succeeded = $true; EffectiveVersion = '1.0' }
        } -ValidateArtifact {
            [pscustomobject]@{ IsValid = $false; ReasonCode = 'MissingIntuneWin'; Reason = 'Expected file is missing.' }
        } -DeployPackage { $deployCalls++ }.GetNewClosure()

        $result.FailureCount | Should -Be 1
        $result.Results[0].ReasonCode | Should -Be 'MissingIntuneWin'
        $deployCalls | Should -Be 0
    }

    It 'continues after a deployment failure' {
        $apps = @(
            [pscustomobject]@{ DisplayName = 'First'; WingetApp = [pscustomobject]@{ Name = 'First'; PackageID = 'Vendor.First'; Version = '2' } }
            [pscustomobject]@{ DisplayName = 'Second'; WingetApp = [pscustomobject]@{ Name = 'Second'; PackageID = 'Vendor.Second'; Version = '3' } }
        )
        $result = Invoke-WinTunerDiscoveryDeployment -Apps $apps -RootPackageFolder $TestDrive -BuildPackage {
            param($Id, $Root, $Version)
            [pscustomobject]@{ Succeeded = $true; EffectiveVersion = $Version }
        } -ValidateArtifact { [pscustomobject]@{ IsValid = $true } } -DeployPackage {
            param($Id)
            if ($Id -eq 'Vendor.First') { throw 'tenant rejected deployment' }
        }

        $result.SuccessCount | Should -Be 1
        $result.FailureCount | Should -Be 1
        $result.Results[0].ReasonCode | Should -Be 'DeploymentFailed'
        $result.Results[1].ReasonCode | Should -Be 'Deployed'
    }

    It 'blocks package creation when PackageId or target version is missing' {
        $buildCalls = 0
        $apps = @(
            [pscustomobject]@{ DisplayName = 'No Id'; WingetApp = [pscustomobject]@{ Name = 'No Id'; PackageID = ''; Version = '1' } }
            [pscustomobject]@{ DisplayName = 'No Version'; WingetApp = [pscustomobject]@{ Name = 'No Version'; PackageID = 'No.Version'; Version = '' } }
        )
        $result = Invoke-WinTunerDiscoveryDeployment -Apps $apps -RootPackageFolder $TestDrive -BuildPackage { $buildCalls++ }.GetNewClosure()
        $result.FailureCount | Should -Be 2
        $result.Results.ReasonCode | Should -Be @('MissingPackageId', 'MissingTargetVersion')
        $buildCalls | Should -Be 0
    }

    It 'reports package-build exceptions without attempting validation or deployment' {
        $laterCalls = 0
        $app = [pscustomobject]@{ DisplayName = 'Build Fails'; WingetApp = [pscustomobject]@{ Name = 'Build Fails'; PackageID = 'Build.Fails'; Version = '1' } }
        $result = Invoke-WinTunerDiscoveryDeployment -Apps @($app) -RootPackageFolder $TestDrive -BuildPackage {
            throw 'download failed'
        } -ValidateArtifact { $laterCalls++ }.GetNewClosure() -DeployPackage { $laterCalls++ }.GetNewClosure()

        $result.Results[0].ReasonCode | Should -Be 'PackageBuildFailed'
        $laterCalls | Should -Be 0
    }

    It 'reports validation exceptions without attempting deployment' {
        $deployCalls = 0
        $app = [pscustomobject]@{ DisplayName = 'Validation Fails'; WingetApp = [pscustomobject]@{ Name = 'Validation Fails'; PackageID = 'Validation.Fails'; Version = '1' } }
        $result = Invoke-WinTunerDiscoveryDeployment -Apps @($app) -RootPackageFolder $TestDrive -BuildPackage {
            [pscustomobject]@{ Succeeded = $true; EffectiveVersion = '1' }
        } -ValidateArtifact {
            throw 'metadata unreadable'
        } -DeployPackage { $deployCalls++ }.GetNewClosure()

        $result.Results[0].ReasonCode | Should -Be 'ValidationError'
        $deployCalls | Should -Be 0
    }
}
