BeforeAll {
    foreach($name in 'AppUpdate','DiscoveryDeployment','PackageUpload') {
        Import-Module (Join-Path $PSScriptRoot "..\Modules\WinTuner.$name.psm1") -Force
        & (Get-Module "WinTuner.$name") {
            # Test-only external command substitute: never contacts Intune.
            function script:Deploy-WtWin32App {
                [CmdletBinding()]
                param($GraphId,[switch]$KeepAssignments,$PackageId,$Version,$RootPackageFolder)
                [pscustomobject]@{GraphId=$GraphId;KeepAssignments=[bool]$KeepAssignments;PackageId=$PackageId;Version=$Version;Root=$RootPackageFolder}
            }
        }
    }
}

Describe 'Fixed package operation boundaries' {
    It 'rejects <Parameter> on <Command> before executing supplied code' -ForEach @(
        foreach($command in 'Invoke-WinTunerAppUpdateBatch','Invoke-WinTunerDiscoveryDeployment','Invoke-WinTunerPackageUpload') {
            foreach($parameter in 'BuildPackage','ValidateArtifact','DeployPackage') {@{Command=$command;Parameter=$parameter}}
        }
    ) {
        $state=@{Executed=$false}
        $arguments=@{RootPackageFolder=$TestDrive}
        if($Command -eq 'Invoke-WinTunerPackageUpload'){$arguments.PackageId='Vendor.App';$arguments.Version='1.0'}else{$arguments.Apps=@([pscustomobject]@{Name='Test'})}
        $arguments[$Parameter]={$state.Executed=$true}.GetNewClosure()
        { & $Command @arguments } | Should -Throw '*parameter*'
        $state.Executed | Should -BeFalse
        (Get-Command $Command).Parameters.ContainsKey($Parameter) | Should -BeFalse
    }

    It 'keeps operation helpers private to their modules' {
        foreach($name in 'AppUpdate','DiscoveryDeployment','PackageUpload') {
            $exports=(Get-Module "WinTuner.$name").ExportedFunctions.Keys
            @($exports | Where-Object {$_ -match '^(New-WinTuner(Update|Discovery)Package|Test-WinTuner(Update|Discovery|Upload)Artifact|Publish-WinTuner(Update|Discovery|Upload)Package)$'}) | Should -HaveCount 0
        }
    }

    It 'preserves Graph identity and KeepAssignments in the fixed update deployment call' {
        $result=& (Get-Module WinTuner.AppUpdate) { Publish-WinTunerUpdatePackage -App ([pscustomobject]@{GraphId='old-id'}) -Id 'Vendor.App' -Version '2.0' -Root 'C:\Packages' }
        $result.GraphId | Should -Be 'old-id'
        $result.KeepAssignments | Should -BeTrue
        $result.PackageId | Should -Be 'Vendor.App'
        $result.Version | Should -Be '2.0'
        $result.Root | Should -Be 'C:\Packages'
    }

    It 'passes code-like package text as data through fixed deployment calls' {
        $id='Vendor.App; throw "injected"'
        foreach($spec in @(
            @{Module='WinTuner.DiscoveryDeployment';Helper='Publish-WinTunerDiscoveryPackage'},
            @{Module='WinTuner.PackageUpload';Helper='Publish-WinTunerUploadPackage'}
        )) {
            $result=& (Get-Module $spec.Module) { param($Helper,$Id) & $Helper -Id $Id -Version '2.0' -Root 'C:\Packages' } $spec.Helper $id
            $result.PackageId | Should -BeExactly $id
            $result.Version | Should -Be '2.0'
            $result.KeepAssignments | Should -BeFalse
        }
        $result=& (Get-Module WinTuner.AppUpdate) { param($Id) Publish-WinTunerUpdatePackage -App ([pscustomobject]@{GraphId=''}) -Id $Id -Version '2.0' -Root 'C:\Packages' } $id
        $result.PackageId | Should -BeExactly $id
        $result.KeepAssignments | Should -BeFalse
    }
}
