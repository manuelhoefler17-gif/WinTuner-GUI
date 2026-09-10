BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\Modules\WinTuner.Core.psm1'
    Import-Module $modulePath -Force
}

Describe 'Get-WinTunerDiscoveryActionState' {
    It 'disables every discovery action while disconnected' {
        $state = Get-WinTunerDiscoveryActionState -Connected:$false -IsScanning:$false -CancelRequested:$false -IsDeploying:$false -ResultCount 3 -CheckedCount 1

        $state.CanScan | Should -BeFalse
        $state.CanDeploy | Should -BeFalse
        $state.CanExport | Should -BeFalse
        $state.CanCheckAll | Should -BeFalse
        $state.CanUncheckAll | Should -BeFalse
    }

    It 'enables only scanning before results exist' {
        $state = Get-WinTunerDiscoveryActionState -Connected:$true -IsScanning:$false -CancelRequested:$false -IsDeploying:$false -ResultCount 0 -CheckedCount 0

        $state.CanScan | Should -BeTrue
        $state.CanDeploy | Should -BeFalse
        $state.CanExport | Should -BeFalse
        $state.CanCheckAll | Should -BeFalse
        $state.CanUncheckAll | Should -BeFalse
    }

    It 'enables result actions but not deployment before a selection' {
        $state = Get-WinTunerDiscoveryActionState -Connected:$true -IsScanning:$false -CancelRequested:$false -IsDeploying:$false -ResultCount 3 -CheckedCount 0

        $state.CanScan | Should -BeTrue
        $state.CanDeploy | Should -BeFalse
        $state.CanExport | Should -BeTrue
        $state.CanCheckAll | Should -BeTrue
        $state.CanUncheckAll | Should -BeFalse
    }

    It 'enables deployment when at least one result is selected' {
        $state = Get-WinTunerDiscoveryActionState -Connected:$true -IsScanning:$false -CancelRequested:$false -IsDeploying:$false -ResultCount 3 -CheckedCount 1

        $state.CanDeploy | Should -BeTrue
        $state.CanExport | Should -BeTrue
        $state.CanCheckAll | Should -BeTrue
        $state.CanUncheckAll | Should -BeTrue
    }

    It 'disables Check All when every result is selected' {
        $state = Get-WinTunerDiscoveryActionState -Connected:$true -IsScanning:$false -CancelRequested:$false -IsDeploying:$false -ResultCount 3 -CheckedCount 3

        $state.CanCheckAll | Should -BeFalse
        $state.CanUncheckAll | Should -BeTrue
        $state.CanDeploy | Should -BeTrue
    }

    It 'leaves only the scan button available as Cancel Scan while scanning' {
        $state = Get-WinTunerDiscoveryActionState -Connected:$true -IsScanning:$true -CancelRequested:$false -IsDeploying:$false -ResultCount 3 -CheckedCount 1

        $state.CanScan | Should -BeTrue
        $state.CanDeploy | Should -BeFalse
        $state.CanExport | Should -BeFalse
        $state.CanCheckAll | Should -BeFalse
        $state.CanUncheckAll | Should -BeFalse
    }

    It 'disables every discovery action after cancellation is requested' {
        $state = Get-WinTunerDiscoveryActionState -Connected:$true -IsScanning:$true -CancelRequested:$true -IsDeploying:$false -ResultCount 3 -CheckedCount 1

        $state.CanScan | Should -BeFalse
        $state.CanDeploy | Should -BeFalse
        $state.CanExport | Should -BeFalse
        $state.CanCheckAll | Should -BeFalse
        $state.CanUncheckAll | Should -BeFalse
    }

    It 'disables every discovery action during deployment' {
        $state = Get-WinTunerDiscoveryActionState -Connected:$true -IsScanning:$false -CancelRequested:$false -IsDeploying:$true -ResultCount 3 -CheckedCount 1

        $state.CanScan | Should -BeFalse
        $state.CanDeploy | Should -BeFalse
        $state.CanExport | Should -BeFalse
        $state.CanCheckAll | Should -BeFalse
        $state.CanUncheckAll | Should -BeFalse
    }

    It 'disables discovery actions while another workflow is active' {
        $state = Get-WinTunerDiscoveryActionState -Connected:$true -IsScanning:$false -CancelRequested:$false -IsDeploying:$false -IsOtherOperationActive:$true -ResultCount 3 -CheckedCount 1

        $state.CanScan | Should -BeFalse
        $state.CanDeploy | Should -BeFalse
        $state.CanExport | Should -BeFalse
        $state.CanCheckAll | Should -BeFalse
        $state.CanUncheckAll | Should -BeFalse
    }
}
