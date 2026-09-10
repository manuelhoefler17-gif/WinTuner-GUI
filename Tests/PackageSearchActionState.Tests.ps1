BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\Modules\WinTuner.Core.psm1'
    Import-Module $modulePath -Force
}

Describe 'Get-WinTunerPackageSearchActionState' {
    It 'enables only search and query editing before results exist' {
        $state = Get-WinTunerPackageSearchActionState -IsBusy:$false -ResultCount 0 -SelectedIndex -1

        $state.CanSearch | Should -BeTrue
        $state.CanEditQuery | Should -BeTrue
        $state.CanSelectResult | Should -BeFalse
        $state.CanSelectVersion | Should -BeFalse
        $state.CanCreatePackage | Should -BeFalse
    }

    It 'enables result actions for a valid selection' {
        $state = Get-WinTunerPackageSearchActionState -IsBusy:$false -ResultCount 3 -SelectedIndex 1

        $state.CanSelectResult | Should -BeTrue
        $state.CanSelectVersion | Should -BeTrue
        $state.CanCreatePackage | Should -BeTrue
    }

    It 'rejects a stale result selection' {
        $state = Get-WinTunerPackageSearchActionState -IsBusy:$false -ResultCount 2 -SelectedIndex 2

        $state.CanSelectResult | Should -BeTrue
        $state.CanSelectVersion | Should -BeFalse
        $state.CanCreatePackage | Should -BeFalse
    }

    It 'disables every search action while another operation is active' {
        $state = Get-WinTunerPackageSearchActionState -IsBusy:$true -ResultCount 3 -SelectedIndex 0

        $state.CanSearch | Should -BeFalse
        $state.CanEditQuery | Should -BeFalse
        $state.CanSelectResult | Should -BeFalse
        $state.CanSelectVersion | Should -BeFalse
        $state.CanCreatePackage | Should -BeFalse
    }
}