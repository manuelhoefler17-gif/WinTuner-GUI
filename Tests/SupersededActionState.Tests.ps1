BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\Modules\WinTuner.Core.psm1'
    Import-Module $modulePath -Force
}

Describe 'Get-WinTunerSupersededActionState' {
    It 'disables every action while disconnected' {
        $state = Get-WinTunerSupersededActionState -Connected:$false -IsBusy:$false -ResultCount 2 -SelectedIndex 0

        $state.CanSearch | Should -BeFalse
        $state.CanDeleteSelected | Should -BeFalse
        $state.CanDeleteAll | Should -BeFalse
        $state.CanLogout | Should -BeFalse
    }

    It 'enables only search before results exist' {
        $state = Get-WinTunerSupersededActionState -Connected:$true -IsBusy:$false -ResultCount 0 -SelectedIndex -1

        $state.CanSearch | Should -BeTrue
        $state.CanDeleteSelected | Should -BeFalse
        $state.CanDeleteAll | Should -BeFalse
        $state.CanLogout | Should -BeTrue
    }

    It 'enables deletion for a valid result selection' {
        $state = Get-WinTunerSupersededActionState -Connected:$true -IsBusy:$false -ResultCount 2 -SelectedIndex 1

        $state.CanSearch | Should -BeTrue
        $state.CanDeleteSelected | Should -BeTrue
        $state.CanDeleteAll | Should -BeTrue
    }

    It 'rejects a stale selection index' {
        $state = Get-WinTunerSupersededActionState -Connected:$true -IsBusy:$false -ResultCount 2 -SelectedIndex 2

        $state.CanDeleteSelected | Should -BeFalse
        $state.CanDeleteAll | Should -BeTrue
    }

    It 'disables every action while an operation is active' {
        $state = Get-WinTunerSupersededActionState -Connected:$true -IsBusy:$true -ResultCount 2 -SelectedIndex 0

        $state.CanSearch | Should -BeFalse
        $state.CanDeleteSelected | Should -BeFalse
        $state.CanDeleteAll | Should -BeFalse
        $state.CanLogout | Should -BeFalse
    }
}