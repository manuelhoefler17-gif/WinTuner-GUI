BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\Modules\WinTuner.Core.psm1'
    Import-Module $modulePath -Force
}

Describe 'Get-WinTunerUpdateActionState' {
    It 'disables every update action while disconnected' {
        $state = Get-WinTunerUpdateActionState -Connected:$false -IsBusy:$false -CandidateCount 3 -CheckedCount 2

        $state.CanSearch | Should -BeFalse
        $state.CanCheckAll | Should -BeFalse
        $state.CanUncheckAll | Should -BeFalse
        $state.CanUpdateSelected | Should -BeFalse
        $state.CanUpdateAll | Should -BeFalse
        $state.CanLogout | Should -BeFalse
    }

    It 'enables only search before candidates exist' {
        $state = Get-WinTunerUpdateActionState -Connected:$true -IsBusy:$false -CandidateCount 0 -CheckedCount 0

        $state.CanSearch | Should -BeTrue
        $state.CanCheckAll | Should -BeFalse
        $state.CanUncheckAll | Should -BeFalse
        $state.CanUpdateSelected | Should -BeFalse
        $state.CanUpdateAll | Should -BeFalse
        $state.CanLogout | Should -BeTrue
    }

    It 'enables candidate actions but not checked actions before a selection' {
        $state = Get-WinTunerUpdateActionState -Connected:$true -IsBusy:$false -CandidateCount 3 -CheckedCount 0

        $state.CanSearch | Should -BeTrue
        $state.CanCheckAll | Should -BeTrue
        $state.CanUncheckAll | Should -BeFalse
        $state.CanUpdateSelected | Should -BeFalse
        $state.CanUpdateAll | Should -BeTrue
    }

    It 'enables checked actions when at least one candidate is selected' {
        $state = Get-WinTunerUpdateActionState -Connected:$true -IsBusy:$false -CandidateCount 3 -CheckedCount 1

        $state.CanCheckAll | Should -BeTrue
        $state.CanUncheckAll | Should -BeTrue
        $state.CanUpdateSelected | Should -BeTrue
        $state.CanUpdateAll | Should -BeTrue
    }

    It 'disables Check All when every candidate is selected' {
        $state = Get-WinTunerUpdateActionState -Connected:$true -IsBusy:$false -CandidateCount 3 -CheckedCount 3

        $state.CanCheckAll | Should -BeFalse
        $state.CanUncheckAll | Should -BeTrue
        $state.CanUpdateSelected | Should -BeTrue
    }

    It 'disables every update action while an operation is active' {
        $state = Get-WinTunerUpdateActionState -Connected:$true -IsBusy:$true -CandidateCount 3 -CheckedCount 2

        $state.CanSearch | Should -BeFalse
        $state.CanCheckAll | Should -BeFalse
        $state.CanUncheckAll | Should -BeFalse
        $state.CanUpdateSelected | Should -BeFalse
        $state.CanUpdateAll | Should -BeFalse
        $state.CanLogout | Should -BeFalse
        $state.SearchButtonText | Should -Be 'Search Updates'
    }

    It 'turns Search Updates into an enabled cancel action during a scan' {
        $state = Get-WinTunerUpdateActionState -Connected:$true -IsBusy:$true -IsScanRunning:$true -CancelRequested:$false -CandidateCount 3 -CheckedCount 2

        $state.CanSearch | Should -BeTrue
        $state.SearchButtonText | Should -Be 'Cancel Scan'
        $state.CanUpdateSelected | Should -BeFalse
        $state.CanLogout | Should -BeFalse
    }

    It 'disables the cancel action after cancellation was requested' {
        $state = Get-WinTunerUpdateActionState -Connected:$true -IsBusy:$true -IsScanRunning:$true -CancelRequested:$true -CandidateCount 3 -CheckedCount 2

        $state.CanSearch | Should -BeFalse
        $state.SearchButtonText | Should -Be 'Cancelling...'
    }
}

Describe 'GUI update result summary wiring' {
    BeforeAll {
        $guiPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'WinTuner_GUI.ps1'
        $script:updateSummaryGuiText = [System.IO.File]::ReadAllText($guiPath)
    }

    It 'keeps the latest batch result visible in the update summary' {
        $script:updateSummaryGuiText | Should -Match '\$script:lastUpdateResultSummary'
        $script:updateSummaryGuiText | Should -Match 'Last update: \$successCount successful, \$failureCount failed'
        $script:updateSummaryGuiText | Should -Match '\$updateSummaryLabel\.Text = if'
    }

    It 'keeps failed candidates available and explains that they can be retried' {
        $script:updateSummaryGuiText | Should -Match 'Failed candidates remain available for retry\.'
    }

    It 'clears the previous update result when the tenant session ends' {
        @([regex]::Matches($script:updateSummaryGuiText, '\$script:lastUpdateResultSummary = ''''')).Count | Should -BeGreaterOrEqual 2
    }
}
