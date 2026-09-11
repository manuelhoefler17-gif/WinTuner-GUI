BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\Modules\WinTuner.Connection.psm1') -Force
}

Describe 'Invoke-WinTunerConnectionVerification' {
    It 'succeeds on the first tenant query even when no apps exist' {
        $result = Invoke-WinTunerConnectionVerification -GetApps { @() } -Delay { throw 'delay must not run' }
        $result.Succeeded | Should -BeTrue
        $result.Attempts | Should -Be 1
    }

    It 'retries transient failures and then succeeds' {
        $state = [pscustomobject]@{ Attempts = 0 }
        $delays = [System.Collections.Generic.List[int]]::new()
        $result = Invoke-WinTunerConnectionVerification -GetApps {
            $state.Attempts++
            if ($state.Attempts -lt 3) { throw 'temporary connection failure' }
            @([pscustomobject]@{ Name = 'App' })
        }.GetNewClosure() -MaxAttempts 4 -RetryDelayMilliseconds 25 -Delay {
            param($Milliseconds)
            $delays.Add($Milliseconds)
        }.GetNewClosure()

        $result.Succeeded | Should -BeTrue
        $result.Attempts | Should -Be 3
        $delays | Should -Be @(25, 25)
    }

    It 'returns the final verification error after all attempts fail' {
        $result = Invoke-WinTunerConnectionVerification -GetApps { throw 'authentication expired' } -MaxAttempts 3 -RetryDelayMilliseconds 0
        $result.Succeeded | Should -BeFalse
        $result.Attempts | Should -Be 3
        $result.ErrorMessage | Should -Be 'authentication expired'
    }
}
