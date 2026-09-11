BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\Modules\WinTuner.SupersededRemoval.psm1') -Force
}

Describe 'Invoke-WinTunerSupersededRemoval' {
    It 'removes every verified app and reports progress' {
        $removed = [System.Collections.Generic.List[string]]::new()
        $progress = [System.Collections.Generic.List[string]]::new()
        $apps = @(
            [pscustomobject]@{ Name = 'Old One'; GraphId = 'graph-1' }
            [pscustomobject]@{ Name = 'Old Two'; GraphId = 'graph-2' }
        )
        $result = Invoke-WinTunerSupersededRemoval -Apps $apps -RemoveApp {
            param($App)
            $removed.Add([string]$App.GraphId)
        }.GetNewClosure() -ReportProgress {
            param($Info)
            $progress.Add("$($Info.Processed)/$($Info.Total):$($Info.AppName)")
        }.GetNewClosure()

        $result.SuccessCount | Should -Be 2
        $result.FailureCount | Should -Be 0
        $removed | Should -Be @('graph-1', 'graph-2')
        $progress | Should -Be @('1/2:Old One', '2/2:Old Two')
    }

    It 'treats an already absent app as successful and continues after another failure' {
        $apps = @(
            [pscustomobject]@{ Name = 'Gone'; GraphId = 'gone' }
            [pscustomobject]@{ Name = 'Fails'; GraphId = 'fails' }
            [pscustomobject]@{ Name = 'Works'; GraphId = 'works' }
        )
        $result = Invoke-WinTunerSupersededRemoval -Apps $apps -RemoveApp {
            param($App)
            if ($App.GraphId -eq 'gone') { throw '404 not found' }
            if ($App.GraphId -eq 'fails') { throw 'tenant denied removal' }
        }

        $result.SuccessCount | Should -Be 2
        $result.FailureCount | Should -Be 1
        $result.Results[0].ReasonCode | Should -Be 'AlreadyAbsent'
        $result.Results[1].ReasonCode | Should -Be 'RemovalFailed'
        $result.Results[2].ReasonCode | Should -Be 'Removed'
    }

    It 'blocks a removal when the Graph identifier is missing' {
        $calls = 0
        $result = Invoke-WinTunerSupersededRemoval -Apps @([pscustomobject]@{ Name = 'Unsafe'; GraphId = '' }) -RemoveApp { $calls++ }.GetNewClosure()
        $result.FailureCount | Should -Be 1
        $result.Results[0].ReasonCode | Should -Be 'MissingGraphId'
        $calls | Should -Be 0
    }
}
