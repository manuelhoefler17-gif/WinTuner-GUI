BeforeAll { Import-Module (Join-Path $PSScriptRoot '..\Modules\WinTuner.Retry.psm1') -Force }

Describe 'Bounded transient retries' {
    It 'retries a transient failure and reports the wait before returning success' {
        $state = @{ Calls = 0 }
        $waits = [Collections.Generic.List[int]]::new()
        $reports = [Collections.Generic.List[object]]::new()
        $result = Invoke-WinTunerRetry -Operation {
            $state.Calls++
            if ($state.Calls -eq 1) { throw 'HTTP 429 Too Many Requests' }
            'success'
        }.GetNewClosure() -Wait { param($ms) $waits.Add($ms) }.GetNewClosure() -ReportRetry { param($info) $reports.Add($info) }.GetNewClosure()
        $result | Should -Be 'success'
        $state.Calls | Should -Be 2
        ($waits | Measure-Object -Sum).Sum | Should -Be 3000
        $reports[0].Attempt | Should -Be 2
        $reports[0].StatusCode | Should -Be 429
    }

    It 'honors Retry-After and never retries earlier than requested' {
        $response = [Net.Http.HttpResponseMessage]::new([Net.HttpStatusCode]::TooManyRequests)
        $response.Headers.Add('Retry-After','7')
        $exception = [Exception]::new('Throttle')
        $exception | Add-Member -NotePropertyName Response -NotePropertyValue $response
        $failure = [Management.Automation.ErrorRecord]::new($exception,'Throttled',[Management.Automation.ErrorCategory]::LimitsExceeded,$null)
        $state = @{ Calls = 0 }
        $waits = [Collections.Generic.List[int]]::new()
        try {
            $null = Invoke-WinTunerRetry -Operation { $state.Calls++; if ($state.Calls -eq 1) { throw $failure }; 'ok' }.GetNewClosure() -Wait { param($ms) $waits.Add($ms) }.GetNewClosure()
            ($waits | Measure-Object -Sum).Sum | Should -Be 7000
        } finally { $response.Dispose() }
    }

    It 'fails without retrying authentication, validation or permanent lookup errors' -TestCases @(
        @{ Message = 'HTTP 401 Unauthorized' }, @{ Message = 'HTTP 403 Forbidden' },
        @{ Message = '404 Not Found' }, @{ Message = 'Hash mismatch' }
    ) {
        param($Message)
        $state = @{ Calls = 0 }
        $failureOperation = { $state.Calls++; throw $Message }.GetNewClosure()
        { Invoke-WinTunerRetry -Operation $failureOperation -Wait { throw 'Must not wait' } } | Should -Throw $Message
        $state.Calls | Should -Be 1
    }

    It 'limits attempts and stops during a retry wait when canceled' {
        $state = @{ Calls = 0 }
        $failureOperation = { $state.Calls++; throw '503 Service Unavailable' }.GetNewClosure()
        { Invoke-WinTunerRetry -Operation $failureOperation -Wait {} } | Should -Throw '*503*'
        $state.Calls | Should -Be 3
        $cancel = @{ Requested = $false }
        $cancelCheck = { $cancel.Requested }.GetNewClosure()
        $cancelWait = { $cancel.Requested = $true }.GetNewClosure()
        { Invoke-WinTunerRetry -Operation { throw '429 Too Many Requests' } -ShouldCancel $cancelCheck -Wait $cancelWait } | Should -Throw '*canceled*'
    }

    It 'stops when the next retry would exceed the wait budget' {
        $attempts = @{ Count = 0 }
        $operation = { $attempts.Count++; throw '429 Too Many Requests' }.GetNewClosure()
        { Invoke-WinTunerRetry -Operation $operation -MaxTotalWaitSeconds 1 -Wait { throw 'Must not wait' } } | Should -Throw '*429*'
        $attempts.Count | Should -Be 1
    }
}
