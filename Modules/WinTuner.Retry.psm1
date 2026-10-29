Set-StrictMode -Version Latest

function Get-WinTunerRetryPolicy {
    param([Parameter(Mandatory)][System.Management.Automation.ErrorRecord]$ErrorRecord)

    $status = 0
    $retryAfter = $null
    $exception = $ErrorRecord.Exception
    while ($exception) {
        $responseProperty = $exception.PSObject.Properties['Response']
        $response = if ($responseProperty) { $responseProperty.Value } else { $null }
        $statusProperty = $exception.PSObject.Properties['StatusCode']
        if ($statusProperty -and $statusProperty.Value) { $status = [int]$statusProperty.Value }
        if ($response) {
            $statusProperty = $response.PSObject.Properties['StatusCode']
            if ($statusProperty -and $statusProperty.Value) { $status = [int]$statusProperty.Value }
            try { $retryAfter = [string](@($response.Headers.GetValues('Retry-After'))[0]) } catch {}
        }
        if ($status -gt 0) { break }
        $exception = $exception.InnerException
    }
    if ($status -eq 0 -and $ErrorRecord.Exception.Message -match '(?<!\d)(408|429|500|502|503|504)(?!\d)') {
        $status = [int]$Matches[1]
    }
    $isTimeout = $ErrorRecord.Exception -is [TimeoutException] -or $ErrorRecord.Exception.InnerException -is [TimeoutException]
    $seconds = $null
    if (-not [string]::IsNullOrWhiteSpace($retryAfter)) {
        $numericDelay = 0.0
        $retryDate = [datetimeoffset]::MinValue
        if ([double]::TryParse($retryAfter, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$numericDelay)) {
            $seconds = [math]::Max(0, [math]::Ceiling($numericDelay))
        } elseif ([datetimeoffset]::TryParse($retryAfter, [ref]$retryDate)) {
            $seconds = [math]::Max(0, [math]::Ceiling(($retryDate - [datetimeoffset]::UtcNow).TotalSeconds))
        }
    }
    [pscustomobject]@{ CanRetry = ($status -in @(408,429,500,502,503,504) -or $isTimeout); StatusCode = $status; RetryAfterSeconds = $seconds }
}

function Invoke-WinTunerRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][scriptblock]$Operation,
        [ValidateRange(1,5)][int]$MaxAttempts = 3,
        [ValidateRange(0,300)][int]$MaxTotalWaitSeconds = 120,
        [scriptblock]$ReportRetry = { param($RetryInfo) },
        [scriptblock]$ShouldCancel = { $false },
        [scriptblock]$Wait = { param($Milliseconds) Start-Sleep -Milliseconds $Milliseconds }
    )

    $totalWait = 0
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        if (& $ShouldCancel) { throw [OperationCanceledException]::new('Operation canceled while waiting to retry.') }
        try { return & $Operation }
        catch {
            $failure = $_
            $policy = Get-WinTunerRetryPolicy -ErrorRecord $failure
            if (-not $policy.CanRetry -or $attempt -ge $MaxAttempts) { throw }
            $delay = if ($null -ne $policy.RetryAfterSeconds) { [double]$policy.RetryAfterSeconds } else { [math]::Pow(2, $attempt) + 1 }
            # Never shorten a server's Retry-After. Fail when it exceeds this operation's budget.
            if ($delay -gt ($MaxTotalWaitSeconds - $totalWait)) { throw }
            try {
                $null = & $ReportRetry ([pscustomobject]@{
                    Attempt = $attempt + 1; MaxAttempts = $MaxAttempts
                    DelaySeconds = $delay; StatusCode = $policy.StatusCode
                })
            } catch {}
            $remainingMs = [int]($delay * 1000)
            while ($remainingMs -gt 0) {
                if (& $ShouldCancel) { throw [OperationCanceledException]::new('Operation canceled while waiting to retry.') }
                $slice = [math]::Min(1000, $remainingMs)
                $null = & $Wait $slice
                $remainingMs -= $slice
            }
            $totalWait += $delay
        }
    }
}

Export-ModuleMember -Function Invoke-WinTunerRetry
