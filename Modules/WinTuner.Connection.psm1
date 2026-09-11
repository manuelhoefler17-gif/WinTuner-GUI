Set-StrictMode -Version Latest

function Invoke-WinTunerConnectionVerification {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][scriptblock]$GetApps,
        [ValidateRange(1, 20)][int]$MaxAttempts = 4,
        [ValidateRange(0, 30000)][int]$RetryDelayMilliseconds = 500,
        [scriptblock]$Delay = { param($Milliseconds) Start-Sleep -Milliseconds $Milliseconds },
        [scriptblock]$ReportAttempt = { param($AttemptInfo) }
    )

    $lastError = ''
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            & $ReportAttempt ([pscustomobject]@{ Attempt = $attempt; MaxAttempts = $MaxAttempts })
        } catch {}

        try {
            $null = @(& $GetApps)
            return [pscustomobject]@{
                Succeeded = $true
                Attempts = $attempt
                ErrorMessage = ''
            }
        } catch {
            $lastError = $_.Exception.Message
            if ($attempt -lt $MaxAttempts -and $RetryDelayMilliseconds -gt 0) {
                & $Delay $RetryDelayMilliseconds
            }
        }
    }

    [pscustomobject]@{
        Succeeded = $false
        Attempts = $MaxAttempts
        ErrorMessage = $lastError
    }
}

Export-ModuleMember -Function Invoke-WinTunerConnectionVerification
