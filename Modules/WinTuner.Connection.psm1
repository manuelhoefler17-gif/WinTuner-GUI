Set-StrictMode -Version Latest

function Invoke-WinTunerGraphPermissionPreflight {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][scriptblock]$GetManagedApps,
        [Parameter(Mandatory)][scriptblock]$GetDetectedApps
    )

    $checks = @(
        [pscustomobject]@{
            Name = 'Intune managed applications'
            Permission = 'DeviceManagementApps.ReadWrite.All'
            Action = $GetManagedApps
        },
        [pscustomobject]@{
            Name = 'Intune detected applications'
            Permission = 'DeviceManagementManagedDevices.Read.All'
            Action = $GetDetectedApps
        }
    )
    $completed = [System.Collections.Generic.List[string]]::new()

    foreach ($check in $checks) {
        try {
            $null = & $check.Action
            $completed.Add([string]$check.Permission)
        } catch {
            $rawMessage = [string]$_.Exception.Message
            $forbidden = $rawMessage -match '(?i)forbidden|\b403\b'
            $errorMessage = if ($forbidden) {
                "Microsoft Graph app-only preflight was denied for $($check.Name) (403 Forbidden). Required Application permission: $($check.Permission). Grant admin consent, allow time for propagation, then sign in again."
            } else {
                "Microsoft Graph app-only preflight failed while reading $($check.Name): $rawMessage"
            }

            return [pscustomobject]@{
                Succeeded = $false
                ChecksCompleted = @($completed)
                MissingPermission = if ($forbidden) { [string]$check.Permission } else { '' }
                Forbidden = $forbidden
                ErrorMessage = $errorMessage
            }
        }
    }

    return [pscustomobject]@{
        Succeeded = $true
        ChecksCompleted = @($completed)
        MissingPermission = ''
        Forbidden = $false
        ErrorMessage = ''
    }
}

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

Export-ModuleMember -Function @(
    'Invoke-WinTunerGraphPermissionPreflight'
    'Invoke-WinTunerConnectionVerification'
)
