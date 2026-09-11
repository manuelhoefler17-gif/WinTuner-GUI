Set-StrictMode -Version Latest

function Get-WinTunerSupersededProperty {
    param([Parameter(Mandatory)][object]$InputObject, [Parameter(Mandatory)][string]$Name)
    $property = $InputObject.PSObject.Properties[$Name]
    if ($property) { return [string]$property.Value }
    return ''
}

function Invoke-WinTunerSupersededRemoval {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateNotNull()][object[]]$Apps,
        [scriptblock]$RemoveApp = { param($App) Remove-WtWin32App -GraphId ([string]$App.GraphId) -ErrorAction Stop },
        [scriptblock]$ReportProgress = { param($ProgressInfo) }
    )

    $items = @($Apps)
    $results = [System.Collections.Generic.List[object]]::new()
    $processed = 0

    foreach ($app in $items) {
        $processed++
        $name = Get-WinTunerSupersededProperty -InputObject $app -Name 'Name'
        $graphId = Get-WinTunerSupersededProperty -InputObject $app -Name 'GraphId'

        try {
            & $ReportProgress ([pscustomobject]@{
                Stage = 'Removing'
                Processed = $processed
                Total = $items.Count
                AppName = $name
            })
        } catch {}

        if ([string]::IsNullOrWhiteSpace($graphId)) {
            $results.Add([pscustomobject]@{
                Succeeded = $false
                AlreadyAbsent = $false
                Name = $name
                GraphId = $graphId
                ReasonCode = 'MissingGraphId'
                Message = "Cannot remove '$name' because its Graph identifier is missing."
            })
            continue
        }

        try {
            $null = & $RemoveApp $app
            $results.Add([pscustomobject]@{
                Succeeded = $true
                AlreadyAbsent = $false
                Name = $name
                GraphId = $graphId
                ReasonCode = 'Removed'
                Message = "Removed '$name'."
            })
        } catch {
            $message = $_.Exception.Message
            if ($message -match '(?i)not found|404') {
                $results.Add([pscustomobject]@{
                    Succeeded = $true
                    AlreadyAbsent = $true
                    Name = $name
                    GraphId = $graphId
                    ReasonCode = 'AlreadyAbsent'
                    Message = "'$name' was already absent from Intune."
                })
            } else {
                $results.Add([pscustomobject]@{
                    Succeeded = $false
                    AlreadyAbsent = $false
                    Name = $name
                    GraphId = $graphId
                    ReasonCode = 'RemovalFailed'
                    Message = "Removal failed for '$name': $message"
                })
            }
        }
    }

    $successCount = @($results | Where-Object Succeeded).Count
    [pscustomobject]@{
        SuccessCount = $successCount
        FailureCount = $results.Count - $successCount
        Results = @($results)
    }
}

Export-ModuleMember -Function Invoke-WinTunerSupersededRemoval
