Set-StrictMode -Version Latest

$script:LogBasePath = $null
$script:OutputBox   = $null
$script:MaxLogSize  = 2MB

function Initialize-WinTunerLogging {
    [CmdletBinding()]
    param(
        [string]$BasePath,
        [object]$OutputBox
    )

    if ([string]::IsNullOrWhiteSpace($BasePath)) {
        $BasePath = [Environment]::GetFolderPath('LocalApplicationData')
    }

    $script:LogBasePath = $BasePath
    $script:OutputBox   = $OutputBox
}

function Get-WinTunerLogPath {
    [CmdletBinding()]
    param()

    $base = $script:LogBasePath

    if ([string]::IsNullOrWhiteSpace($base)) {
        $base = [Environment]::GetFolderPath('LocalApplicationData')
    }

    return (Join-Path $base 'WinTuner_GUI.log')
}

function Write-Log {
    [CmdletBinding()]
    param(
        [string]$Message
    )

    if ([string]::IsNullOrWhiteSpace($Message)) {
        return
    }

    try {
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $logLine   = "$timestamp - $Message"
        $logPath   = Get-WinTunerLogPath
        $base      = Split-Path -Parent $logPath

        if (-not (Test-Path -LiteralPath $base)) {
            New-Item -ItemType Directory -Path $base -Force -ErrorAction Stop |
                Out-Null
        }

        try {
            if (Test-Path -LiteralPath $logPath) {
                $logFile = Get-Item -LiteralPath $logPath -ErrorAction Stop

                if ($logFile.Length -gt $script:MaxLogSize) {
                    $oldLogPath = Join-Path $base 'WinTuner_GUI_old.log'

                    Move-Item `
                        -LiteralPath $logPath `
                        -Destination $oldLogPath `
                        -Force `
                        -ErrorAction SilentlyContinue
                }
            }

            Add-Content `
                -LiteralPath $logPath `
                -Value $logLine `
                -Encoding utf8 `
                -ErrorAction SilentlyContinue
        }
        catch {
        }

        $outputBox = $script:OutputBox

        if ($outputBox -and -not $outputBox.IsDisposed) {
            try {
                if ($outputBox.InvokeRequired) {
                    $line = $logLine
                    $outputBox.Invoke(
                        [Action]{
                            $outputBox.AppendText("$line`r`n")
                        }
                    )
                }
                else {
                    $outputBox.AppendText("$logLine`r`n")
                }
            }
            catch {
            }
        }
    }
    catch {
    }
}

function Write-LogSafe {
    [CmdletBinding()]
    param(
        [string]$Message
    )

    if ([string]::IsNullOrWhiteSpace($Message)) {
        return
    }

    try {
        Write-Log -Message $Message
    }
    catch {
        try {
            $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            $logPath   = Get-WinTunerLogPath

            Add-Content `
                -LiteralPath $logPath `
                -Value "$timestamp - $Message" `
                -Encoding utf8 `
                -ErrorAction SilentlyContinue
        }
        catch {
        }
    }
}

Export-ModuleMember -Function `
    Initialize-WinTunerLogging, `
    Get-WinTunerLogPath, `
    Write-Log, `
    Write-LogSafe
