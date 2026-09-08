Set-StrictMode -Version Latest

function New-WinTunerDefaultSettings {
    [CmdletBinding()]
    param()

    return @{
        RememberMe         = $false
        LastUser           = ""
        RecentUsers        = @()
        MaxRecentUsers     = 3
        WingetOverrides    = @{}
        DefaultPackagePath = "C:\Temp"
        AutoCheckUpdates   = $false
    }
}

function Import-WinTunerSettings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $settings = New-WinTunerDefaultSettings

    if (-not (Test-Path -LiteralPath $Path)) {
        return $settings
    }

    try {
        $o = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop |
             ConvertFrom-Json -ErrorAction Stop

        if (-not $o) {
            return $settings
        }

        if ($o.PSObject.Properties['RememberMe']) {
            $settings.RememberMe = [bool]$o.RememberMe
        }

        if ($o.PSObject.Properties['LastUser']) {
            $settings.LastUser = [string]$o.LastUser
        }

        if ($o.PSObject.Properties['RecentUsers']) {
            $settings.RecentUsers = @([string[]]$o.RecentUsers)
        }

        if (
            $o.PSObject.Properties['MaxRecentUsers'] -and
            [int]$o.MaxRecentUsers -gt 0
        ) {
            $settings.MaxRecentUsers = [int]$o.MaxRecentUsers
        }

        if ($o.PSObject.Properties['DefaultPackagePath']) {
            $settings.DefaultPackagePath = [string]$o.DefaultPackagePath
        }

        if ($o.PSObject.Properties['AutoCheckUpdates']) {
            $settings.AutoCheckUpdates = [bool]$o.AutoCheckUpdates
        }

        if ($o.PSObject.Properties['WingetOverrides']) {
            $overrides = @{}

            if ($null -ne $o.WingetOverrides) {
                foreach ($p in $o.WingetOverrides.PSObject.Properties) {
                    $overrides[$p.Name] = [string]$p.Value
                }
            }

            $settings.WingetOverrides = $overrides
        }

        return $settings
    }
    catch {
        Write-Verbose "Failed to load settings from '$Path': $($_.Exception.Message)"
        return $settings
    }
}

function Export-WinTunerSettings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Settings,

        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    try {
        $directory = Split-Path -Parent $Path

        if (
            -not [string]::IsNullOrWhiteSpace($directory) -and
            -not (Test-Path -LiteralPath $directory)
        ) {
            New-Item -ItemType Directory -Path $directory -Force -ErrorAction Stop |
                Out-Null
        }

        $Settings |
            ConvertTo-Json -Compress -Depth 10 |
            Set-Content -LiteralPath $Path -Encoding utf8 -ErrorAction Stop

        return $true
    }
    catch {
        Write-Verbose "Failed to save settings to '$Path': $($_.Exception.Message)"
        return $false
    }
}

Export-ModuleMember -Function `
    New-WinTunerDefaultSettings, `
    Import-WinTunerSettings, `
    Export-WinTunerSettings
