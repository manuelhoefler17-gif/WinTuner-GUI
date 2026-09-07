# WinTuner.Core.psm1
# Core helper functions for WinTuner GUI.

Set-StrictMode -Version Latest

function Test-IsNewerVersion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Latest,

        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Current
    )

    if ([string]::IsNullOrWhiteSpace($Latest) -or [string]::IsNullOrWhiteSpace($Current)) {
        return $false
    }

    try {
        return ([version]$Latest -gt [version]$Current)
    }
    catch {
        # Some WinGet versions contain suffixes such as "1.2.3-beta".
        # Fall back to the leading numeric version (up to four components),
        # preserving the behavior of the original monolithic script.
        $latestMatch  = [regex]::Match($Latest, '^\s*(\d+(?:\.\d+){0,3})')
        $currentMatch = [regex]::Match($Current, '^\s*(\d+(?:\.\d+){0,3})')

        if (-not $latestMatch.Success -or -not $currentMatch.Success) {
            return $false
        }

        $latestNumeric  = $latestMatch.Groups[1].Value
        $currentNumeric = $currentMatch.Groups[1].Value

        try {
            return ([version]$latestNumeric -gt [version]$currentNumeric)
        }
        catch {
            $latestParts  = @($latestNumeric.Split('.')  | ForEach-Object { [int]$_ })
            $currentParts = @($currentNumeric.Split('.') | ForEach-Object { [int]$_ })
            $length = [Math]::Max($latestParts.Count, $currentParts.Count)

            for ($index = 0; $index -lt $length; $index++) {
                $latestPart  = if ($index -lt $latestParts.Count)  { $latestParts[$index] }  else { 0 }
                $currentPart = if ($index -lt $currentParts.Count) { $currentParts[$index] } else { 0 }

                if ($latestPart -gt $currentPart) { return $true }
                if ($latestPart -lt $currentPart) { return $false }
            }

            return $false
        }
    }
}

Export-ModuleMember -Function Test-IsNewerVersion
