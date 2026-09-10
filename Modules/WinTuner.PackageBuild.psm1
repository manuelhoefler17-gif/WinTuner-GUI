Set-StrictMode -Version Latest

function New-WinTunerPackageBuildResult {
    param(
        [bool]$Succeeded,
        [AllowNull()][string]$EffectiveVersion,
        [AllowNull()][string]$ErrorMessage,
        [AllowNull()][string]$ChoiceRequired
    )

    [pscustomobject]@{
        Succeeded        = $Succeeded
        EffectiveVersion = $EffectiveVersion
        ErrorMessage     = $ErrorMessage
        ChoiceRequired   = $ChoiceRequired
    }
}

function Invoke-WinTunerPackageAttempt {
    param(
        [Parameter(Mandatory)][string]$PackageId,
        [Parameter(Mandatory)][string]$PackageFolder,
        [AllowNull()][string]$Version,
        [Parameter(Mandatory)][scriptblock]$CreatePackage
    )

    try {
        $null = & $CreatePackage $PackageId $PackageFolder $Version
        return New-WinTunerPackageBuildResult -Succeeded $true -EffectiveVersion $Version
    }
    catch {
        return New-WinTunerPackageBuildResult -Succeeded $false -ErrorMessage $_.Exception.Message
    }
}

function Invoke-WinTunerPreviousPackageAttempt {
    param(
        [Parameter(Mandatory)][string]$PackageId,
        [Parameter(Mandatory)][string]$PackageFolder,
        [AllowNull()][string]$LatestVersion,
        [AllowNull()][string]$InstalledVersion,
        [Parameter(Mandatory)][scriptblock]$CreatePackage,
        [Parameter(Mandatory)][scriptblock]$GetPreviousVersion,
        [Parameter(Mandatory)][scriptblock]$IsNewerVersion,
        [AllowNull()][string]$OriginalError
    )

    try {
        $previousVersion = [string](& $GetPreviousVersion $PackageId $LatestVersion)
    }
    catch {
        return New-WinTunerPackageBuildResult -Succeeded $false -ErrorMessage $_.Exception.Message
    }

    if ([string]::IsNullOrWhiteSpace($previousVersion)) {
        $message = if ([string]::IsNullOrWhiteSpace($OriginalError)) {
            'No previous WinGet version is available.'
        }
        else {
            $OriginalError
        }
        return New-WinTunerPackageBuildResult -Succeeded $false -ErrorMessage $message
    }

    if (
        -not [string]::IsNullOrWhiteSpace($InstalledVersion) -and
        -not [bool](& $IsNewerVersion $previousVersion $InstalledVersion)
    ) {
        $message = if ([string]::IsNullOrWhiteSpace($OriginalError)) {
            "Previous WinGet version '$previousVersion' is not newer than installed version '$InstalledVersion'."
        }
        else {
            $OriginalError
        }
        return New-WinTunerPackageBuildResult -Succeeded $false -ErrorMessage $message
    }

    return Invoke-WinTunerPackageAttempt `
        -PackageId $PackageId `
        -PackageFolder $PackageFolder `
        -Version $previousVersion `
        -CreatePackage $CreatePackage
}

function Invoke-WinTunerPackageBuild {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$PackageId,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$PackageFolder,
        [AllowNull()][string]$DesiredVersion,
        [AllowNull()][string]$LatestVersion,
        [AllowNull()][string]$InstalledVersion,
        [ValidateSet('Automatic', 'RetrySame', 'Previous')][string]$Mode = 'Automatic',
        [scriptblock]$CreatePackage = {
            param($Id, $Folder, $Version)
            if ([string]::IsNullOrWhiteSpace([string]$Version)) {
                New-WtWingetPackage -PackageId $Id -PackageFolder $Folder -ErrorAction Stop
            }
            else {
                New-WtWingetPackage -PackageId $Id -PackageFolder $Folder -Version $Version -ErrorAction Stop
            }
        },
        [scriptblock]$GetPreviousVersion = {
            param($Id, $Version)
            Get-PreviousWingetVersion -PackageId $Id -LatestVersion $Version
        },
        [scriptblock]$IsNewerVersion = {
            param($Candidate, $Installed)
            Test-IsNewerVersion -Latest $Candidate -Current $Installed
        }
    )

    $attemptVersion = if ([string]::IsNullOrWhiteSpace($DesiredVersion)) {
        $LatestVersion
    }
    else {
        $DesiredVersion
    }

    if ($Mode -eq 'Previous') {
        return Invoke-WinTunerPreviousPackageAttempt `
            -PackageId $PackageId `
            -PackageFolder $PackageFolder `
            -LatestVersion $attemptVersion `
            -InstalledVersion $InstalledVersion `
            -CreatePackage $CreatePackage `
            -GetPreviousVersion $GetPreviousVersion `
            -IsNewerVersion $IsNewerVersion
    }

    $attempt = Invoke-WinTunerPackageAttempt `
        -PackageId $PackageId `
        -PackageFolder $PackageFolder `
        -Version $attemptVersion `
        -CreatePackage $CreatePackage

    if ($attempt.Succeeded -or $Mode -eq 'RetrySame') {
        return $attempt
    }

    if ($attempt.ErrorMessage -match '404|Not Found') {
        return Invoke-WinTunerPreviousPackageAttempt `
            -PackageId $PackageId `
            -PackageFolder $PackageFolder `
            -LatestVersion $attemptVersion `
            -InstalledVersion $InstalledVersion `
            -CreatePackage $CreatePackage `
            -GetPreviousVersion $GetPreviousVersion `
            -IsNewerVersion $IsNewerVersion `
            -OriginalError $attempt.ErrorMessage
    }

    if ($attempt.ErrorMessage -match 'Hash mismatch') {
        return New-WinTunerPackageBuildResult `
            -Succeeded $false `
            -ErrorMessage $attempt.ErrorMessage `
            -ChoiceRequired 'HashMismatch'
    }

    return $attempt
}

Export-ModuleMember -Function Invoke-WinTunerPackageBuild