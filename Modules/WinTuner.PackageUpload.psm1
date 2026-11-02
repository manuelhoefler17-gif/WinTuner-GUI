Set-StrictMode -Version Latest

function Test-WinTunerUploadArtifact {
    param($Root, $Id, $Version)
    Test-WinTunerPackageArtifact -RootPackageFolder $Root -PackageId $Id -Version $Version
}

function Publish-WinTunerUploadPackage {
    param($Id, $Version, $Root)
    Deploy-WtWin32App -PackageId $Id -Version $Version -RootPackageFolder $Root -ErrorAction Stop
}

function New-WinTunerPackageUploadResult {
    param(
        [bool]$Succeeded,
        [AllowNull()][string]$ErrorMessage,
        [AllowNull()][string]$ReasonCode,
        [AllowNull()][string]$IntuneWinPath
    )

    [pscustomobject]@{
        Succeeded     = $Succeeded
        ErrorMessage  = $ErrorMessage
        ReasonCode    = $ReasonCode
        IntuneWinPath = $IntuneWinPath
    }
}

function Invoke-WinTunerPackageUpload {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$PackageId,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Version,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$RootPackageFolder
    )

    try {
        $validation = Test-WinTunerUploadArtifact $RootPackageFolder $PackageId $Version
    }
    catch {
        return New-WinTunerPackageUploadResult `
            -Succeeded $false `
            -ErrorMessage $_.Exception.Message `
            -ReasonCode 'ValidationError'
    }

    if (-not $validation -or -not [bool]$validation.IsValid) {
        $reason = if ($validation -and -not [string]::IsNullOrWhiteSpace([string]$validation.Reason)) {
            [string]$validation.Reason
        }
        else {
            'Package artifact validation failed.'
        }
        $reasonCode = if ($validation -and -not [string]::IsNullOrWhiteSpace([string]$validation.ReasonCode)) {
            [string]$validation.ReasonCode
        }
        else {
            'InvalidArtifact'
        }

        return New-WinTunerPackageUploadResult `
            -Succeeded $false `
            -ErrorMessage $reason `
            -ReasonCode $reasonCode
    }

    try {
        $null = Publish-WinTunerUploadPackage $PackageId $Version $RootPackageFolder
    }
    catch {
        return New-WinTunerPackageUploadResult `
            -Succeeded $false `
            -ErrorMessage $_.Exception.Message `
            -ReasonCode 'DeploymentFailed' `
            -IntuneWinPath ([string]$validation.IntuneWinPath)
    }

    return New-WinTunerPackageUploadResult `
        -Succeeded $true `
        -ReasonCode 'Uploaded' `
        -IntuneWinPath ([string]$validation.IntuneWinPath)
}

Export-ModuleMember -Function Invoke-WinTunerPackageUpload