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

function Test-WinTunerPackageRoot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$RootPackageFolder
    )

    $result = [ordered]@{
        IsValid    = $false
        ReasonCode = 'Unknown'
        Reason     = 'Package root validation did not complete.'
        FullPath   = $null
    }

    if ([string]::IsNullOrWhiteSpace($RootPackageFolder)) {
        $result['ReasonCode'] = 'InvalidRoot'
        $result['Reason'] = 'Package root folder is empty.'
        return [pscustomobject]$result
    }

    try {
        $fullPath = [System.IO.Path]::GetFullPath($RootPackageFolder.Trim())
    }
    catch {
        $result['ReasonCode'] = 'InvalidRoot'
        $result['Reason'] = "Package root folder is invalid: $($_.Exception.Message)"
        return [pscustomobject]$result
    }

    $result['FullPath'] = $fullPath

    if (Test-Path -LiteralPath $fullPath -PathType Leaf) {
        $result['ReasonCode'] = 'RootIsFile'
        $result['Reason'] = 'Package root points to a file instead of a directory.'
        return [pscustomobject]$result
    }

    $normalizedPath = $fullPath.TrimEnd(
        [System.IO.Path]::DirectorySeparatorChar,
        [System.IO.Path]::AltDirectorySeparatorChar
    )
    $volumeRoot = [System.IO.Path]::GetPathRoot($fullPath).TrimEnd(
        [System.IO.Path]::DirectorySeparatorChar,
        [System.IO.Path]::AltDirectorySeparatorChar
    )

    if ($normalizedPath.Equals($volumeRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        $result['ReasonCode'] = 'ProtectedRoot'
        $result['Reason'] = 'The selected package root cannot be a drive root.'
        return [pscustomobject]$result
    }

    $protectedPaths = @(
        [Environment]::GetFolderPath('Windows'),
        [Environment]::GetFolderPath('System'),
        $env:ProgramFiles,
        ${env:ProgramFiles(x86)}
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }

    foreach ($protectedPath in $protectedPaths) {
        try {
            $normalizedProtectedPath = [System.IO.Path]::GetFullPath([string]$protectedPath).TrimEnd(
                [System.IO.Path]::DirectorySeparatorChar,
                [System.IO.Path]::AltDirectorySeparatorChar
            )
        }
        catch {
            continue
        }

        if (
            $normalizedPath.Equals($normalizedProtectedPath, [System.StringComparison]::OrdinalIgnoreCase) -or
            $normalizedPath.StartsWith(
                $normalizedProtectedPath + [System.IO.Path]::DirectorySeparatorChar,
                [System.StringComparison]::OrdinalIgnoreCase
            )
        ) {
            $result['ReasonCode'] = 'ProtectedRoot'
            $result['Reason'] = 'The selected package root is a protected system directory.'
            return [pscustomobject]$result
        }
    }

    $result['IsValid'] = $true
    $result['ReasonCode'] = 'Valid'
    $result['Reason'] = 'Package root is valid.'
    return [pscustomobject]$result
}

function Test-WinTunerPackageArtifact {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$RootPackageFolder,

        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$PackageId,

        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Version
    )

    $result = [ordered]@{
        IsValid      = $false
        ReasonCode   = 'Unknown'
        Reason       = 'Package artifact validation did not complete.'
        RootPath     = $null
        PackagePath  = $null
        MetadataPath = $null
        IntuneWinPath = $null
        FileName     = $null
        DisplayVersion = $null
    }


    foreach ($component in @(
        @{ Name = 'PackageId'; Value = $PackageId },
        @{ Name = 'Version'; Value = $Version }
    )) {
        $value = [string]$component.Value

        if (
            [string]::IsNullOrWhiteSpace($value) -or
            [System.IO.Path]::IsPathRooted($value) -or
            $value -in '.', '..' -or
            $value.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars()) -ge 0
        ) {
            $result['ReasonCode'] = "Invalid$($component.Name)"
            $result['Reason'] = "$($component.Name) is not a safe package path component."
            return [pscustomobject]$result
        }
    }

    $rootValidation = Test-WinTunerPackageRoot -RootPackageFolder $RootPackageFolder
    $result['RootPath'] = $rootValidation.FullPath

    if (-not $rootValidation.IsValid) {
        $result['ReasonCode'] = $rootValidation.ReasonCode
        $result['Reason'] = $rootValidation.Reason
        return [pscustomobject]$result
    }

    $rootPath = $rootValidation.FullPath

    if (-not (Test-Path -LiteralPath $rootPath -PathType Container)) {
        $result['ReasonCode'] = 'RootNotFound'
        $result['Reason'] = 'Package root folder does not exist.'
        return [pscustomobject]$result
    }

    try {
        $packagePath = [System.IO.Path]::GetFullPath(
            [System.IO.Path]::Combine($rootPath, $PackageId, $Version)
        )
    }
    catch {
        $result['ReasonCode'] = 'InvalidPackagePath'
        $result['Reason'] = "Package directory path is invalid: $($_.Exception.Message)"
        return [pscustomobject]$result
    }

    $rootPrefix = $rootPath.TrimEnd(
        [System.IO.Path]::DirectorySeparatorChar,
        [System.IO.Path]::AltDirectorySeparatorChar
    ) + [System.IO.Path]::DirectorySeparatorChar

    if (-not $packagePath.StartsWith($rootPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        $result['ReasonCode'] = 'InvalidPackagePath'
        $result['Reason'] = 'Package directory resolves outside the selected package root.'
        return [pscustomobject]$result
    }

    $result['PackagePath'] = $packagePath

    if (-not (Test-Path -LiteralPath $packagePath -PathType Container)) {
        $result['ReasonCode'] = 'PackageNotFound'
        $result['Reason'] = 'Package directory does not exist for the selected package and version.'
        return [pscustomobject]$result
    }

    $metadataPath = Join-Path -Path $packagePath -ChildPath 'win32LobApp.json'
    $result['MetadataPath'] = $metadataPath

    if (-not (Test-Path -LiteralPath $metadataPath -PathType Leaf)) {
        $result['ReasonCode'] = 'MetadataNotFound'
        $result['Reason'] = 'win32LobApp.json was not found in the package directory.'
        return [pscustomobject]$result
    }

    try {
        $metadata = Get-Content -LiteralPath $metadataPath -Raw -ErrorAction Stop |
            ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        $result['ReasonCode'] = 'InvalidMetadata'
        $result['Reason'] = "win32LobApp.json could not be read: $($_.Exception.Message)"
        return [pscustomobject]$result
    }

    $fileName = ''
    if ($null -ne $metadata) {
        $fileNameProperty = $metadata.PSObject.Properties['fileName']
        if ($null -ne $fileNameProperty) {
            $fileName = ([string]$fileNameProperty.Value).Trim()
        }
    }

    if ([string]::IsNullOrWhiteSpace($fileName)) {
        $result['ReasonCode'] = 'MissingFileName'
        $result['Reason'] = 'win32LobApp.json does not contain a package filename.'
        return [pscustomobject]$result
    }

    $displayVersion = ''
    if ($null -ne $metadata) {
        $displayVersionProperty = $metadata.PSObject.Properties['displayVersion']
        if ($null -ne $displayVersionProperty) {
            $displayVersion = ([string]$displayVersionProperty.Value).Trim()
        }
    }

    $result['DisplayVersion'] = $displayVersion

    if ([string]::IsNullOrWhiteSpace($displayVersion)) {
        $result['ReasonCode'] = 'MissingDisplayVersion'
        $result['Reason'] = 'win32LobApp.json does not contain a display version.'
        return [pscustomobject]$result
    }

    if (-not $displayVersion.Equals($Version, [System.StringComparison]::OrdinalIgnoreCase)) {
        $result['ReasonCode'] = 'VersionMismatch'
        $result['Reason'] = "The package metadata version '$displayVersion' does not match selected version '$Version'."
        return [pscustomobject]$result
    }

    if (
        [System.IO.Path]::IsPathRooted($fileName) -or
        $fileName -ne [System.IO.Path]::GetFileName($fileName) -or
        $fileName.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars()) -ge 0
    ) {
        $result['ReasonCode'] = 'InvalidFileName'
        $result['Reason'] = 'The package filename in win32LobApp.json is not a safe file name.'
        return [pscustomobject]$result
    }

    if (-not $fileName.EndsWith('.intunewin', [System.StringComparison]::OrdinalIgnoreCase)) {
        $result['ReasonCode'] = 'InvalidFileType'
        $result['Reason'] = 'The package filename in win32LobApp.json is not an .intunewin file.'
        return [pscustomobject]$result
    }

    $intuneWinPath = Join-Path -Path $packagePath -ChildPath $fileName
    $result['FileName'] = $fileName
    $result['IntuneWinPath'] = $intuneWinPath

    if (-not (Test-Path -LiteralPath $intuneWinPath -PathType Leaf)) {
        $result['ReasonCode'] = 'IntuneWinNotFound'
        $result['Reason'] = 'The exact .intunewin file referenced by win32LobApp.json was not found.'
        return [pscustomobject]$result
    }

    try {
        $intuneWinFile = Get-Item -LiteralPath $intuneWinPath -ErrorAction Stop
    }
    catch {
        $result['ReasonCode'] = 'IntuneWinNotReadable'
        $result['Reason'] = "The referenced .intunewin file could not be read: $($_.Exception.Message)"
        return [pscustomobject]$result
    }

    if ($intuneWinFile.Length -le 0) {
        $result['ReasonCode'] = 'IntuneWinEmpty'
        $result['Reason'] = 'The referenced .intunewin file is empty.'
        return [pscustomobject]$result
    }

    if (($intuneWinFile.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
        $result['ReasonCode'] = 'IntuneWinReparsePoint'
        $result['Reason'] = 'The referenced .intunewin file is a reparse point and cannot be trusted.'
        return [pscustomobject]$result
    }

    $result['IsValid'] = $true
    $result['ReasonCode'] = 'Valid'
    $result['Reason'] = 'Package metadata and the referenced .intunewin file are valid.'
    return [pscustomobject]$result
}

function Get-WinTunerUpdateActionState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [bool]$Connected,

        [Parameter(Mandatory)]
        [bool]$IsBusy,

        [Parameter(Mandatory)]
        [ValidateRange(0, [int]::MaxValue)]
        [int]$CandidateCount,

        [Parameter(Mandatory)]
        [ValidateRange(0, [int]::MaxValue)]
        [int]$CheckedCount
    )

    $effectiveCheckedCount = [Math]::Min($CheckedCount, $CandidateCount)
    $canInteract = $Connected -and -not $IsBusy

    return [pscustomobject]@{
        CanSearch         = $canInteract
        CanCheckAll       = $canInteract -and $CandidateCount -gt 0 -and $effectiveCheckedCount -lt $CandidateCount
        CanUncheckAll     = $canInteract -and $effectiveCheckedCount -gt 0
        CanUpdateSelected = $canInteract -and $effectiveCheckedCount -gt 0
        CanUpdateAll      = $canInteract -and $CandidateCount -gt 0
        CanLogout         = $canInteract
    }
}

function Get-WinTunerDiscoveryActionState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [bool]$Connected,

        [Parameter(Mandatory)]
        [bool]$IsScanning,

        [Parameter(Mandatory)]
        [bool]$CancelRequested,

        [Parameter(Mandatory)]
        [bool]$IsDeploying,

        [Parameter(Mandatory)]
        [ValidateRange(0, [int]::MaxValue)]
        [int]$ResultCount,

        [Parameter(Mandatory)]
        [ValidateRange(0, [int]::MaxValue)]
        [int]$CheckedCount
    )

    $effectiveCheckedCount = [Math]::Min($CheckedCount, $ResultCount)
    $isIdle = $Connected -and -not $IsScanning -and -not $IsDeploying

    return [pscustomobject]@{
        CanScan       = $Connected -and -not $IsDeploying -and -not $CancelRequested
        CanDeploy     = $isIdle -and $effectiveCheckedCount -gt 0
        CanExport     = $isIdle -and $ResultCount -gt 0
        CanCheckAll   = $isIdle -and $ResultCount -gt 0 -and $effectiveCheckedCount -lt $ResultCount
        CanUncheckAll = $isIdle -and $effectiveCheckedCount -gt 0
    }
}

Export-ModuleMember -Function Test-IsNewerVersion, Test-WinTunerPackageRoot, Test-WinTunerPackageArtifact, Get-WinTunerUpdateActionState, Get-WinTunerDiscoveryActionState
