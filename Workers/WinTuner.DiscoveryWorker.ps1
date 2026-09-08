[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$InputPath,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

function Write-WorkerOutput {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Results
    )

    $tempPath = "$OutputPath.tmp"

    try {
        @($Results) |
            ConvertTo-Json -Depth 8 |
            Set-Content `
                -Path $tempPath `
                -Encoding utf8 `
                -NoNewline

        Move-Item `
            -Path $tempPath `
            -Destination $OutputPath `
            -Force
    }
    finally {
        Remove-Item `
            -Path $tempPath `
            -Force `
            -ErrorAction SilentlyContinue
    }
}

if (-not (Test-Path $InputPath)) {
    throw "Input file not found: $InputPath"
}

$modulePath = (
    Get-Module WinTuner -ListAvailable |
        Sort-Object Version -Descending |
        Select-Object -First 1
).Path

if (-not $modulePath) {
    throw 'WinTuner module not found.'
}

Import-Module `
    $modulePath `
    -Force `
    -ErrorAction Stop

$inputData = @(
    Get-Content `
        -Path $InputPath `
        -Raw `
        -Encoding utf8 |
        ConvertFrom-Json
)

$output = [System.Collections.Generic.List[object]]::new()

foreach ($item in $inputData) {
    $query = [string]$item

    if ([string]::IsNullOrWhiteSpace($query)) {
        continue
    }

    try {
        $rawResults = @(
            Search-WtWinGetPackage `
                -SearchQuery $query `
                -ErrorAction SilentlyContinue `
                2>$null 3>$null 4>$null 5>$null 6>$null
        )

        $cleanResults = @(
            foreach ($result in $rawResults) {
                if ($null -eq $result) {
                    continue
                }

                $packageId = [string]$result.PackageID

                if ([string]::IsNullOrWhiteSpace($packageId)) {
                    continue
                }

                [pscustomobject]@{
                    Name      = [string]$result.Name
                    PackageID = $packageId
                }
            }
        )

        $cleanResults = @(
            $cleanResults |
                Sort-Object PackageID -Unique
        )

        $output.Add(
            [pscustomobject]@{
                Query   = $query
                Success = $true
                Results = $cleanResults
                Error   = $null
            }
        )
    }
    catch {
        $output.Add(
            [pscustomobject]@{
                Query   = $query
                Success = $false
                Results = @()
                Error   = $_.Exception.Message
            }
        )
    }
}

Write-WorkerOutput -Results @($output)