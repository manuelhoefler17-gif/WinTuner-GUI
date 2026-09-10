[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path -Parent $PSScriptRoot

$syntaxFiles = @(
    Get-ChildItem -LiteralPath $repositoryRoot -Recurse -File |
        Where-Object { $_.Extension -in '.ps1', '.psm1', '.psd1' }
)

$syntaxErrors = [System.Collections.Generic.List[object]]::new()
foreach ($file in $syntaxFiles) {
    $tokens = $null
    $errors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile(
        $file.FullName,
        [ref]$tokens,
        [ref]$errors
    )

    foreach ($errorRecord in @($errors)) {
        $syntaxErrors.Add([pscustomobject]@{
            Path    = $file.FullName
            Line    = $errorRecord.Extent.StartLineNumber
            Column  = $errorRecord.Extent.StartColumnNumber
            Message = $errorRecord.Message
        })
    }
}

if ($syntaxErrors.Count -gt 0) {
    $syntaxErrors | Format-Table -AutoSize | Out-String | Write-Error
    throw "PowerShell syntax validation failed with $($syntaxErrors.Count) error(s)."
}

Write-Host "PowerShell syntax validation passed: $($syntaxFiles.Count) file(s)."

Import-Module Pester -MinimumVersion 5.7.1 -ErrorAction Stop
$configuration = New-PesterConfiguration
$configuration.Run.Path = $PSScriptRoot
$configuration.Run.PassThru = $true
$configuration.Run.Exit = $false
$configuration.Output.Verbosity = 'Detailed'

$result = Invoke-Pester -Configuration $configuration
if ($result.FailedCount -gt 0) {
    throw "Pester failed: $($result.FailedCount) failed, $($result.PassedCount) passed."
}

Write-Host "Pester passed: $($result.PassedCount) test(s)."