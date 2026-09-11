BeforeAll {
function Invoke-WinTunerIsolatedTestScript {
    param(
        [Parameter(Mandatory)][string]$Script,
        [Parameter(Mandatory)][object[]]$Arguments
    )

    $powerShell = [System.Management.Automation.PowerShell]::Create()
    try {
        $null = $powerShell.AddScript($Script)
        foreach ($argument in $Arguments) { $null = $powerShell.AddArgument($argument) }
        $asyncResult = $powerShell.BeginInvoke()
        $output = @($powerShell.EndInvoke($asyncResult))
        if ($powerShell.Streams.Error.Count -gt 0) { throw $powerShell.Streams.Error[0].Exception }
        return [string]$output[$output.Count - 1]
    } finally {
        $powerShell.Dispose()
    }
}
}

Describe 'Isolated workflow runspaces' {
    It 'completes Discovery matching across a runspace and JSON boundary' {
        $repositoryRoot = Split-Path -Parent $PSScriptRoot
        $scriptText = @'
param($RepositoryRoot)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $RepositoryRoot 'Modules\WinTuner.DiscoveryScan.psm1') -Force -ErrorAction Stop
$result = Invoke-WinTunerDiscoveryScan -ConnectGraph { $true } -GetExistingApps { @() } -ResolvePackageId { '' } -GetDetectedApps {
    [pscustomobject]@{
        Apps = @([pscustomobject]@{ displayName = 'Contoso Tool 2.0'; publisher = 'Contoso'; deviceCount = 5 })
        FromCache = $false
        PageCount = 1
        LimitReached = $false
    }
} -SearchPackages {
    [pscustomobject]@{
        Canceled = $false
        TotalQueries = 1
        CacheHits = 0
        WorkerQueries = 1
        WorkerCount = 1
        Results = @([pscustomobject]@{
            Query = 'Contoso Tool'
            Success = $true
            Results = @([pscustomobject]@{ Name = 'Contoso Tool'; PackageID = 'Contoso.Tool'; Version = '2.0' })
        })
    }
}
$result | ConvertTo-Json -Depth 7 -Compress
'@
        $result = (Invoke-WinTunerIsolatedTestScript -Script $scriptText -Arguments @($repositoryRoot)) | ConvertFrom-Json
        $result.ErrorMessage | Should -BeNullOrEmpty
        $result.UniquePackageCount | Should -Be 1
        $result.Apps[0].WingetApp.PackageID | Should -Be 'Contoso.Tool'
    }

    It 'completes non-writing deployment and removal callbacks in one isolated runspace' {
        $repositoryRoot = Split-Path -Parent $PSScriptRoot
        $scriptText = @'
param($RepositoryRoot)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $RepositoryRoot 'Modules\WinTuner.DiscoveryDeployment.psm1') -Force -ErrorAction Stop
Import-Module (Join-Path $RepositoryRoot 'Modules\WinTuner.SupersededRemoval.psm1') -Force -ErrorAction Stop
$app = [pscustomobject]@{
    DisplayName = 'Contoso Tool'
    WingetApp = [pscustomobject]@{ Name = 'Contoso Tool'; PackageID = 'Contoso.Tool'; Version = '2.0' }
}
$deployment = Invoke-WinTunerDiscoveryDeployment -Apps @($app) -RootPackageFolder $RepositoryRoot -BuildPackage {
    [pscustomobject]@{ Succeeded = $true; EffectiveVersion = '2.0' }
} -ValidateArtifact {
    [pscustomobject]@{ IsValid = $true }
} -DeployPackage {
    param($Id, $Version, $Root)
    [pscustomobject]@{ WouldDeploy = "$Id|$Version|$Root" }
}
$removal = Invoke-WinTunerSupersededRemoval -Apps @([pscustomobject]@{ Name = 'Old Tool'; GraphId = 'graph-old' }) -RemoveApp {
    param($App)
    [pscustomobject]@{ WouldRemove = $App.GraphId }
}
[pscustomobject]@{ Deployment = $deployment; Removal = $removal } | ConvertTo-Json -Depth 7 -Compress
'@
        $result = (Invoke-WinTunerIsolatedTestScript -Script $scriptText -Arguments @($repositoryRoot)) | ConvertFrom-Json
        $result.Deployment.SuccessCount | Should -Be 1
        $result.Removal.SuccessCount | Should -Be 1
    }

    It 'retries connection verification inside an isolated runspace' {
        $repositoryRoot = Split-Path -Parent $PSScriptRoot
        $scriptText = @'
param($RepositoryRoot)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $RepositoryRoot 'Modules\WinTuner.Connection.psm1') -Force -ErrorAction Stop
$state = [pscustomobject]@{ Attempts = 0 }
$result = Invoke-WinTunerConnectionVerification -GetApps {
    $state.Attempts++
    if ($state.Attempts -eq 1) { throw 'transient' }
    @()
}.GetNewClosure() -MaxAttempts 3 -RetryDelayMilliseconds 0
$result | ConvertTo-Json -Depth 4 -Compress
'@
        $result = (Invoke-WinTunerIsolatedTestScript -Script $scriptText -Arguments @($repositoryRoot)) | ConvertFrom-Json
        $result.Succeeded | Should -BeTrue
        $result.Attempts | Should -Be 2
    }
}
