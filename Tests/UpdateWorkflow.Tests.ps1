BeforeAll {
    $repositoryRoot = Split-Path $PSScriptRoot -Parent
    $tokens = $null; $errors = $null
    $guiAst = [Management.Automation.Language.Parser]::ParseFile((Join-Path $repositoryRoot 'WinTuner_GUI.ps1'), [ref]$tokens, [ref]$errors)
    $assignment = @($guiAst.FindAll({
        param($node)
        $node -is [Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq '$updateScript'
    }, $true))
    $script:workerText = @($assignment[0].Right.FindAll({ param($node) $node -is [Management.Automation.Language.StringConstantExpressionAst] }, $true))[0].Value
    $script:repo = $repositoryRoot

    function Invoke-UpdateWorkflowFixture {
        param($Plan, $HistoryPath, [switch]$ChangedAssignments)
        $runner = [powershell]::Create()
        $fixture = @'
param($WorkerText, $Root, $PlanJson, $HistoryPath, $ChangedAssignments)
$ErrorActionPreference = 'Stop'
$global:fixtureState = @{ Builds = 0; Deploys = 0; Patches = 0; Changed = $ChangedAssignments }
function Import-Module {
    [CmdletBinding()]
    param([string]$Name, [switch]$Force)
    if ($Name -eq 'WinTuner' -or $Name -match 'WinTuner\.(Core|Winget|PackageBuild)\.psm1$') { return }
    Microsoft.PowerShell.Core\Import-Module -Name $Name -Force -Global
}
function Get-WtToken { [CmdletBinding()] param([switch]$ShowToken,[switch]$DecodeToken) 'fixture-token' }
function Invoke-RestMethod {
    [CmdletBinding()]
    param($Uri,$Method,$Headers,$Body,$ContentType,$MaximumRedirection)
    if ($Method -eq 'PATCH') { $global:fixtureState.Patches++; return @{} }
    if ([string]$Uri -match '/assignments') {
        return @{ value = @(@{
            id = 'assignment'; intent = if ($global:fixtureState.Changed) { 'required' } else { 'available' }
            target = @{ '@odata.type' = '#microsoft.graph.groupAssignmentTarget'; groupId = 'group-a' }
            settings = @{ '@odata.type' = '#microsoft.graph.win32LobAppAssignmentSettings'; notifications = 'showAll' }
        }) }
    }
    # Match the API contract: v1.0 Win32 apps do not expose displayVersion.
    if ([string]$Uri -eq 'https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/old-app') {
        return @{ id = 'old-app'; displayVersion = '1.0' }
    }
    @{ id = 'old-app' }
}
function Invoke-WinTunerPackageBuild {
    [CmdletBinding()]
    param($PackageId,$PackageFolder,$DesiredVersion,$LatestVersion,$InstalledVersion,$ReportRetry)
    $global:fixtureState.Builds++
    [pscustomobject]@{ Succeeded = $true; EffectiveVersion = '2.0' }
}
function Test-WinTunerPackageArtifact {
    [CmdletBinding()]
    param($RootPackageFolder,$PackageId,$Version)
    [pscustomobject]@{ IsValid = $true }
}
function Deploy-WtWin32App {
    [CmdletBinding()]
    param($GraphId,[switch]$KeepAssignments,$PackageId,$Version,$RootPackageFolder)
    $global:fixtureState.Deploys++
    [pscustomobject]@{ Id = 'new-app' }
}
$context = [pscustomobject]@{ Account = 'admin@contoso.com'; TenantId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'; AuthType = 'Delegated' }
[AppDomain]::CurrentDomain.SetData('WinTuner.GUI.DelegatedGraphContext', $context)
try {
    $appJson = '{"Name":"App","GraphId":"old-app","PackageId":"Vendor.App","CurrentVersion":"1.0","LatestVersion":"2.0"}'
    $output = @(& ([scriptblock]::Create($WorkerText)) $Root $appJson $HistoryPath (Join-Path $HistoryPath 'progress.json') $context.TenantId $PlanJson ([guid]::NewGuid().ToString()) $HistoryPath)
    [pscustomobject]@{ Result = ([string]$output[-1] | ConvertFrom-Json); Builds = $global:fixtureState.Builds; Deploys = $global:fixtureState.Deploys; Patches = $global:fixtureState.Patches } | ConvertTo-Json -Depth 40 -Compress
} finally { [AppDomain]::CurrentDomain.SetData('WinTuner.GUI.DelegatedGraphContext', $null) }
'@
        $planJson = if ($Plan) { $Plan | ConvertTo-Json -Depth 40 -Compress } else { '' }
        try {
            $null = $runner.AddScript($fixture).AddArgument($script:workerText).AddArgument($script:repo).AddArgument($planJson).AddArgument($HistoryPath).AddArgument([bool]$ChangedAssignments)
            $output = @($runner.Invoke())
            if ($runner.HadErrors) { throw $runner.Streams.Error[0].Exception }
            [string]$output[-1] | ConvertFrom-Json
        } finally { $runner.Dispose() }
    }
}

Describe 'Actual update worker preview and deployment phases' {
    It 'reads the preview without building or deploying and deploys only the approved plan' {
        $historyPath = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $null = New-Item -ItemType Directory -Path $historyPath
        $preview = Invoke-UpdateWorkflowFixture -HistoryPath $historyPath
        $preview.Builds | Should -Be 0
        $preview.Deploys | Should -Be 0
        $preview.Patches | Should -Be 0
        $preview.Result.Items | Should -HaveCount 1
        $deployment = Invoke-UpdateWorkflowFixture -Plan $preview.Result -HistoryPath $historyPath
        $deployment.Result.SuccessCount | Should -Be 1
        $deployment.Builds | Should -Be 1
        $deployment.Deploys | Should -Be 1
        $deployment.Patches | Should -Be 1
        @(Get-ChildItem -LiteralPath (Join-Path $historyPath 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa') -Filter '*.json') | Should -HaveCount 1
    }

    It 'blocks a stale assignment snapshot inside the actual worker before deployment' {
        $historyPath = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $null = New-Item -ItemType Directory -Path $historyPath
        $preview = Invoke-UpdateWorkflowFixture -HistoryPath $historyPath
        $deployment = Invoke-UpdateWorkflowFixture -Plan $preview.Result -HistoryPath $historyPath -ChangedAssignments
        $deployment.Result.FailureCount | Should -Be 1
        $deployment.Result.Results[0].ReasonCode | Should -Be 'PlanChanged'
        $deployment.Deploys | Should -Be 0
        $deployment.Patches | Should -Be 0
    }
}
