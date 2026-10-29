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
    Import-Module (Join-Path $repositoryRoot 'Modules\WinTuner.UpdateRules.psm1') -Force

    function Invoke-UpdateWorkflowFixture {
        param($Plan, $HistoryPath, [switch]$ChangedAssignments, [string]$Mode = 'Checked', [switch]$ChangeRulesDuringBuild)
        $runner = [powershell]::Create()
        $fixture = @'
param($WorkerText, $Root, $PlanJson, $HistoryPath, $ChangedAssignments, $UpdateMode, $ChangeRulesDuringBuild)
$ErrorActionPreference = 'Stop'
$global:fixtureState = @{ Builds = 0; Deploys = 0; Patches = 0; Changed = $ChangedAssignments; ChangeRules = $ChangeRulesDuringBuild; RulesPath = (Join-Path $HistoryPath 'rules') }
function Import-Module {
    [CmdletBinding()]
    param([string]$Name, [switch]$Force)
    if ($Name -eq 'WinTuner' -or $Name -match 'WinTuner\.(Core|Winget|PackageBuild)\.psm1$') { return }
    Microsoft.PowerShell.Core\Import-Module -Name $Name -Force -Global
}
function Get-WtToken { [CmdletBinding()] param([switch]$ShowToken,[switch]$DecodeToken) 'fixture-token' }
function Test-IsNewerVersion { param($Latest,$Current) [version]$Latest -gt [version]$Current }
function Get-WingetVersions { param($PackageId) @('2.0','1.5','1.0') }
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
    if ($global:fixtureState.ChangeRules) {
        Set-WinTunerUpdateRule -TenantId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -PackageId 'Vendor.App' -ExcludeFromAll $true -BasePath $global:fixtureState.RulesPath
    }
    [pscustomobject]@{ Succeeded = $true; EffectiveVersion = $DesiredVersion }
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
    $output = @(& ([scriptblock]::Create($WorkerText)) $Root $appJson $HistoryPath (Join-Path $HistoryPath 'progress.json') $context.TenantId $PlanJson ([guid]::NewGuid().ToString()) $HistoryPath $UpdateMode $global:fixtureState.RulesPath)
    [pscustomobject]@{ Result = ([string]$output[-1] | ConvertFrom-Json); Builds = $global:fixtureState.Builds; Deploys = $global:fixtureState.Deploys; Patches = $global:fixtureState.Patches } | ConvertTo-Json -Depth 40 -Compress
} finally { [AppDomain]::CurrentDomain.SetData('WinTuner.GUI.DelegatedGraphContext', $null) }
'@
        $planJson = if ($Plan) { $Plan | ConvertTo-Json -Depth 40 -Compress } else { '' }
        try {
            $null = $runner.AddScript($fixture).AddArgument($script:workerText).AddArgument($script:repo).AddArgument($planJson).AddArgument($HistoryPath).AddArgument([bool]$ChangedAssignments).AddArgument($Mode).AddArgument([bool]$ChangeRulesDuringBuild)
            $output = @($runner.Invoke())
            if ($runner.HadErrors) { throw $runner.Streams.Error[0].Exception }
            [string]$output[-1] | ConvertFrom-Json
        } finally { $runner.Dispose() }
    }
}

Describe 'Actual update worker preview and deployment phases' {
    It 'excludes a package from All while allowing a deliberately checked update' {
        $historyPath = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $null = New-Item -ItemType Directory -Path $historyPath
        Set-WinTunerUpdateRule -TenantId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -PackageId 'Vendor.App' -ExcludeFromAll $true -BasePath (Join-Path $historyPath 'rules')
        { Invoke-UpdateWorkflowFixture -HistoryPath $historyPath -Mode All } | Should -Throw '*No apps are eligible*'
        $preview = Invoke-UpdateWorkflowFixture -HistoryPath $historyPath -Mode Checked
        $preview.Result.Items[0].RuleDescription | Should -Match 'Excluded from Update All'
        $deployment = Invoke-UpdateWorkflowFixture -HistoryPath $historyPath -Plan $preview.Result -Mode Checked
        $deployment.Result.SuccessCount | Should -Be 1
        $deployment.Deploys | Should -Be 1
    }

    It 'builds and deploys the exact pinned version instead of the scanned latest version' {
        $historyPath = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $null = New-Item -ItemType Directory -Path $historyPath
        Set-WinTunerUpdateRule -TenantId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -PackageId 'Vendor.App' -PinnedVersion '1.5' -BasePath (Join-Path $historyPath 'rules')
        $preview = Invoke-UpdateWorkflowFixture -HistoryPath $historyPath
        $preview.Result.Items[0].LatestVersion | Should -Be '1.5'
        $deployment = Invoke-UpdateWorkflowFixture -HistoryPath $historyPath -Plan $preview.Result
        $deployment.Result.SuccessCount | Should -Be 1
        $deployment.Result.Results[0].EffectiveVersion | Should -Be '1.5'
    }

    It 'rechecks rules after building and blocks upload if an exclusion was added' {
        $historyPath = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $null = New-Item -ItemType Directory -Path $historyPath
        $preview = Invoke-UpdateWorkflowFixture -HistoryPath $historyPath -Mode All
        $deployment = Invoke-UpdateWorkflowFixture -HistoryPath $historyPath -Plan $preview.Result -Mode All -ChangeRulesDuringBuild
        $deployment.Result.Results[0].ReasonCode | Should -Be 'PlanChanged'
        $deployment.Deploys | Should -Be 0
        $deployment.Builds | Should -Be 1
    }

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
