BeforeAll {
    $root = Split-Path $PSScriptRoot -Parent
    Import-Module (Join-Path $root 'Modules\WinTuner.Logging.psm1') -Force
    Import-Module (Join-Path $root 'Modules\WinTuner.PackageBuild.psm1') -Force
    Import-Module (Join-Path $root 'Modules\WinTuner.AppUpdate.psm1') -Force
    $tokens=$null; $errors=$null
    $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'WinTuner_GUI.ps1'),[ref]$tokens,[ref]$errors)
    foreach ($name in @('Update-Status','Update-WinTunerAppUpdateRetryStatus')) {
        $definition = $ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)
        Invoke-Expression $definition.Extent.Text
    }
    function Invoke-UiAction { param($Control,$Action) & $Action }
}

Describe 'Retry progress across build, JSON and GUI logging' {
    It 'preserves HTTP status from a retried build and logs each retry once' {
        Mock Start-Sleep -ModuleName WinTuner.Retry {}
        $state = @{ Builds = 0; Deploys = 0 }
        $progress = [Collections.Generic.List[object]]::new()
        $creator = {
            param($Id,$Folder,$Version)
            $state.Builds++
            if ($state.Builds -eq 1) { throw 'HTTP 429 Too Many Requests' }
        }.GetNewClosure()
        $builder = {
            param($Id,$Folder,$Target,$Installed,$RetryReporter)
            Invoke-WinTunerPackageBuild -PackageId $Id -PackageFolder $Folder -DesiredVersion $Target -CreatePackage $creator -ReportRetry $RetryReporter
        }.GetNewClosure()
        $app = [pscustomobject]@{ Name='Claude'; CurrentVersion='1.0'; LatestVersion='2.0'; GraphId='old-app'; PackageId='Anthropic.Claude' }
        $result = Invoke-WinTunerAppUpdateBatch -Apps @($app) -RootPackageFolder $TestDrive -BuildPackage $builder -ValidateArtifact {
            [pscustomobject]@{ IsValid=$true }
        } -DeployPackage {
            $state.Deploys++
            [pscustomobject]@{ Id='new-app' }
        }.GetNewClosure() -ConfigureAvailableAssignmentAutoUpdate {
            [pscustomobject]@{ Summary='No Available assignments found.' }
        } -ReportProgress {param($info) $progress.Add($info)}.GetNewClosure()
        $result.SuccessCount | Should -Be 1
        $state.Builds | Should -Be 2
        $state.Deploys | Should -Be 1
        $retries = @($progress | Where-Object Stage -EQ 'Retry')
        $retries.Count | Should -Be 1
        $roundTrip = $retries[0] | ConvertTo-Json -Compress | ConvertFrom-Json
        $roundTrip.StatusCode | Should -Be 429
        $roundTrip.DelaySeconds | Should -Be 15
        Initialize-WinTunerLogging -BasePath $TestDrive
        $script:statusLabel = [pscustomobject]@{ Text='' }
        Update-WinTunerAppUpdateRetryStatus -ProgressInfo $roundTrip
        $lines = @(Get-Content -LiteralPath (Join-Path $TestDrive 'WinTuner_GUI.log'))
        $lines.Count | Should -Be 1
        $lines[0] | Should -Match 'Waiting 15s before retry 2/3: Claude.*HTTP 429 \(Too Many Requests\)'
        $script:statusLabel.Text | Should -Match 'HTTP 429'
    }

    It 'shows numeric server failures and legacy or timeout progress without logging raw service errors' {
        $directory = Join-Path $TestDrive 'other-progress'
        Initialize-WinTunerLogging -BasePath $directory
        $script:statusLabel = [pscustomobject]@{ Text='' }
        $info = [pscustomobject]@{ DelaySeconds=3; Attempt=2; MaxAttempts=3; AppName='App'; StatusCode=503; ErrorMessage='sensitive-service-response' }
        Update-WinTunerAppUpdateRetryStatus -ProgressInfo $info
        $script:statusLabel.Text | Should -Match 'HTTP 503'
        $info.PSObject.Properties.Remove('StatusCode')
        Update-WinTunerAppUpdateRetryStatus -ProgressInfo $info
        $script:statusLabel.Text | Should -Match 'Temporary network error or timeout'
        $lines = @(Get-Content -LiteralPath (Join-Path $directory 'WinTuner_GUI.log'))
        $lines.Count | Should -Be 2
        ($lines -join '') | Should -Not -Match 'sensitive-service-response'
    }
}
