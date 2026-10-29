BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\Modules\WinTuner.AppUpdate.psm1'
    Import-Module $modulePath -Force
}

Describe 'Invoke-WinTunerAppUpdateBatch' {
    It 'builds, validates, and deploys an update while preserving the Graph identifier' {
        $calls = [System.Collections.Generic.List[string]]::new()
        $app = [pscustomobject]@{
            Name           = 'Vendor App'
            CurrentVersion = '1.0.0'
            LatestVersion  = '2.0.0'
            GraphId        = 'graph-123'
            PackageId      = 'Vendor.App'
        }

        $result = Invoke-WinTunerAppUpdateBatch `
            -Apps @($app) `
            -RootPackageFolder $TestDrive `
            -BuildPackage {
                param($Id, $Folder, $Target, $Installed)
                $calls.Add("build|$Id|$Folder|$Target|$Installed")
                [pscustomobject]@{ Succeeded = $true; EffectiveVersion = $Target }
            }.GetNewClosure() `
            -ValidateArtifact {
                param($Root, $Id, $Version)
                $calls.Add("validate|$Root|$Id|$Version")
                [pscustomobject]@{ IsValid = $true; IntuneWinPath = (Join-Path $Root 'package.intunewin') }
            }.GetNewClosure() `
            -DeployPackage {
                param($App, $Id, $Version, $Root)
                $calls.Add("deploy|$($App.GraphId)|$Id|$Version|$Root")
                [pscustomobject]@{ Id = 'new-graph-456' }
            }.GetNewClosure() `
            -ConfigureAvailableAssignmentAutoUpdate {
                param($GraphId)
                $calls.Add("autoupdate|$GraphId")
                [pscustomobject]@{ Summary = 'Enabled AutoUpdate for 1 Available assignment(s); already enabled for 0.' }
            }.GetNewClosure()

        $result.SuccessCount | Should -Be 1
        $result.FailureCount | Should -Be 0
        $result.Results[0].ReasonCode | Should -Be 'Updated'
        $calls | Should -HaveCount 4
        $calls[0] | Should -Be "build|Vendor.App|$TestDrive|2.0.0|1.0.0"
        $calls[1] | Should -Be "validate|$TestDrive|Vendor.App|2.0.0"
        $calls[2] | Should -Be "deploy|graph-123|Vendor.App|2.0.0|$TestDrive"
        $calls[3] | Should -Be 'autoupdate|new-graph-456'
        $result.Results[0].AutoUpdateSummary | Should -Be 'Enabled AutoUpdate for 1 Available assignment(s); already enabled for 0.'
    }

    It 'blocks deployment when exact artifact validation fails' {
        $deployCalls = 0
        $result = Invoke-WinTunerAppUpdateBatch `
            -Apps @([pscustomobject]@{
                Name = 'Unsafe App'; CurrentVersion = '1'; LatestVersion = '2'; PackageId = 'Unsafe.App'
            }) `
            -RootPackageFolder $TestDrive `
            -BuildPackage { [pscustomobject]@{ Succeeded = $true; EffectiveVersion = '2' } } `
            -ValidateArtifact { [pscustomobject]@{ IsValid = $false; ReasonCode = 'MissingIntuneWin'; Reason = 'Expected file is missing.' } } `
            -DeployPackage { $deployCalls++ }.GetNewClosure()

        $result.SuccessCount | Should -Be 0
        $result.FailureCount | Should -Be 1
        $result.Results[0].ReasonCode | Should -Be 'MissingIntuneWin'
        $deployCalls | Should -Be 0
    }

    It 'blocks package creation when the verified scan result has no PackageId' {
        $buildCalls = 0
        $result = Invoke-WinTunerAppUpdateBatch `
            -Apps @([pscustomobject]@{
                Name = 'Unknown App'; CurrentVersion = '1'; LatestVersion = '2'; PackageId = ''
            }) `
            -RootPackageFolder $TestDrive `
            -BuildPackage { $buildCalls++ }.GetNewClosure()

        $result.FailureCount | Should -Be 1
        $result.Results[0].ReasonCode | Should -Be 'MissingPackageId'
        $buildCalls | Should -Be 0
    }

    It 'continues with later apps after a deployment failure' {
        $deployCalls = [System.Collections.Generic.List[string]]::new()
        $apps = @(
            [pscustomobject]@{ Name = 'First'; CurrentVersion = '1'; LatestVersion = '2'; PackageId = 'Vendor.First' }
            [pscustomobject]@{ Name = 'Second'; CurrentVersion = '1'; LatestVersion = '2'; PackageId = 'Vendor.Second' }
        )

        $result = Invoke-WinTunerAppUpdateBatch `
            -Apps $apps `
            -RootPackageFolder $TestDrive `
            -BuildPackage { param($Id, $Folder, $Target) [pscustomobject]@{ Succeeded = $true; EffectiveVersion = $Target } } `
            -ValidateArtifact { [pscustomobject]@{ IsValid = $true } } `
            -DeployPackage {
                param($App, $Id)
                $deployCalls.Add($Id)
                if ($Id -eq 'Vendor.First') { throw 'Tenant rejected deployment' }
            }.GetNewClosure()

        $result.SuccessCount | Should -Be 1
        $result.FailureCount | Should -Be 1
        $result.Results[0].ReasonCode | Should -Be 'DeploymentFailed'
        $result.Results[1].ReasonCode | Should -Be 'Updated'
        $deployCalls | Should -Be @('Vendor.First', 'Vendor.Second')
    }

    It 'enables AutoUpdate only for Available assignments and preserves their existing settings' {
        $requests = [System.Collections.Generic.List[object]]::new()
        $graphRequest = {
            param($Method, $Uri, $Body)
            $requests.Add([pscustomobject]@{ Method = $Method; Uri = $Uri; Body = $Body })
            if ($Method -eq 'GET') {
                return @{ value = @(
                    [pscustomobject]@{
                        id = 'available-assignment'
                        intent = 'available'
                        settings = [pscustomobject]@{
                            '@odata.type' = '#microsoft.graph.win32LobAppAssignmentSettings'
                            notifications = 'showAll'
                        }
                    },
                    [pscustomobject]@{
                        id = 'required-assignment'
                        intent = 'required'
                        settings = [pscustomobject]@{ '@odata.type' = '#microsoft.graph.win32LobAppAssignmentSettings' }
                    }
                ) }
            }
        }.GetNewClosure()

        $module = Get-Module WinTuner.AppUpdate
        $result = & $module {
            param($Request)
            Set-WinTunerAvailableAssignmentAutoUpdate -GraphId 'new-app-id' -InvokeGraphRequest $Request
        } $graphRequest

        $result.AvailableCount | Should -Be 1
        $result.EnabledCount | Should -Be 1
        $requests | Should -HaveCount 2
        $requests[0].Method | Should -Be 'GET'
        $requests[1].Method | Should -Be 'PATCH'
        $requests[1].Uri | Should -Match '/assignments/available-assignment$'
        $body = $requests[1].Body | ConvertFrom-Json -AsHashtable
        $body.settings.notifications | Should -Be 'showAll'
        $body.settings.autoUpdateSettings.autoUpdateSupersededAppsState | Should -Be 'enabled'
    }
}
