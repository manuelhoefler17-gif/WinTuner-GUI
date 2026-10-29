BeforeAll { Import-Module (Join-Path $PSScriptRoot '..\Modules\WinTuner.AppUpdate.psm1') -Force }

Describe 'Read-only deployment planning' {
    BeforeEach {
        $app = [pscustomobject]@{ Name = 'Test App'; PackageId = 'Vendor.App'; GraphId = 'app-id'; CurrentVersion = '1.0'; LatestVersion = '2.0' }
        $tenant = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
        $getApp = { [pscustomobject]@{ id = 'app-id'; displayVersion = '1.0' } }
    }

    It 'shows assignments and AutoUpdate intent without deployment' {
        $plan = New-WinTunerUpdatePlan -Apps @($app) -TenantId $tenant -RootPackageFolder $TestDrive -GetApp $getApp -GetAssignments {
            [pscustomobject]@{ id = 'a'; intent = 'available'; target = @{ '@odata.type' = '#microsoft.graph.groupAssignmentTarget'; groupId = 'group-a' }; settings = @{ notifications = 'showAll' } }
            [pscustomobject]@{ id = 'b'; intent = 'required'; target = @{ '@odata.type' = '#microsoft.graph.allDevicesAssignmentTarget' }; settings = @{} }
        }
        $plan.Items | Should -HaveCount 1
        $plan.Items[0].Assignments | Should -HaveCount 2
        $text = Format-WinTunerUpdatePlan $plan
        $text | Should -Match '1.0 -> 2.0'
        $text | Should -Match 'Group: group-a'
        $text | Should -Match 'AutoUpdate: enabled on new app'
        $text | Should -Match 'required \| All devices'
    }

    It 'blocks missing identifiers and a changed Intune version' {
        { New-WinTunerUpdatePlan -Apps @($app) -TenantId $tenant -RootPackageFolder $TestDrive -GetApp { @{ id = 'app-id'; displayVersion = '1.5' } } -GetAssignments {} } | Should -Throw '*changed*'
        $app.PackageId = ''
        { New-WinTunerUpdatePlan -Apps @($app) -TenantId $tenant -RootPackageFolder $TestDrive -GetApp $getApp -GetAssignments {} } | Should -Throw '*PackageId*'
    }

    It 'distinguishes a missing version field from a changed version and fails closed' {
        { New-WinTunerUpdatePlan -Apps @($app) -TenantId $tenant -RootPackageFolder $TestDrive -GetApp { @{ id = 'app-id' } } -GetAssignments { throw 'Assignments must not be read after version verification failed' } } | Should -Throw '*displayVersion is missing*'
    }

    It 'canonicalizes assignment property order but detects changed assignment settings' {
        $first = New-WinTunerUpdatePlan -Apps @($app) -TenantId $tenant -RootPackageFolder $TestDrive -GetApp $getApp -GetAssignments { @{ id = 'a'; intent = 'available'; settings = [ordered]@{ notifications = 'showAll'; restartSettings = @{ gracePeriodInMinutes = 10 } } } }
        $same = New-WinTunerUpdatePlan -Apps @($app) -TenantId $tenant -RootPackageFolder $TestDrive -GetApp $getApp -GetAssignments { [ordered]@{ settings = [ordered]@{ restartSettings = @{ gracePeriodInMinutes = 10 }; notifications = 'showAll' }; intent = 'available'; id = 'a' } }
        $changed = New-WinTunerUpdatePlan -Apps @($app) -TenantId $tenant -RootPackageFolder $TestDrive -GetApp $getApp -GetAssignments { @{ id = 'a'; intent = 'required'; settings = @{} } }
        $first.Items[0].AssignmentHash | Should -BeExactly $same.Items[0].AssignmentHash
        $first.Items[0].AssignmentHash | Should -Not -Be $changed.Items[0].AssignmentHash
    }

    It 'blocks deployment when final plan validation fails and records the failure' {
        $state = @{ Deploys = 0; History = @() }
        $batch = Invoke-WinTunerAppUpdateBatch -Apps @($app) -RootPackageFolder $TestDrive `
            -BuildPackage { @{ Succeeded = $true; EffectiveVersion = '2.0' } } `
            -ValidateArtifact { @{ IsValid = $true } } `
            -ValidatePlan { throw 'The assignments changed after preview.' } `
            -DeployPackage { $state.Deploys++ }.GetNewClosure() `
            -RecordResult { param($result) $state.History += $result }.GetNewClosure()
        $batch.FailureCount | Should -Be 1
        $batch.Results[0].ReasonCode | Should -Be 'PlanChanged'
        $state.Deploys | Should -Be 0
        $state.History.Count | Should -Be 1
    }

    It 'does not repeat a successful deployment when local history cannot be saved' {
        $app.GraphId = ''
        $state = @{ Deploys = 0 }
        $batch = Invoke-WinTunerAppUpdateBatch -Apps @($app) -RootPackageFolder $TestDrive `
            -BuildPackage { @{ Succeeded = $true; EffectiveVersion = '2.0' } } `
            -ValidateArtifact { @{ IsValid = $true } } `
            -DeployPackage { $state.Deploys++ }.GetNewClosure() -RecordResult { throw 'Disk full' }
        $batch.SuccessCount | Should -Be 1
        $batch.Results[0].HistoryWarning | Should -Match 'could not be saved'
        $state.Deploys | Should -Be 1
    }

    It 'never automatically replays a complete deployment after a 429 response' {
        $state = @{ Deploys = 0 }
        $batch = Invoke-WinTunerAppUpdateBatch -Apps @($app) -RootPackageFolder $TestDrive `
            -BuildPackage { @{ Succeeded = $true; EffectiveVersion = '2.0' } } `
            -ValidateArtifact { @{ IsValid = $true } } `
            -DeployPackage { $state.Deploys++; throw 'HTTP 429 Too Many Requests during upload' }.GetNewClosure()
        $batch.FailureCount | Should -Be 1
        $batch.Results[0].ReasonCode | Should -Be 'DeploymentFailed'
        $state.Deploys | Should -Be 1
    }
}
