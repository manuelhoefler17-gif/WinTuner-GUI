BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\Modules\WinTuner.UpdateRules.psm1') -Force
    Import-Module (Join-Path $PSScriptRoot '..\Modules\WinTuner.Core.psm1') -Force
}

Describe 'Tenant-scoped persisted update rules' {
    BeforeEach {
        $tenant = [guid]::NewGuid()
        $otherTenant = [guid]::NewGuid()
    }

    It 'preserves other package rules, updates case-insensitively and isolates tenants' {
        Set-WinTunerUpdateRule -TenantId $tenant -PackageId 'Vendor.App' -ExcludeFromAll $true -Reason 'Pilot rollout' -BasePath $TestDrive
        Set-WinTunerUpdateRule -TenantId $tenant -PackageId 'Vendor.Other' -PinnedVersion '1.2' -BasePath $TestDrive
        Set-WinTunerUpdateRule -TenantId $tenant -PackageId 'vendor.app' -PinnedVersion '2.0' -BasePath $TestDrive
        $rules = @(Get-WinTunerUpdateRules -TenantId $tenant -BasePath $TestDrive)
        $rules | Should -HaveCount 2
        @($rules | Where-Object PackageId -eq 'vendor.app')[0].PinnedVersion | Should -Be '2.0'
        @(Get-WinTunerUpdateRules -TenantId $otherTenant -BasePath $TestDrive) | Should -HaveCount 0
        Set-WinTunerUpdateRule -TenantId $tenant -PackageId 'Vendor.App' -Remove -BasePath $TestDrive
        @(Get-WinTunerUpdateRules -TenantId $tenant -BasePath $TestDrive) | Should -HaveCount 1
    }

    It 'fails closed when the rule document is malformed instead of ignoring exclusions' {
        Set-WinTunerUpdateRule -TenantId $tenant -PackageId 'Vendor.App' -ExcludeFromAll $true -BasePath $TestDrive
        $path = Join-Path $TestDrive ($tenant.ToString('D') + '.json')
        Set-Content -LiteralPath $path -Value '{ invalid json'
        { Get-WinTunerUpdateRules -TenantId $tenant -BasePath $TestDrive } | Should -Throw '*blocked*'
        { Set-WinTunerUpdateRule -TenantId $tenant -PackageId 'Vendor.Other' -ExcludeFromAll $true -BasePath $TestDrive } | Should -Throw '*blocked*'
        (Get-Content -LiteralPath $path -Raw) | Should -Match 'invalid json'
    }

    It 'rejects wildcard pins and invalid expiry values' {
        { New-WinTunerUpdateRule -PackageId 'Vendor.App' -PinnedVersion '2.*' } | Should -Throw
        { New-WinTunerUpdateRule -PackageId 'Vendor.App' -ExpiresUtc 'not-a-date' } | Should -Throw
    }
}

Describe 'Update exclusion and pin decisions' {
    BeforeEach {
        $app = [pscustomobject]@{ Name = 'App'; GraphId = 'old-id'; PackageId = 'Vendor.App'; CurrentVersion = '1.0'; LatestVersion = '3.0' }
    }

    It 'keeps ordinary update behavior when there is no rule' {
        $decision = Get-WinTunerUpdateRuleDecision -App $app
        $decision.CanUpdateAll | Should -BeTrue
        $decision.CanUpdateSelected | Should -BeTrue
        $decision.TargetVersion | Should -Be '3.0'
    }

    It 'excludes only All and leaves an explicitly selected update available' {
        $rule = New-WinTunerUpdateRule -PackageId 'Vendor.App' -ExcludeFromAll $true -Reason 'Await approval'
        $decision = Get-WinTunerUpdateRuleDecision -App $app -Rules @($rule)
        $decision.CanUpdateAll | Should -BeFalse
        $decision.CanUpdateSelected | Should -BeTrue
        @(Resolve-WinTunerUpdateCandidates -Apps @($app) -Rules @($rule) -Mode All).Skipped | Should -HaveCount 1
    }

    It 'pins the exact target and leaves the original scanned candidate unchanged' {
        $rule = New-WinTunerUpdateRule -PackageId 'Vendor.App' -PinnedVersion '2.0'
        $selection = Resolve-WinTunerUpdateCandidates -Apps @($app) -Rules @($rule)
        $selection.Apps[0].LatestVersion | Should -Be '2.0'
        $selection.Apps[0].ScannedLatestVersion | Should -Be '3.0'
        $app.LatestVersion | Should -Be '3.0'
        $selection.Apps[0].GraphId | Should -Be 'old-id'
    }

    It 'holds the installed version and prevents downgrades for both actions' -TestCases @(@{ Pin = '1.0' },@{ Pin = '0.9' }) {
        param($Pin)
        $rule = New-WinTunerUpdateRule -PackageId 'Vendor.App' -PinnedVersion $Pin
        $decision = Get-WinTunerUpdateRuleDecision -App $app -Rules @($rule)
        $decision.CanUpdateAll | Should -BeFalse
        $decision.CanUpdateSelected | Should -BeFalse
    }

    It 'ignores a rule at its expiry boundary and restores the scanned latest target' {
        $rule = New-WinTunerUpdateRule -PackageId 'Vendor.App' -ExcludeFromAll $true -PinnedVersion '1.0' -ExpiresUtc '2026-10-29T15:00:00Z'
        $decision = Get-WinTunerUpdateRuleDecision -App $app -Rules @($rule) -Now ([datetimeoffset]'2026-10-29T15:00:00Z')
        $decision.CanUpdateAll | Should -BeTrue
        $decision.TargetVersion | Should -Be '3.0'
        $decision.Description | Should -Match 'expired'
    }

    It 'keeps rules attached to the package when a deployment replaces the Graph ID' {
        $rule = New-WinTunerUpdateRule -PackageId 'Vendor.App' -ExcludeFromAll $true
        $app.GraphId = 'replacement-id'
        (Get-WinTunerUpdateRuleDecision -App $app -Rules @($rule)).CanUpdateAll | Should -BeFalse
    }

    It 'calculates button safety separately for All and checked eligibility' {
        $state = Get-WinTunerUpdateActionState -Connected $true -IsBusy $false -CandidateCount 2 -CheckedCount 1 -EligibleAllCount 0 -EligibleCheckedCount 1 -SelectableCount 1
        $state.CanUpdateAll | Should -BeFalse
        $state.CanUpdateSelected | Should -BeTrue
        $held = Get-WinTunerUpdateActionState -Connected $true -IsBusy $false -CandidateCount 2 -CheckedCount 1 -EligibleAllCount 0 -EligibleCheckedCount 0 -SelectableCount 0
        $held.CanUpdateSelected | Should -BeFalse
        $held.CanCheckAll | Should -BeFalse
        $held.CanUncheckAll | Should -BeTrue
    }
}
