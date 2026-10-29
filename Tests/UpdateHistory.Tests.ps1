BeforeAll { Import-Module (Join-Path $PSScriptRoot '..\Modules\WinTuner.UpdateHistory.psm1') -Force }

Describe 'Tenant-scoped update history' {
    BeforeEach {
        $tenantA = [guid]::NewGuid().ToString('D')
        $tenantB = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'
        $result = [pscustomobject]@{
            Name = 'Vendor App'; PackageId = 'Vendor.App'; GraphId = 'graph-id'
            CurrentVersion = '1.0'; TargetVersion = '2.0'; EffectiveVersion = '2.0'
            Succeeded = $true; ReasonCode = 'Updated'; AutoUpdateSummary = ''
            Message = 'must-not-be-persisted'; AccessToken = 'must-not-be-persisted'
        }
    }

    It 'persists only whitelisted fields and keeps tenants separate' {
        Add-WinTunerUpdateHistory -TenantId $tenantA -OperationId ([guid]::NewGuid()) -Result $result -BasePath $TestDrive
        $history = @(Get-WinTunerUpdateHistory -TenantId $tenantA -BasePath $TestDrive)
        $history | Should -HaveCount 1
        $history[0].Status | Should -Be 'Updated'
        $history[0].TimestampUtc | Should -Match '\+00:00$'
        @(Get-WinTunerUpdateHistory -TenantId $tenantB -BasePath $TestDrive) | Should -HaveCount 0
        (Get-ChildItem $TestDrive -Recurse -Filter '*.json' | Get-Content -Raw) | Should -Not -Match 'must-not-be-persisted'
    }

    It 'records warnings and failures with distinct status and no overwritten entries' {
        $result.AutoUpdateSummary = 'Warning: configuration could not be completed'
        Add-WinTunerUpdateHistory -TenantId $tenantA -OperationId ([guid]::NewGuid()) -Result $result -BasePath $TestDrive
        $result.Succeeded = $false
        $result.ReasonCode = 'PackageBuildFailed'
        Add-WinTunerUpdateHistory -TenantId $tenantA -OperationId ([guid]::NewGuid()) -Result $result -BasePath $TestDrive
        $history = @(Get-WinTunerUpdateHistory -TenantId $tenantA -BasePath $TestDrive)
        $history | Should -HaveCount 2
        $history.Status | Should -Contain 'UpdatedWithWarning'
        $history.Status | Should -Contain 'Failed'
    }

    It 'exports only approved fields and neutralizes spreadsheet formulas' {
        $result.Name = '=HYPERLINK("https://example.com")'
        Add-WinTunerUpdateHistory -TenantId $tenantA -OperationId ([guid]::NewGuid()) -Result $result -BasePath $TestDrive
        $path = Join-Path $TestDrive 'history.csv'
        Export-WinTunerUpdateHistory -Entries @(Get-WinTunerUpdateHistory -TenantId $tenantA -BasePath $TestDrive) -Path $path
        $csv = @(Import-Csv $path)
        $csv[0].Name | Should -BeExactly ("'" + $result.Name)
        $csv[0].PSObject.Properties.Name | Should -Not -Contain 'AccessToken'
    }
}
