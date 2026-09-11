$ErrorActionPreference = 'Stop'

Describe 'Microsoft Graph permission documentation' {
    BeforeAll {
        $script:repositoryRoot = Split-Path -Parent $PSScriptRoot
        $script:readme = Get-Content -LiteralPath (Join-Path $script:repositoryRoot 'README.md') -Raw
        $script:intuneModule = Get-Content -LiteralPath (Join-Path $script:repositoryRoot 'Modules/WinTuner.Intune.psm1') -Raw
    }

    It 'documents every scope explicitly requested by the GUI Graph connection' {
        $guiScopes = @(
            'DeviceManagementApps.ReadWrite.All'
            'DeviceManagementManagedDevices.Read.All'
            'Directory.Read.All'
        )

        foreach ($scope in $guiScopes) {
            $script:intuneModule | Should -Match ([regex]::Escape($scope))
            $script:readme | Should -Match ([regex]::Escape($scope))
        }
    }

    It 'documents the WinTuner default scope and tenant approval requirements' {
        $script:readme | Should -Match ([regex]::Escape('DeviceManagementConfiguration.ReadWrite.All'))
        $script:readme | Should -Match 'delegated permissions'
        $script:readme | Should -Match 'administrator must grant consent'
        $script:readme | Should -Match 'appropriate Intune role'
    }
}