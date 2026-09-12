$ErrorActionPreference = 'Stop'

Describe 'Microsoft Graph permission documentation' {
    BeforeAll {
        $script:repositoryRoot = Split-Path -Parent $PSScriptRoot
        $script:readme = Get-Content -LiteralPath (Join-Path $script:repositoryRoot 'README.md') -Raw
        $script:intuneModule = Get-Content -LiteralPath (Join-Path $script:repositoryRoot 'Modules/WinTuner.Intune.psm1') -Raw
        $script:gui = Get-Content -LiteralPath (Join-Path $script:repositoryRoot 'WinTuner_GUI.ps1') -Raw
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

    It 'exposes the permission summary from Settings before sign-in' {
        $script:gui | Should -Match ([regex]::Escape('$graphPermissionsButton.Text = "Graph Permissions..."'))
        $script:gui | Should -Match ([regex]::Escape('"Required Microsoft Graph Permissions"'))
        $script:readme | Should -Match ([regex]::Escape('Settings > Graph Permissions...'))
    }

    It 'keeps every documented scope in the Settings permission dialog' {
        $effectiveScopes = @(
            'DeviceManagementApps.ReadWrite.All'
            'DeviceManagementConfiguration.ReadWrite.All'
            'DeviceManagementManagedDevices.Read.All'
            'Directory.Read.All'
        )

        foreach ($scope in $effectiveScopes) {
            $script:gui | Should -Match ([regex]::Escape($scope))
        }
    }

    It 'separates interactive User Login from customer-owned Entra Application requirements' {
        foreach ($text in @(
            'User login (interactive)'
            'No customer-owned Entra App Registration is required'
            'Entra application (client secret or certificate)'
            'only these Microsoft Graph **application permissions**'
            'Delegated permissions and'
            'User.Read'
            'are not required on this App Registration'
        )) {
            $script:readme | Should -Match ([regex]::Escape($text))
        }

        $script:gui | Should -Match ([regex]::Escape('USER LOGIN (Interactive)'))
        $script:gui | Should -Match ([regex]::Escape('ENTRA APPLICATION (Client secret or certificate)'))
        $script:gui | Should -Match ([regex]::Escape('Add only these Microsoft Graph > Application permissions:'))
        $script:gui | Should -Match ([regex]::Escape('Delegated permissions and User.Read are not required'))
    }
}
