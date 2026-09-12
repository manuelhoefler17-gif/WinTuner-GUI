BeforeAll {
    $guiPath = Join-Path $PSScriptRoot '..\WinTuner_GUI.ps1'
    $tokens = $null
    $parseErrors = $null
    $script:guiAst = [System.Management.Automation.Language.Parser]::ParseFile($guiPath, [ref]$tokens, [ref]$parseErrors)
    $parseErrors | Should -HaveCount 0
    $script:guiText = Get-Content -LiteralPath $guiPath -Raw
}

Describe 'GUI authentication modes' {
    It 'declares version 0.10.21 and bootstraps the authentication module' {
        $script:guiText | Should -Match '\$script:AppVersion\s*=\s*"0\.10\.21"'
        $script:guiText | Should -Match "Modules/WinTuner\.Authentication\.psm1"
        $script:guiText | Should -Match '\$authenticationModulePath\s*=\s*Join-Path.*WinTuner\.Authentication\.psm1'
        $script:guiText | Should -Match 'Import-Module\s+\$authenticationModulePath\s+-Force'
    }

    It 'stores client secrets only through a masked DPAPI-protected field' {
        $function = @($script:guiAst.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $node.Name -eq 'Show-WinTunerAuthenticationSettings'
        }, $true))
        $function | Should -HaveCount 1
        $text = $function[0].Extent.Text
        $text.Contains('$script:settings.ClientSecret') | Should -BeFalse
        $text | Should -Match ([regex]::Escape('$secretBox.UseSystemPasswordChar = $true'))
        $text | Should -Match 'Protect-WinTunerClientSecretForCurrentUser'
        $text | Should -Match 'EntraClientSecretProtected'
        $text | Should -Match 'plaintext is never saved'
        $text | Should -Match ([regex]::Escape('$secretBox.Clear()'))
    }

    It 'captures the main settings instance before creating the Save event closure' {
        $function = @($script:guiAst.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $node.Name -eq 'Show-WinTunerAuthenticationSettings'
        }, $true))
        $function | Should -HaveCount 1
        $text = $function[0].Extent.Text
        $text | Should -Match '\$settingsReference\s*=\s*\$script:settings'
        $text | Should -Match 'Export-WinTunerSettings\s+-Settings\s+\$settingsReference\s+-Path\s+\$settingsPathReference'
        $saveHandler = [regex]::Match($text, '\$saveButton\.Add_Click\(\{(?s).*?\}\.GetNewClosure\(\)\)').Value
        $saveHandler | Should -Not -Match '\$script:settings'
    }
    It 'masks and clears the one-login client secret' {
        $function = @($script:guiAst.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $node.Name -eq 'Show-WinTunerClientSecretPrompt'
        }, $true))
        $function | Should -HaveCount 1
        $function[0].Extent.Text | Should -Match 'UseSystemPasswordChar\s*=\s*\$true'
        $function[0].Extent.Text | Should -Match 'finally(?s).*?\$secretBox\.Clear\(\)'
    }

    It 'uses a saved protected secret before falling back to the one-login prompt' {
        $function = @($script:guiAst.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $node.Name -eq 'Start-WinTunerLogin'
        }, $true))
        $function | Should -HaveCount 1
        $text = $function[0].Extent.Text
        $text | Should -Match 'Unprotect-WinTunerClientSecretForCurrentUser'
        $text | Should -Match 'EntraClientSecretProtected'
        $text | Should -Match 'Show-WinTunerClientSecretPrompt'
        $text.IndexOf('Unprotect-WinTunerClientSecretForCurrentUser') | Should -BeLessThan $text.IndexOf('Show-WinTunerClientSecretPrompt')
    }

    It 'shows actionable app-only Intune permission guidance for a 403 response' {
        $function = @($script:guiAst.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $node.Name -eq 'Show-WinTunerLoginError'
        }, $true))
        $function | Should -HaveCount 1
        $text = $function[0].Extent.Text
        $text | Should -Match 'Forbidden'
        $text | Should -Match 'DeviceManagementApps\.ReadWrite\.All'
        $text | Should -Match 'DeviceManagementManagedDevices\.Read\.All'
        $text | Should -Match 'fresh app-only token'
        $text | Should -Match 'Application permissions'
        $text | Should -Match 'Grant admin consent'
    }

    It 'clears only the WinTuner app-only token cache and checks both Graph permissions before accepting login' {
        $function = @($script:guiAst.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $node.Name -eq 'Start-WinTunerLogin'
        }, $true))
        $function | Should -HaveCount 1
        $text = $function[0].Extent.Text
        $text | Should -Match 'Clear-WinTunerAppOnlyTokenCache'
        $text | Should -Match 'Invoke-WinTunerGraphPermissionPreflight'
        $text | Should -Match 'deviceAppManagement/mobileApps\?\$top=1'
        $text | Should -Match 'deviceManagement/detectedApps\?\$top=1'
        $text | Should -Match 'PermissionPreflightPassed'
        $text | Should -Not -Match 'mg\.msal\.cache'
    }
    It 'validates certificate availability and clears client-secret connection parameters' {
        $function = @($script:guiAst.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $node.Name -eq 'Start-WinTunerLogin'
        }, $true))
        $function | Should -HaveCount 1
        $function[0].Extent.Text | Should -Match 'Test-WinTunerCertificateAvailable'
        $function[0].Extent.Text | Should -Match "Remove\('ClientSecret'\)"
        $function[0].Extent.Text | Should -Match '\$clientSecret\s*=\s*\$null'
    }

    It 'reuses only validated app-only Graph context in Discovery workers' {
        $function = @($script:guiAst.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $node.Name -eq 'Start-WinTunerDiscoveryScan'
        }, $true))
        $function | Should -HaveCount 1
        $function[0].Extent.Text | Should -Match 'Confirm-WinTunerGraphContext'
        $function[0].Extent.Text | Should -Match 'AddArgument\(\$script:currentAuthenticationMode\)'
        $function[0].Extent.Text | Should -Match 'AddArgument\(\$script:currentGraphTenantId\)'
        $function[0].Extent.Text | Should -Not -Match 'AddArgument\(\$clientSecret\)'
    }

    It 'clears app-only runtime identity state during logout' {
        $script:guiText | Should -Match '\$logoutButton\.Add_Click\(\{(?s).*?\$script:currentIdentityLabel\s*=\s*''''.*?\$script:currentAuthenticationMode\s*=\s*''Interactive''.*?\$script:currentGraphTenantId\s*=\s*''''.*?\$script:currentClientId\s*=\s*'''''
    }
}
