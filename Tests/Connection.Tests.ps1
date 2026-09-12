BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\Modules\WinTuner.Connection.psm1') -Force
}

Describe 'Invoke-WinTunerConnectionVerification' {
    It 'succeeds on the first tenant query even when no apps exist' {
        $result = Invoke-WinTunerConnectionVerification -GetApps { @() } -Delay { throw 'delay must not run' }
        $result.Succeeded | Should -BeTrue
        $result.Attempts | Should -Be 1
    }

    It 'retries transient failures and then succeeds' {
        $state = [pscustomobject]@{ Attempts = 0 }
        $delays = [System.Collections.Generic.List[int]]::new()
        $result = Invoke-WinTunerConnectionVerification -GetApps {
            $state.Attempts++
            if ($state.Attempts -lt 3) { throw 'temporary connection failure' }
            @([pscustomobject]@{ Name = 'App' })
        }.GetNewClosure() -MaxAttempts 4 -RetryDelayMilliseconds 25 -Delay {
            param($Milliseconds)
            $delays.Add($Milliseconds)
        }.GetNewClosure()

        $result.Succeeded | Should -BeTrue
        $result.Attempts | Should -Be 3
        $delays | Should -Be @(25, 25)
    }

    It 'returns the final verification error after all attempts fail' {
        $result = Invoke-WinTunerConnectionVerification -GetApps { throw 'authentication expired' } -MaxAttempts 3 -RetryDelayMilliseconds 0
        $result.Succeeded | Should -BeFalse
        $result.Attempts | Should -Be 3
        $result.ErrorMessage | Should -Be 'authentication expired'
    }
}

Describe 'GUI connection state integration' {
    BeforeAll {
        $guiPath = Join-Path $PSScriptRoot '..\WinTuner_GUI.ps1'
        $tokens = $null
        $parseErrors = $null
        $script:guiAst = [System.Management.Automation.Language.Parser]::ParseFile($guiPath, [ref]$tokens, [ref]$parseErrors)
        $parseErrors | Should -HaveCount 0
        $script:intuneModuleText = Get-Content -LiteralPath (Join-Path $PSScriptRoot '..\Modules\WinTuner.Intune.psm1') -Raw
    }

    It 'recomputes the Login button state when connection state changes' {
        $function = @($script:guiAst.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $node.Name -eq 'Set-ConnectedUIState'
        }, $true))
        $function | Should -HaveCount 1
        $function[0].Extent.Text | Should -Match 'Update-WinTunerLoginButtonState'
        $function[0].Extent.Text | Should -Match 'Update-WinTunerAuthenticationUI'
    }

    It 'connects Microsoft Graph during the initial login workflow' {
        $function = @($script:guiAst.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $node.Name -eq 'Start-WinTunerLogin'
        }, $true))
        $function | Should -HaveCount 1
        $function[0].Extent.Text | Should -Match 'Connect-WinTunerGraph\s+-AuthenticationMode\s+Interactive\s+-UserPrincipalName\s+\$UserPrincipalName'
        $function[0].Extent.Text | Should -Match 'Confirm-WinTunerGraphContext'
        $function[0].Extent.Text | Should -Match 'AddArgument\(\$configuration\.Mode\).*AddArgument\(\$upn\).*AddArgument\(\$graphTenantId\)'
    }

    It 'stores the Microsoft Graph context for reuse across runspaces' {
        $script:intuneModuleText | Should -Match 'Connect-MgGraph(?s).*?-ContextScope\s+Process'
    }
}
Describe 'App-only Graph permission preflight' {
    It 'validates both required read-only Graph paths' {
        $calls = [System.Collections.Generic.List[string]]::new()
        $result = Invoke-WinTunerGraphPermissionPreflight -GetManagedApps {
            $calls.Add('apps')
            [pscustomobject]@{ value = @() }
        }.GetNewClosure() -GetDetectedApps {
            $calls.Add('detected')
            [pscustomobject]@{ value = @() }
        }.GetNewClosure()

        $result.Succeeded | Should -BeTrue
        $result.ChecksCompleted | Should -Be @(
            'DeviceManagementApps.ReadWrite.All',
            'DeviceManagementManagedDevices.Read.All'
        )
        $calls | Should -Be @('apps', 'detected')
    }

    It 'maps an apps 403 to the exact required Application permission and stops' {
        $detectedCalled = $false
        $result = Invoke-WinTunerGraphPermissionPreflight -GetManagedApps {
            throw 'HTTP 403 Forbidden with raw service details'
        } -GetDetectedApps {
            $detectedCalled = $true
        }.GetNewClosure()

        $result.Succeeded | Should -BeFalse
        $result.Forbidden | Should -BeTrue
        $result.MissingPermission | Should -Be 'DeviceManagementApps.ReadWrite.All'
        $result.ErrorMessage | Should -Match 'Grant admin consent'
        $result.ErrorMessage | Should -Not -Match 'raw service details'
        $detectedCalled | Should -BeFalse
    }

    It 'maps a detected-apps 403 to the exact required Application permission' {
        $result = Invoke-WinTunerGraphPermissionPreflight -GetManagedApps { @() } -GetDetectedApps {
            throw 'Forbidden'
        }

        $result.Succeeded | Should -BeFalse
        $result.Forbidden | Should -BeTrue
        $result.MissingPermission | Should -Be 'DeviceManagementManagedDevices.Read.All'
        $result.ChecksCompleted | Should -Be @('DeviceManagementApps.ReadWrite.All')
    }

    It 'keeps non-permission failures distinguishable' {
        $result = Invoke-WinTunerGraphPermissionPreflight -GetManagedApps { throw 'network timeout' } -GetDetectedApps { @() }
        $result.Succeeded | Should -BeFalse
        $result.Forbidden | Should -BeFalse
        $result.MissingPermission | Should -BeNullOrEmpty
        $result.ErrorMessage | Should -Match 'network timeout'
    }
}
