Import-Module (Join-Path $PSScriptRoot '..\Modules\WinTuner.Intune.psm1') -Force

Describe 'Shared interactive permission request' {
    BeforeAll {
        Import-Module (Join-Path $PSScriptRoot '..\Modules\WinTuner.Authentication.psm1') -Force
    }
    It 'requests the existing WinTuner and Discovery permissions in one login' {
        $parameters = New-WinTunerModuleConnectionParameters -Mode Interactive -UserPrincipalName 'admin@contoso.com'
        $parameters.Username | Should -Be 'admin@contoso.com'
        $parameters.Scopes | Should -HaveCount 4
        $parameters.Scopes | Should -Contain 'DeviceManagementApps.ReadWrite.All'
        $parameters.Scopes | Should -Contain 'DeviceManagementConfiguration.ReadWrite.All'
        $parameters.Scopes | Should -Contain 'DeviceManagementManagedDevices.Read.All'
        $parameters.Scopes | Should -Contain 'Directory.Read.All'
    }
}

InModuleScope WinTuner.Intune {
    BeforeAll {
        # Test-only stubs keep this suite independent of installed auth modules.
        function script:Get-WtToken { param($ShowToken, $DecodeToken) throw 'Unmocked token request' }
        function script:Get-MgContext { throw 'Unmocked Graph context request' }
        function script:Disconnect-MgGraph { throw 'Unmocked Graph disconnect' }
        function script:Invoke-MgGraphRequest { param($Uri, $Method, $Body, $ContentType) throw 'Unmocked Graph request' }
    }

    Describe 'Shared interactive Graph session' {
        BeforeEach {
            [AppDomain]::CurrentDomain.SetData('WinTuner.GUI.DelegatedGraphContext', $null)
            Mock Get-WtToken { 'test-access-token' }
            Mock Get-MgContext { $null }
            Mock Invoke-MgGraphRequest { throw 'A second Graph auth client must not be used' }
            Mock Invoke-RestMethod {
                param($Uri)
                if ($Uri.AbsolutePath -eq '/v1.0/me') {
                    return [pscustomobject]@{ userPrincipalName = 'admin@contoso.com' }
                }
                if ($Uri.AbsolutePath -eq '/v1.0/organization') {
                    return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' }) }
                }
                return [pscustomobject]@{ value = @() }
            }
        }

        AfterEach {
            [AppDomain]::CurrentDomain.SetData('WinTuner.GUI.DelegatedGraphContext', $null)
        }

        It 'verifies user and tenant using the WinTuner token without a second sign-in' {
            $context = Connect-WinTunerGraph -UserPrincipalName 'admin@contoso.com'
            $context.Account | Should -Be 'admin@contoso.com'
            $context.TenantId | Should -Be 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
            $context.AuthType | Should -Be 'Delegated'
            $context.PSObject.Properties.Name | Should -Not -Contain 'AccessToken'
            Should -Invoke Get-WtToken -Times 2 -Exactly -ParameterFilter { -not $ShowToken -and -not $DecodeToken }
            Should -Invoke Invoke-MgGraphRequest -Times 0
        }

        It 'blocks a different account and does not publish session metadata' {
            { Connect-WinTunerGraph -UserPrincipalName 'other@contoso.com' } | Should -Throw '*does not match*'
            [AppDomain]::CurrentDomain.GetData('WinTuner.GUI.DelegatedGraphContext') | Should -BeNullOrEmpty
        }

        It 'does not publish a session when tenant verification fails' {
            Mock Invoke-RestMethod { throw 'Forbidden' } -ParameterFilter { $Uri.AbsolutePath -eq '/v1.0/organization' }
            { Connect-WinTunerGraph -UserPrincipalName 'admin@contoso.com' } | Should -Throw '*Forbidden*'
            [AppDomain]::CurrentDomain.GetData('WinTuner.GUI.DelegatedGraphContext') | Should -BeNullOrEmpty
        }

        It 'retrieves tokens through the provider for each later request and preserves PATCH content' {
            $null = Connect-WinTunerGraph -UserPrincipalName 'admin@contoso.com'
            $script:observedHeaders = [System.Collections.Generic.List[string]]::new()
            $script:observedBodies = [System.Collections.Generic.List[string]]::new()
            Mock Get-WtToken { 'renewed-test-token' }
            Mock Invoke-RestMethod {
                param($Headers, $Body)
                $script:observedHeaders.Add([string]$Headers.Authorization)
                $script:observedBodies.Add([string]$Body)
                [pscustomobject]@{ value = @() }
            }
            $null = Invoke-WinTunerGraphRequest -Uri 'https://graph.microsoft.com/v1.0/deviceManagement/detectedApps'
            $null = Invoke-WinTunerGraphRequest -Uri 'https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps/app/assignments/assignment' -Method PATCH -Body '{"settings":{}}'
            $script:observedHeaders | Should -Be @('Bearer renewed-test-token', 'Bearer renewed-test-token')
            $script:observedBodies[1] | Should -Be '{"settings":{}}'
            Should -Invoke Invoke-RestMethod -Times 2 -Exactly -ParameterFilter { $MaximumRedirection -eq 0 -and $Uri.AbsolutePath -like '/v1.0/device*' }
            Should -Invoke Invoke-MgGraphRequest -Times 0
        }

        It 'rejects an external pagination URL before requesting a token' {
            { Invoke-WinTunerDelegatedGraphRequest -Uri 'https://example.com/token-leak' } | Should -Throw '*graph.microsoft.com*'
            { Invoke-WinTunerDelegatedGraphRequest -Uri 'http://graph.microsoft.com/v1.0/me' } | Should -Throw '*graph.microsoft.com*'
            Should -Invoke Get-WtToken -Times 0
            Should -Invoke Invoke-RestMethod -Times 0
        }

        It 'reuses identity metadata across a new runspace and clears it on logout' {
            $null = Connect-WinTunerGraph -UserPrincipalName 'admin@contoso.com'
            $modulePath = $ExecutionContext.SessionState.Module.Path
            $worker = [powershell]::Create()
            try {
                $null = $worker.AddScript('param($ModulePath) Import-Module $ModulePath -Force; (Confirm-WinTunerGraphContext -UserPrincipalName "admin@contoso.com").TenantId').AddArgument($modulePath)
                $result = @($worker.Invoke())
                $worker.HadErrors | Should -BeFalse
                $result[0] | Should -Be 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
            }
            finally { $worker.Dispose() }
            Disconnect-WinTunerGraph
            [AppDomain]::CurrentDomain.GetData('WinTuner.GUI.DelegatedGraphContext') | Should -BeNullOrEmpty
            { Confirm-WinTunerGraphContext -UserPrincipalName 'admin@contoso.com' } | Should -Throw '*No Microsoft Graph*'
        }

        It 'keeps app-only requests on the existing Graph SDK path' {
            Mock Invoke-MgGraphRequest { [pscustomobject]@{ value = @() } }
            $null = Invoke-WinTunerGraphRequest -Uri 'https://graph.microsoft.com/v1.0/deviceManagement/detectedApps'
            Should -Invoke Invoke-MgGraphRequest -Times 1 -Exactly
            Should -Invoke Get-WtToken -Times 0
        }
    }

    AfterAll {
        [AppDomain]::CurrentDomain.SetData('WinTuner.GUI.DelegatedGraphContext', $null)
        Remove-Item Function:script:Get-WtToken, Function:script:Get-MgContext, Function:script:Disconnect-MgGraph, Function:script:Invoke-MgGraphRequest -ErrorAction SilentlyContinue
    }
}
