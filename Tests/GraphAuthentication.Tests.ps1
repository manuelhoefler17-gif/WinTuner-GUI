BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\Modules\WinTuner.Intune.psm1') -Force
    $script:intuneText = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot '..\Modules\WinTuner.Intune.psm1'))
}

Describe 'Microsoft Graph authentication context validation' {
    It 'accepts the requested interactive account' {
        $context = [pscustomobject]@{ Account = 'admin@contoso.com'; ClientId = 'ignored'; TenantId = 'tenant'; AuthType = 'Delegated' }
        (Get-WinTunerGraphContextValidation -AuthenticationMode Interactive -UserPrincipalName 'admin@contoso.com' -Context $context).IsValid | Should -BeTrue
    }

    It 'rejects a different interactive account' {
        $context = [pscustomobject]@{ Account = 'other@contoso.com'; ClientId = 'ignored'; TenantId = 'tenant'; AuthType = 'Delegated' }
        $result = Get-WinTunerGraphContextValidation -AuthenticationMode Interactive -UserPrincipalName 'admin@contoso.com' -Context $context
        $result.IsValid | Should -BeFalse
        $result.Reason | Should -Match 'does not match'
    }

    It 'accepts a matching app-only client and tenant' {
        $context = [pscustomobject]@{
            Account = $null
            ClientId = '11111111-1111-1111-1111-111111111111'
            TenantId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
            AuthType = 'AppOnly'
        }
        $result = Get-WinTunerGraphContextValidation -AuthenticationMode ClientSecret -TenantId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -ClientId '11111111-1111-1111-1111-111111111111' -Context $context
        $result.IsValid | Should -BeTrue
    }

    It 'rejects an app-only client or tenant mismatch' {
        $context = [pscustomobject]@{
            ClientId = '11111111-1111-1111-1111-111111111111'
            TenantId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
            AuthType = 'AppOnly'
        }
        (Get-WinTunerGraphContextValidation -AuthenticationMode Certificate -TenantId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -ClientId '22222222-2222-2222-2222-222222222222' -Context $context).IsValid | Should -BeFalse
        (Get-WinTunerGraphContextValidation -AuthenticationMode Certificate -TenantId 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb' -ClientId '11111111-1111-1111-1111-111111111111' -Context $context).IsValid | Should -BeFalse
    }

    It 'rejects a delegated context for an app-only request' {
        $context = [pscustomobject]@{
            ClientId = '11111111-1111-1111-1111-111111111111'
            TenantId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
            AuthType = 'Delegated'
        }
        (Get-WinTunerGraphContextValidation -AuthenticationMode ClientSecret -TenantId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -ClientId '11111111-1111-1111-1111-111111111111' -Context $context).IsValid | Should -BeFalse
    }

    It 'uses native Graph client-secret and certificate authentication parameters' {
        $script:intuneText | Should -Match '-ClientSecretCredential \$credential'
        $script:intuneText | Should -Match '-CertificateThumbprint \$normalizedThumbprint'
        $script:intuneText | Should -Match '-ContextScope Process'
        $script:intuneText | Should -Not -Match 'Write-(Host|Verbose|Output).*ClientSecret'
    }
}
