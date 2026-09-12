BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\Modules\WinTuner.Authentication.psm1') -Force
}

Describe 'WinTuner authentication configuration' {
    It 'accepts a valid interactive user and rejects an invalid one' {
        (Test-WinTunerAuthenticationConfiguration -Mode Interactive -UserPrincipalName 'admin@contoso.com').IsValid | Should -BeTrue
        (Test-WinTunerAuthenticationConfiguration -Mode Interactive -UserPrincipalName 'invalid').IsValid | Should -BeFalse
    }

    It 'validates client-secret fields without requiring a stored secret' {
        $base = @{
            Mode = 'ClientSecret'
            TenantId = 'contoso.onmicrosoft.com'
            ClientId = '11111111-1111-1111-1111-111111111111'
        }
        (Test-WinTunerAuthenticationConfiguration @base).IsValid | Should -BeTrue
        (Test-WinTunerAuthenticationConfiguration @base -RequireSecret).IsValid | Should -BeFalse
        (Test-WinTunerAuthenticationConfiguration @base -RequireSecret -ClientSecret 'temporary-value').IsValid | Should -BeTrue
    }

    It 'builds the expected WinTuner client-secret parameters' {
        $parameters = New-WinTunerModuleConnectionParameters -Mode ClientSecret -TenantId tenant.example -ClientId '11111111-1111-1111-1111-111111111111' -ClientSecret 'temporary-value'
        $parameters.TenantId | Should -Be 'tenant.example'
        $parameters.ClientId | Should -Be '11111111-1111-1111-1111-111111111111'
        $parameters.ClientSecret | Should -Be 'temporary-value'
        $parameters.Keys | Should -Not -Contain 'Username'
    }

    It 'normalizes and validates certificate thumbprints' {
        $thumbprint = 'AA BB CC DD EE FF 00 11 22 33 44 55 66 77 88 99 AA BB CC DD'
        $parameters = New-WinTunerModuleConnectionParameters -Mode Certificate -TenantId tenant.example -ClientId '22222222-2222-2222-2222-222222222222' -CertificateThumbprint $thumbprint
        $parameters.ClientCertificateThumbprint | Should -Be 'AABBCCDDEEFF00112233445566778899AABBCCDD'
        (Test-WinTunerAuthenticationConfiguration -Mode Certificate -TenantId tenant.example -ClientId '22222222-2222-2222-2222-222222222222' -CertificateThumbprint 'invalid').IsValid | Should -BeFalse
    }

    It 'returns only sanitized configuration metadata' {
        $result = Test-WinTunerAuthenticationConfiguration -Mode ClientSecret -TenantId tenant.example -ClientId '11111111-1111-1111-1111-111111111111' -ClientSecret 'do-not-return' -RequireSecret
        $result.PSObject.Properties.Name | Should -Not -Contain 'ClientSecret'
        ($result | ConvertTo-Json -Compress) | Should -Not -Match 'do-not-return'
    }

    It 'rejects an empty certificate lookup without touching a certificate store item' {
        $result = Test-WinTunerCertificateAvailable -Thumbprint ' '
        $result.IsAvailable | Should -BeFalse
        $result.Certificate | Should -BeNullOrEmpty
    }

    It 'redacts a client secret repeated by an upstream error' {
        $message = Protect-WinTunerAuthenticationErrorMessage -Message 'Login failed for secret-value.' -Secret 'secret-value'
        $message | Should -Be 'Login failed for [REDACTED].'
        $message | Should -Not -Match 'secret-value'
    }

    It 'protects a client secret for the current Windows user without exposing plaintext' {
        $plainText = 'test-client-secret-value'
        $protected = Protect-WinTunerClientSecretForCurrentUser -ClientSecret $plainText
        $protected | Should -Not -BeNullOrEmpty
        $protected | Should -Not -Match ([regex]::Escape($plainText))
        (Unprotect-WinTunerClientSecretForCurrentUser -ProtectedClientSecret $protected) | Should -Be $plainText
    }

    It 'rejects an invalid protected client secret with a safe recovery message' {
        { Unprotect-WinTunerClientSecretForCurrentUser -ProtectedClientSecret 'not-valid-base64' } | Should -Throw '*save it again*'
    }
}
