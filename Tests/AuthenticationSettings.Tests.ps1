BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\Modules\WinTuner.Settings.psm1') -Force
    Import-Module (Join-Path $PSScriptRoot '..\Modules\WinTuner.Authentication.psm1') -Force
}

Describe 'Authentication settings persistence' {
    It 'defaults to interactive authentication without app credentials' {
        $settings = New-WinTunerDefaultSettings
        $settings.AuthenticationMode | Should -Be 'Interactive'
        $settings.EntraTenantId | Should -BeNullOrEmpty
        $settings.EntraClientId | Should -BeNullOrEmpty
        $settings.EntraCertificateThumbprint | Should -BeNullOrEmpty
        $settings.EntraClientSecretProtected | Should -BeNullOrEmpty
        $settings.Keys | Should -Not -Contain 'ClientSecret'
    }

    It 'round-trips app-only configuration without a plaintext client secret' {
        $path = Join-Path $TestDrive 'settings.json'
        $settings = New-WinTunerDefaultSettings
        $settings.AuthenticationMode = 'Certificate'
        $settings.EntraTenantId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
        $settings.EntraClientId = '11111111-1111-1111-1111-111111111111'
        $settings.EntraCertificateThumbprint = 'AABBCCDDEEFF00112233445566778899AABBCCDD'

        (Export-WinTunerSettings -Settings $settings -Path $path) | Should -BeTrue
        $loaded = Import-WinTunerSettings -Path $path
        $loaded.AuthenticationMode | Should -Be 'Certificate'
        $loaded.EntraTenantId | Should -Be $settings.EntraTenantId
        $loaded.EntraClientId | Should -Be $settings.EntraClientId
        $loaded.EntraCertificateThumbprint | Should -Be $settings.EntraCertificateThumbprint
        $plainTextSecret = 'test-settings-secret-value'
        $settings.EntraClientSecretProtected = Protect-WinTunerClientSecretForCurrentUser -ClientSecret $plainTextSecret
        (Export-WinTunerSettings -Settings $settings -Path $path) | Should -BeTrue
        $loaded = Import-WinTunerSettings -Path $path
        (Unprotect-WinTunerClientSecretForCurrentUser -ProtectedClientSecret $loaded.EntraClientSecretProtected) | Should -Be $plainTextSecret
        ([System.IO.File]::ReadAllText($path)).Contains($plainTextSecret) | Should -BeFalse
        ([System.IO.File]::ReadAllText($path)).Contains('"ClientSecret":') | Should -BeFalse
    }

    It 'falls back to interactive mode for an unknown persisted mode' {
        $path = Join-Path $TestDrive 'invalid-settings.json'
        [System.IO.File]::WriteAllText($path, '{"AuthenticationMode":"Unknown"}')
        (Import-WinTunerSettings -Path $path).AuthenticationMode | Should -Be 'Interactive'
    }
}
