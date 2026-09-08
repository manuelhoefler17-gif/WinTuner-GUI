BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\Modules\WinTuner.Core.psm1'
    Import-Module $modulePath -Force
}

Describe 'Test-IsNewerVersion' {
    It 'returns true when the latest version is newer' {
        Test-IsNewerVersion -Latest '1.2.4' -Current '1.2.3' | Should -BeTrue
    }

    It 'returns false when versions are equal' {
        Test-IsNewerVersion -Latest '1.2.3' -Current '1.2.3' | Should -BeFalse
    }

    It 'returns false when the latest version is older' {
        Test-IsNewerVersion -Latest '1.2.2' -Current '1.2.3' | Should -BeFalse
    }

    It 'handles different numeric component counts' {
        Test-IsNewerVersion -Latest '2.0' -Current '1.9.9.9' | Should -BeTrue
    }

    It 'handles common version suffixes by comparing the numeric prefix' {
        Test-IsNewerVersion -Latest '1.2.4-beta' -Current '1.2.3' | Should -BeTrue
    }

    It 'returns false for an empty latest version' {
        Test-IsNewerVersion -Latest '' -Current '1.2.3' | Should -BeFalse
    }

    It 'returns false for an unparseable version' {
        Test-IsNewerVersion -Latest 'unknown' -Current '1.2.3' | Should -BeFalse
    }
}
