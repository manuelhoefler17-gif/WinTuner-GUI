BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\Modules\WinTuner.Core.psm1'
    Import-Module $modulePath -Force
}

Describe 'Test-WinTunerPackageRoot' {
    It 'accepts a normal package path that does not exist yet' {
        $path = Join-Path $TestDrive 'Packages\Nested'

        $result = Test-WinTunerPackageRoot -RootPackageFolder $path

        $result.IsValid | Should -BeTrue
        $result.ReasonCode | Should -Be 'Valid'
        $result.FullPath | Should -Be ([System.IO.Path]::GetFullPath($path))
    }

    It 'rejects an empty package path' {
        $result = Test-WinTunerPackageRoot -RootPackageFolder ''

        $result.IsValid | Should -BeFalse
        $result.ReasonCode | Should -Be 'InvalidRoot'
    }

    It 'rejects a path that points to a file' {
        $filePath = Join-Path $TestDrive 'package-root.txt'
        Set-Content -LiteralPath $filePath -Value 'not a directory'

        $result = Test-WinTunerPackageRoot -RootPackageFolder $filePath

        $result.IsValid | Should -BeFalse
        $result.ReasonCode | Should -Be 'RootIsFile'
    }

    It 'rejects a drive root' {
        $driveRoot = [System.IO.Path]::GetPathRoot($TestDrive)

        $result = Test-WinTunerPackageRoot -RootPackageFolder $driveRoot

        $result.IsValid | Should -BeFalse
        $result.ReasonCode | Should -Be 'ProtectedRoot'
    }

    It 'rejects protected Windows directories and their children' {
        $protectedPath = Join-Path ([Environment]::GetFolderPath('Windows')) 'Temp'

        $result = Test-WinTunerPackageRoot -RootPackageFolder $protectedPath

        $result.IsValid | Should -BeFalse
        $result.ReasonCode | Should -Be 'ProtectedRoot'
    }
}