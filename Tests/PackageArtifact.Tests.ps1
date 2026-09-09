BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\Modules\WinTuner.Core.psm1'
    Import-Module $modulePath -Force
}

Describe 'Test-WinTunerPackageArtifact' {
    BeforeEach {
        $rootPath = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $packageId = 'Contoso.App'
        $version = '1.2.3'
        $packagePath = Join-Path (Join-Path $rootPath $packageId) $version
        New-Item -ItemType Directory -Path $packagePath -Force | Out-Null
    }

    It 'accepts metadata that references an existing non-empty IntuneWin file' {
        '{"fileName":"Contoso.App.intunewin","displayVersion":"1.2.3"}' | Set-Content -LiteralPath (Join-Path $packagePath 'win32LobApp.json')
        'package-content' | Set-Content -LiteralPath (Join-Path $packagePath 'Contoso.App.intunewin') -NoNewline

        $result = Test-WinTunerPackageArtifact -RootPackageFolder $rootPath -PackageId $packageId -Version $version

        $result.IsValid | Should -BeTrue
        $result.ReasonCode | Should -Be 'Valid'
        $result.FileName | Should -Be 'Contoso.App.intunewin'
        $result.DisplayVersion | Should -Be '1.2.3'
    }

    It 'rejects a package without metadata' {
        $result = Test-WinTunerPackageArtifact -RootPackageFolder $rootPath -PackageId $packageId -Version $version

        $result.IsValid | Should -BeFalse
        $result.ReasonCode | Should -Be 'MetadataNotFound'
    }

    It 'rejects malformed metadata' {
        '{not-json' | Set-Content -LiteralPath (Join-Path $packagePath 'win32LobApp.json')

        $result = Test-WinTunerPackageArtifact -RootPackageFolder $rootPath -PackageId $packageId -Version $version

        $result.IsValid | Should -BeFalse
        $result.ReasonCode | Should -Be 'InvalidMetadata'
    }

    It 'rejects a package root that does not exist' {
        $missingRoot = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))

        $result = Test-WinTunerPackageArtifact -RootPackageFolder $missingRoot -PackageId $packageId -Version $version

        $result.IsValid | Should -BeFalse
        $result.ReasonCode | Should -Be 'RootNotFound'
    }

    It 'rejects metadata without a package filename' {
        '{"displayName":"Contoso App","displayVersion":"1.2.3"}' | Set-Content -LiteralPath (Join-Path $packagePath 'win32LobApp.json')

        $result = Test-WinTunerPackageArtifact -RootPackageFolder $rootPath -PackageId $packageId -Version $version

        $result.IsValid | Should -BeFalse
        $result.ReasonCode | Should -Be 'MissingFileName'
    }
    It 'rejects metadata without a display version' {
        '{"fileName":"Contoso.App.intunewin"}' | Set-Content -LiteralPath (Join-Path $packagePath 'win32LobApp.json')
        'package-content' | Set-Content -LiteralPath (Join-Path $packagePath 'Contoso.App.intunewin') -NoNewline

        $result = Test-WinTunerPackageArtifact -RootPackageFolder $rootPath -PackageId $packageId -Version $version

        $result.IsValid | Should -BeFalse
        $result.ReasonCode | Should -Be 'MissingDisplayVersion'
    }

    It 'rejects metadata for a different version' {
        '{"fileName":"Contoso.App.intunewin","displayVersion":"1.2.2"}' | Set-Content -LiteralPath (Join-Path $packagePath 'win32LobApp.json')
        'package-content' | Set-Content -LiteralPath (Join-Path $packagePath 'Contoso.App.intunewin') -NoNewline

        $result = Test-WinTunerPackageArtifact -RootPackageFolder $rootPath -PackageId $packageId -Version $version

        $result.IsValid | Should -BeFalse
        $result.ReasonCode | Should -Be 'VersionMismatch'
    }

    It 'rejects a metadata filename that escapes the package directory' {
        '{"fileName":"..\\outside.intunewin","displayVersion":"1.2.3"}' | Set-Content -LiteralPath (Join-Path $packagePath 'win32LobApp.json')

        $result = Test-WinTunerPackageArtifact -RootPackageFolder $rootPath -PackageId $packageId -Version $version

        $result.IsValid | Should -BeFalse
        $result.ReasonCode | Should -Be 'InvalidFileName'
    }

    It 'requires the exact IntuneWin file referenced by metadata' {
        '{"fileName":"expected.intunewin","displayVersion":"1.2.3"}' | Set-Content -LiteralPath (Join-Path $packagePath 'win32LobApp.json')
        'other-package' | Set-Content -LiteralPath (Join-Path $packagePath 'other.intunewin') -NoNewline

        $result = Test-WinTunerPackageArtifact -RootPackageFolder $rootPath -PackageId $packageId -Version $version

        $result.IsValid | Should -BeFalse
        $result.ReasonCode | Should -Be 'IntuneWinNotFound'
    }

    It 'rejects an empty referenced IntuneWin file' {
        '{"fileName":"empty.intunewin","displayVersion":"1.2.3"}' | Set-Content -LiteralPath (Join-Path $packagePath 'win32LobApp.json')
        [System.IO.File]::WriteAllBytes((Join-Path $packagePath 'empty.intunewin'), [byte[]]@())

        $result = Test-WinTunerPackageArtifact -RootPackageFolder $rootPath -PackageId $packageId -Version $version

        $result.IsValid | Should -BeFalse
        $result.ReasonCode | Should -Be 'IntuneWinEmpty'
    }

    It 'rejects an unsafe package identifier' {
        $result = Test-WinTunerPackageArtifact -RootPackageFolder $rootPath -PackageId '..\Contoso.App' -Version $version

        $result.IsValid | Should -BeFalse
        $result.ReasonCode | Should -Be 'InvalidPackageId'
    }

    It 'rejects a non-IntuneWin metadata filename' {
        '{"fileName":"Contoso.App.zip","displayVersion":"1.2.3"}' | Set-Content -LiteralPath (Join-Path $packagePath 'win32LobApp.json')
        'package-content' | Set-Content -LiteralPath (Join-Path $packagePath 'Contoso.App.zip') -NoNewline

        $result = Test-WinTunerPackageArtifact -RootPackageFolder $rootPath -PackageId $packageId -Version $version

        $result.IsValid | Should -BeFalse
        $result.ReasonCode | Should -Be 'InvalidFileType'
    }
}
