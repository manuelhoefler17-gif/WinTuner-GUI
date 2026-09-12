BeforeAll {
    $scriptPath = Join-Path $PSScriptRoot 'Run-TenantE2E.ps1'
    $tokens = $null
    $parseErrors = $null
    $script:e2eAst = [System.Management.Automation.Language.Parser]::ParseFile(
        $scriptPath,
        [ref]$tokens,
        [ref]$parseErrors
    )
    $parseErrors | Should -HaveCount 0
    $script:e2eText = [System.IO.File]::ReadAllText($scriptPath)
}

Describe 'Read-only tenant E2E runner' {
    It 'resolves all repository helper modules from the Modules directory' {
        $script:e2eText.Contains('ModulesWinTuner') | Should -BeFalse
        $repositoryRoot = Split-Path -Parent $PSScriptRoot
        Test-Path -LiteralPath (Join-Path (Join-Path $repositoryRoot 'Modules') 'WinTuner.Connection.psm1') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path (Join-Path $repositoryRoot 'Modules') 'WinTuner.Intune.psm1') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path (Join-Path $repositoryRoot 'Modules') 'WinTuner.Authentication.psm1') | Should -BeTrue
    }
    It 'covers both tenant APIs and always disconnects them' {
        $commandNames = @($script:e2eAst.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.CommandAst]
        }, $true) | ForEach-Object { $_.GetCommandName() })

        $commandNames | Should -Contain 'Connect-WtWinTuner'
        $commandNames | Should -Contain 'Invoke-WinTunerConnectionVerification'
        $commandNames | Should -Contain 'Connect-WinTunerGraph'
        $commandNames | Should -Contain 'Get-WinTunerDetectedApps'
        $commandNames | Should -Contain 'Disconnect-WinTunerGraph'
        $commandNames | Should -Contain 'Disconnect-WtWinTuner'
    }

    It 'contains no package deployment or tenant removal commands' {
        $commandNames = @($script:e2eAst.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.CommandAst]
        }, $true) | ForEach-Object { $_.GetCommandName() })

        $forbiddenCommands = @(
            'Deploy-WtWin32App',
            'New-WtWin32App',
            'Remove-WtWin32App',
            'Invoke-WinTunerPackageBuild',
            'Invoke-WinTunerPackageUpload',
            'Invoke-WinTunerDiscoveryDeployment',
            'Invoke-WinTunerSupersededRemoval'
        )
        @($commandNames | Where-Object { $_ -in $forbiddenCommands }) | Should -HaveCount 0
    }
}

Describe 'Tenant E2E Discovery cache validation' {
    It 'supports a two-pass fresh-to-cached Discovery validation' {
        $script:e2eText | Should -Match '\[switch\]\$ValidateDiscoveryCache'
        $script:e2eText | Should -Match 'Get-WinTunerDetectedApps @detectedParameters -ForceRefresh'
        $script:e2eText | Should -Match 'Clear-WinTunerDetectedAppsCache'
        $script:e2eText | Should -Match '\$freshDetectedResult\.LimitReached'
        $script:e2eText | Should -Match 'if \(-not \$cachedDetectedResult\.FromCache\)'
        $script:e2eText | Should -Match 'fresh and cached counts differ'
        $script:e2eText | Should -Match 'CacheReuseVerified = \[bool\]\$ValidateDiscoveryCache'
        $script:e2eText | Should -Match 'TenantWritesPerformed = 0'
    }

    It 'rejects contradictory Discovery refresh switches' {
        $script:e2eText | Should -Match '\$ForceFreshDiscovery -and \$ValidateDiscoveryCache'
        $script:e2eText | Should -Match 'Use either ForceFreshDiscovery or ValidateDiscoveryCache'
    }
}
Describe 'Tenant E2E app-only authentication' {
    It 'supports interactive, client-secret, and certificate modes' {
        $script:e2eText | Should -Match "\[ValidateSet\('Interactive', 'ClientSecret', 'Certificate'\)\]"
        $script:e2eText | Should -Match '\[securestring\]\$ClientSecret'
        $script:e2eText | Should -Match 'Test-WinTunerCertificateAvailable'
        $script:e2eText | Should -Match 'New-WinTunerModuleConnectionParameters'
    }

    It 'clears only the dedicated WinTuner app-only cache before connecting' {
        $script:e2eText | Should -Match 'Clear-WinTunerAppOnlyTokenCache'
        $script:e2eText | Should -Not -Match 'mg\.msal\.cache'
        $script:e2eText.IndexOf('Clear-WinTunerAppOnlyTokenCache') | Should -BeLessThan $script:e2eText.IndexOf('Connect-WtWinTuner @connectionParameters')
    }

    It 'checks both required Application permissions through read-only endpoints' {
        $script:e2eText | Should -Match 'Invoke-WinTunerGraphPermissionPreflight'
        $script:e2eText | Should -Match 'deviceAppManagement/mobileApps\?\$top=1'
        $script:e2eText | Should -Match 'deviceManagement/detectedApps\?\$top=1'
        $script:e2eText | Should -Match 'PermissionPreflightPassed'
    }

    It 'converts and clears the SecureString without writing the plaintext secret' {
        $script:e2eText | Should -Match 'SecureStringToBSTR'
        $script:e2eText | Should -Match 'ZeroFreeBSTR'
        $script:e2eText | Should -Match '\$plainClientSecret\s*=\s*\$null'
        $script:e2eText | Should -Not -Match 'Write-(Host|Output|Verbose).*plainClientSecret'
        $script:e2eText | Should -Not -Match 'ClientSecret\s*=\s*\$plainClientSecret'
    }
}
