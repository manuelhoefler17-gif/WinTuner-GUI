BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\Modules\WinTuner.Settings.psm1') -Force
    Import-Module (Join-Path $PSScriptRoot '..\Modules\WinTuner.Authentication.psm1') -Force
    Import-Module (Join-Path $PSScriptRoot '..\Modules\WinTuner.TenantProfiles.psm1') -Force
    $script:tenantA = '11111111-1111-1111-1111-111111111111'
    $script:tenantB = '22222222-2222-2222-2222-222222222222'
    $script:client = '33333333-3333-3333-3333-333333333333'
    function New-TestProfile {
        param([string]$Name = 'Customer A', [string]$TenantId = $script:tenantA)
        New-WinTunerTenantProfile -Name $Name -TenantId $TenantId -AuthenticationMode Interactive -UserPrincipalName 'admin@example.com'
    }
}

Describe 'Actual login worker tenant boundary' {
    BeforeAll {
        $script:root = Split-Path $PSScriptRoot -Parent
        $tokens = $null; $errors = $null
        $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'WinTuner_GUI.ps1'), [ref]$tokens, [ref]$errors)
        $login = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Start-WinTunerLogin' }, $true)
        $assignment = $login.Find({ param($n) $n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$workerScript' }, $true)
        $script:loginWorker = $assignment.Right.Find({ param($n) $n -is [Management.Automation.Language.StringConstantExpressionAst] }, $true).Value
        function Invoke-LoginTenantFixture {
            param([string]$ExpectedTenant, [string]$ActualTenant, [string]$Mode = 'Interactive')
            $runner = [powershell]::Create()
            $fixture = @'
param($Root,$Worker,$Expected,$Actual,$Mode)
$ErrorActionPreference = 'Stop'
$global:tenantFixture = @{ Apps = 0; Preflight = 0; Actual = $Actual }
function Import-Module {
    [CmdletBinding()] param([string]$Name,[switch]$Force)
    if ($Name -match 'WinTuner\.TenantProfiles\.psm1$') { Microsoft.PowerShell.Core\Import-Module $Name -Force -Global }
}
function Connect-WinTunerGraph { [CmdletBinding()] param($AuthenticationMode,$UserPrincipalName) [pscustomobject]@{ TenantId = $global:tenantFixture.Actual } }
function Confirm-WinTunerGraphContext { [CmdletBinding()] param($AuthenticationMode,$TenantId,$ClientId) [pscustomobject]@{ TenantId = $global:tenantFixture.Actual } }
function Invoke-WinTunerGraphPermissionPreflight { param($GetManagedApps,$GetDetectedApps) $global:tenantFixture.Preflight++; [pscustomobject]@{ Succeeded = $true } }
function Invoke-WinTunerConnectionVerification { param($GetApps,$MaxAttempts,$RetryDelayMilliseconds) $global:tenantFixture.Apps++; [pscustomobject]@{ Succeeded = $true; Attempts = 1 } }
$message = ''
try { $null = & ([scriptblock]::Create($Worker)) $Root $Mode 'admin@example.com' $Actual '33333333-3333-3333-3333-333333333333' $Expected }
catch { $message = $_.Exception.Message }
[pscustomobject]@{ Error = $message; Apps = $global:tenantFixture.Apps; Preflight = $global:tenantFixture.Preflight } | ConvertTo-Json -Compress
'@
            try {
                $null = $runner.AddScript($fixture).AddArgument($root).AddArgument($loginWorker).AddArgument($ExpectedTenant).AddArgument($ActualTenant).AddArgument($Mode)
                $output = @($runner.Invoke())
                if ($output.Count -eq 0) { throw 'The login fixture returned no outcome.' }
                [string]$output[-1] | ConvertFrom-Json
            } finally { $runner.Dispose() }
        }
    }
    It 'rejects another tenant before any tenant inventory reads or app-only permission probes' -ForEach @(@{ Mode = 'Interactive' }, @{ Mode = 'ClientSecret' }, @{ Mode = 'Certificate' }) {
        $result = Invoke-LoginTenantFixture -ExpectedTenant $tenantA -ActualTenant $tenantB -Mode $Mode
        $result.Error | Should -Match 'does not match'
        $result.Apps | Should -Be 0
        $result.Preflight | Should -Be 0
    }
    It 'continues verification for the selected tenant and for legacy manual login' {
        foreach ($expected in @($tenantA, '')) {
            $result = Invoke-LoginTenantFixture -ExpectedTenant $expected -ActualTenant $tenantA
            $result.Error | Should -Be ''
            $result.Apps | Should -Be 1
        }
    }
}

Describe 'Tenant profile validation and storage' {
    It 'round-trips separate tenant configurations without credentials leaking between profiles' {
        $path = Join-Path $TestDrive 'profiles.json'
        $a = New-TestProfile
        $b = New-WinTunerTenantProfile -Name 'Customer B' -TenantId $tenantB -AuthenticationMode Certificate -ClientId $client -CertificateThumbprint ('A' * 40) -DefaultPackagePath 'C:\CustomerB'
        Set-WinTunerTenantProfile -Profile $a -Path $path
        Set-WinTunerTenantProfile -Profile $b -Path $path
        $profiles = @(Get-WinTunerTenantProfiles -Path $path)
        $profiles.Count | Should -Be 2
        $profiles[0].UserPrincipalName | Should -Be 'admin@example.com'
        $profiles[1].UserPrincipalName | Should -Be ''
        $profiles[1].CertificateThumbprint | Should -Be ('A' * 40)
        $profiles[1].DefaultPackagePath | Should -Be 'C:\CustomerB'
        $profiles[0] | Add-Member Token 'must-not-persist'
        Set-WinTunerTenantProfile -Profile $profiles[0] -Path $path
        Get-Content $path -Raw | Should -Not -Match 'must-not-persist|"Token"|"ClientSecret"'
        @(Get-ChildItem $TestDrive -Filter '*.tmp').Count | Should -Be 0
    }
    It 'refuses duplicate names and a tenant change to an existing profile without altering its file' {
        $path = Join-Path $TestDrive 'immutable.json'
        $a = New-TestProfile
        Set-WinTunerTenantProfile -Profile $a -Path $path
        $before = Get-Content $path -Raw
        { Set-WinTunerTenantProfile -Profile (New-TestProfile -Name 'customer a' -TenantId $tenantB) -Path $path } | Should -Throw '*already exists*'
        $a.TenantId = $tenantB
        { Set-WinTunerTenantProfile -Profile $a -Path $path } | Should -Throw '*another tenant*'
        Get-Content $path -Raw | Should -Be $before
    }
    It 'removes only the chosen profile' {
        $path = Join-Path $TestDrive 'remove.json'
        $a = New-TestProfile
        Set-WinTunerTenantProfile -Profile $a -Path $path
        Set-WinTunerTenantProfile -Profile (New-TestProfile -Name 'B' -TenantId $tenantB) -Path $path
        Set-WinTunerTenantProfile -Profile $a -Remove -Path $path
        $remaining = @(Get-WinTunerTenantProfiles -Path $path)
        $remaining.Count | Should -Be 1
        $remaining[0].Name | Should -Be 'B'
    }
    It 'blocks corrupt storage and does not overwrite it on save' {
        $path = Join-Path $TestDrive 'corrupt.json'
        Set-Content $path '{bad-json'
        { Get-WinTunerTenantProfiles -Path $path } | Should -Throw '*could not be read*'
        { Set-WinTunerTenantProfile -Profile (New-TestProfile) -Path $path } | Should -Throw '*could not be read*'
        (Get-Content $path -Raw).Trim() | Should -Be '{bad-json'
    }
    It 'rejects invalid tenants, usernames, certificates and package folders' {
        { New-TestProfile -TenantId ([guid]::Empty) } | Should -Throw
        { New-WinTunerTenantProfile -Name 'A' -TenantId $tenantA -AuthenticationMode Interactive -UserPrincipalName 'invalid' } | Should -Throw
        { New-WinTunerTenantProfile -Name 'A' -TenantId $tenantA -AuthenticationMode Certificate -ClientId $client -CertificateThumbprint 'invalid' } | Should -Throw
        { New-WinTunerTenantProfile -Name 'A' -TenantId $tenantA -AuthenticationMode Interactive -UserPrincipalName 'admin@example.com' -DefaultPackagePath '..\other' } | Should -Throw
    }
    It 'accepts only a protected client secret and never decrypts it during profile switching' {
        $protected = Protect-WinTunerClientSecretForCurrentUser -ClientSecret 'test-only-profile-secret'
        $profile = New-WinTunerTenantProfile -Name 'Secret' -TenantId $tenantA -AuthenticationMode ClientSecret -ClientId $client -EntraClientSecretProtected $protected
        $path = Join-Path $TestDrive 'secret.json'
        Set-WinTunerTenantProfile -Profile $profile -Path $path
        Get-Content $path -Raw | Should -Not -Match 'test-only-profile-secret'
        $loaded = @(Get-WinTunerTenantProfiles -Path $path)[0]
        $loaded.EntraClientSecretProtected | Should -Be $protected
        $settings = Get-WinTunerTenantProfileSettings -Settings (New-WinTunerDefaultSettings) -Profile $loaded
        $settings.EntraClientSecretProtected | Should -Be $protected
        { New-WinTunerTenantProfile -Name 'Secret' -TenantId $tenantA -AuthenticationMode ClientSecret -ClientId $client -EntraClientSecretProtected 'plaintext' } | Should -Throw
    }
}

Describe 'Tenant profile session boundaries' {
    It 'clears all previous app credentials on a switch and leaves unrelated settings intact' {
        $settings = New-WinTunerDefaultSettings
        $settings.EntraClientSecretProtected = 'previous-protected-value'
        $settings.EntraClientId = $client
        $settings.EntraCertificateThumbprint = 'previous-certificate'
        $settings.AutoCheckUpdates = $true
        $profile = New-TestProfile -Name 'B' -TenantId $tenantB
        $result = Get-WinTunerTenantProfileSettings -Settings $settings -Profile $profile
        $result.EntraTenantId | Should -Be $tenantB
        $result.EntraClientSecretProtected | Should -Be ''
        $result.EntraClientId | Should -Be ''
        $result.EntraCertificateThumbprint | Should -Be ''
        $result.ActiveTenantProfileId | Should -Be $profile.Id
        $result.AutoCheckUpdates | Should -BeTrue
        $settings.EntraClientSecretProtected | Should -Be 'previous-protected-value'
        $manual = Get-WinTunerTenantProfileSettings -Settings $result -Profile $null
        $manual.ActiveTenantProfileId | Should -Be ''
        $manual.LastUser | Should -Be ''
        $manual.EntraTenantId | Should -Be ''
    }
    It 'blocks both connected and busy profile switches, including manual configuration' {
        $settings = New-WinTunerDefaultSettings
        { Get-WinTunerTenantProfileSettings -Settings $settings -Profile (New-TestProfile) -Connected $true } | Should -Throw '*Sign out*'
        { Get-WinTunerTenantProfileSettings -Settings $settings -Profile $null -IsBusy $true } | Should -Throw '*Sign out*'
    }
    It 'rejects wrong or missing actual tenants and accepts only the selected tenant' {
        { Assert-WinTunerTenantProfileContext -ExpectedTenantId $tenantA -ActualTenantId $tenantB } | Should -Throw '*does not match*'
        { Assert-WinTunerTenantProfileContext -ExpectedTenantId $tenantA -ActualTenantId '' } | Should -Throw '*does not match*'
        { Assert-WinTunerTenantProfileContext -ExpectedTenantId $tenantA -ActualTenantId $tenantA } | Should -Not -Throw
        { Assert-WinTunerTenantProfileContext -ExpectedTenantId '' -ActualTenantId $tenantB } | Should -Not -Throw
    }
    It 'preserves legacy manual settings and persists the explicit profile choice' {
        $path = Join-Path $TestDrive 'settings.json'
        Set-Content $path '{"RememberMe":true,"LastUser":"admin@example.com"}'
        $settings = Import-WinTunerSettings -Path $path
        $settings.ActiveTenantProfileId | Should -Be ''
        $settings.LastUser | Should -Be 'admin@example.com'
        $profile = New-TestProfile
        $settings = Get-WinTunerTenantProfileSettings -Settings $settings -Profile $profile
        Export-WinTunerSettings -Settings $settings -Path $path | Should -BeTrue
        (Import-WinTunerSettings -Path $path).ActiveTenantProfileId | Should -Be $profile.Id
    }
}

Describe 'GUI profile restoration and identity display' {
    BeforeAll {
        $tokens = $null; $errors = $null
        $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\WinTuner_GUI.ps1'), [ref]$tokens, [ref]$errors)
        foreach ($name in @('Initialize-WinTunerActiveTenantProfile','Test-WinTunerProfileSwitchBusy','Get-WinTunerConfiguredAuthentication','Update-WinTunerTenantProfileSummary')) {
            $definition = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
            Invoke-Expression $definition.Extent.Text
        }
    }
    BeforeEach {
        $script:settings = New-WinTunerDefaultSettings
        $script:storedProfile = New-TestProfile
        $script:settings.ActiveTenantProfileId = $storedProfile.Id
        $script:usernameBox = [pscustomobject]@{ Text = '' }
        $script:pathBox = [pscustomobject]@{ Text = '' }
        $script:defaultPathTextBox = [pscustomobject]@{ Text = '' }
        $script:tenantProfileSummaryLabel = [pscustomobject]@{ Text = '' }
        $script:isConnected = $false
        foreach ($flag in @('isLoginOperationActive','isUpdateOperationActive','updateScanRunning','isSupersededOperationActive','discoveryScanRunning','isDiscoveryDeploymentActive','isPackageSearchActive','isVersionLookupActive','isPackageBuildActive','isPackageUploadActive')) { Set-Variable -Scope Script -Name $flag -Value $false }
        Mock Get-WinTunerTenantProfiles { $script:storedProfile }
    }
    It 'restores the explicit profile username even with Remember me off' {
        Initialize-WinTunerActiveTenantProfile
        $script:usernameBox.Text | Should -Be 'admin@example.com'
        $script:tenantProfileError | Should -Be ''
        $script:settings.RememberMe | Should -BeFalse
        $script:pathBox.Text | Should -Be 'C:\Temp'
        (Get-WinTunerConfiguredAuthentication).IsValid | Should -BeTrue
        Update-WinTunerTenantProfileSummary
        $script:tenantProfileSummaryLabel.Text | Should -Match $tenantA
    }
    It 'blocks login for a removed selected profile rather than using the previous credentials' {
        Mock Get-WinTunerTenantProfiles { @() }
        $script:usernameBox.Text = 'previous@example.com'
        Initialize-WinTunerActiveTenantProfile
        $script:activeTenantProfile | Should -BeNullOrEmpty
        $script:tenantProfileError | Should -Match 'no longer exists'
        (Get-WinTunerConfiguredAuthentication).IsValid | Should -BeFalse
    }
    It 'shows the verified session tenant and mode independently of edited configuration' {
        $script:isConnected = $true
        $script:currentTenantProfileName = 'Verified customer'
        $script:currentGraphTenantId = $tenantB
        $script:currentAuthenticationMode = 'Certificate'
        Update-WinTunerTenantProfileSummary
        $script:tenantProfileSummaryLabel.Text | Should -Match "Verified customer.*$tenantB.*Certificate"
    }
}
