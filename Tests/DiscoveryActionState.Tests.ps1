BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\Modules\WinTuner.Core.psm1'
    Import-Module $modulePath -Force
}

Describe 'Get-WinTunerDiscoveryActionState' {
    It 'disables every discovery action while disconnected' {
        $state = Get-WinTunerDiscoveryActionState -Connected:$false -IsScanning:$false -CancelRequested:$false -IsDeploying:$false -ResultCount 3 -CheckedCount 1

        $state.CanScan | Should -BeFalse
        $state.CanDeploy | Should -BeFalse
        $state.CanExport | Should -BeFalse
        $state.CanCheckAll | Should -BeFalse
        $state.CanUncheckAll | Should -BeFalse
    }

    It 'enables only scanning before results exist' {
        $state = Get-WinTunerDiscoveryActionState -Connected:$true -IsScanning:$false -CancelRequested:$false -IsDeploying:$false -ResultCount 0 -CheckedCount 0

        $state.CanScan | Should -BeTrue
        $state.CanDeploy | Should -BeFalse
        $state.CanExport | Should -BeFalse
        $state.CanCheckAll | Should -BeFalse
        $state.CanUncheckAll | Should -BeFalse
    }

    It 'enables result actions but not deployment before a selection' {
        $state = Get-WinTunerDiscoveryActionState -Connected:$true -IsScanning:$false -CancelRequested:$false -IsDeploying:$false -ResultCount 3 -CheckedCount 0

        $state.CanScan | Should -BeTrue
        $state.CanDeploy | Should -BeFalse
        $state.CanExport | Should -BeTrue
        $state.CanCheckAll | Should -BeTrue
        $state.CanUncheckAll | Should -BeFalse
    }

    It 'enables deployment when at least one result is selected' {
        $state = Get-WinTunerDiscoveryActionState -Connected:$true -IsScanning:$false -CancelRequested:$false -IsDeploying:$false -ResultCount 3 -CheckedCount 1

        $state.CanDeploy | Should -BeTrue
        $state.CanExport | Should -BeTrue
        $state.CanCheckAll | Should -BeTrue
        $state.CanUncheckAll | Should -BeTrue
    }

    It 'disables Check All when every result is selected' {
        $state = Get-WinTunerDiscoveryActionState -Connected:$true -IsScanning:$false -CancelRequested:$false -IsDeploying:$false -ResultCount 3 -CheckedCount 3

        $state.CanCheckAll | Should -BeFalse
        $state.CanUncheckAll | Should -BeTrue
        $state.CanDeploy | Should -BeTrue
    }

    It 'leaves only the scan button available as Cancel Scan while scanning' {
        $state = Get-WinTunerDiscoveryActionState -Connected:$true -IsScanning:$true -CancelRequested:$false -IsDeploying:$false -ResultCount 3 -CheckedCount 1

        $state.CanScan | Should -BeTrue
        $state.CanDeploy | Should -BeFalse
        $state.CanExport | Should -BeFalse
        $state.CanCheckAll | Should -BeFalse
        $state.CanUncheckAll | Should -BeFalse
    }

    It 'disables every discovery action after cancellation is requested' {
        $state = Get-WinTunerDiscoveryActionState -Connected:$true -IsScanning:$true -CancelRequested:$true -IsDeploying:$false -ResultCount 3 -CheckedCount 1

        $state.CanScan | Should -BeFalse
        $state.CanDeploy | Should -BeFalse
        $state.CanExport | Should -BeFalse
        $state.CanCheckAll | Should -BeFalse
        $state.CanUncheckAll | Should -BeFalse
    }

    It 'disables every discovery action during deployment' {
        $state = Get-WinTunerDiscoveryActionState -Connected:$true -IsScanning:$false -CancelRequested:$false -IsDeploying:$true -ResultCount 3 -CheckedCount 1

        $state.CanScan | Should -BeFalse
        $state.CanDeploy | Should -BeFalse
        $state.CanExport | Should -BeFalse
        $state.CanCheckAll | Should -BeFalse
        $state.CanUncheckAll | Should -BeFalse
    }

    It 'disables discovery actions while another workflow is active' {
        $state = Get-WinTunerDiscoveryActionState -Connected:$true -IsScanning:$false -CancelRequested:$false -IsDeploying:$false -IsOtherOperationActive:$true -ResultCount 3 -CheckedCount 1

        $state.CanScan | Should -BeFalse
        $state.CanDeploy | Should -BeFalse
        $state.CanExport | Should -BeFalse
        $state.CanCheckAll | Should -BeFalse
        $state.CanUncheckAll | Should -BeFalse
    }
}
Describe 'GUI Discovery filter layout wiring' {
    It 'repositions the filter controls whenever the Discovery tab is resized' {
        $guiPath = Join-Path $PSScriptRoot '..\WinTuner_GUI.ps1'
        $tokens = $null
        $parseErrors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($guiPath, [ref]$tokens, [ref]$parseErrors)
        $parseErrors.Count | Should -Be 0

        $functions = @($ast.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $node.Name -eq 'Update-WinTunerDiscoveryFilterLayout'
        }, $true))

        $functions.Count | Should -Be 1
        $layoutText = $functions[0].Extent.Text
        $layoutText | Should -Match '\$discoveredAppSearchBox\.Location\s*='
        $layoutText | Should -Match '\$discoveredPublisherBox\.Location\s*='
        $layoutText | Should -Match '\$discoveredSortBox\.Location\s*='
        [System.IO.File]::ReadAllText($guiPath) | Should -Match '\$tabDiscovered\.Add_Resize\(\{\s*Update-WinTunerDiscoveryFilterLayout\s*\}\)'
    }
}

Describe 'GUI Discovery force-refresh wiring' {
    BeforeAll {
        $guiPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'WinTuner_GUI.ps1'
        $tokens = $null
        $parseErrors = $null
        $script:forceRefreshGuiAst = [System.Management.Automation.Language.Parser]::ParseFile(
            $guiPath,
            [ref]$tokens,
            [ref]$parseErrors
        )
        $parseErrors | Should -HaveCount 0
        $script:forceRefreshGuiText = [System.IO.File]::ReadAllText($guiPath)
    }

    It 'forwards the selected refresh mode into the Discovery worker' {
        $functions = @($script:forceRefreshGuiAst.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $node.Name -eq 'Start-WinTunerDiscoveryScan'
        }, $true))

        $functions | Should -HaveCount 1
        $functionText = $functions[0].Extent.Text
        $functionText.Contains('$forceGraphRefresh = [bool]$forceFreshDiscoveryCheckBox.Checked') | Should -BeTrue
        $functionText.Contains('Get-WinTunerDetectedApps -PageSize 500 -MaxPages 1000 -ForceRefresh') | Should -BeTrue
        $functionText.Contains('.AddArgument($forceGraphRefresh)') | Should -BeTrue
    }

    It 'disables the refresh option while a Discovery scan is active' {
        $functions = @($script:forceRefreshGuiAst.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $node.Name -eq 'Update-DiscoveryActionState'
        }, $true))

        $functions | Should -HaveCount 1
        $functions[0].Extent.Text.Contains('$forceFreshDiscoveryCheckBox.Enabled = ($state.CanScan -and -not $script:discoveryScanRunning)') | Should -BeTrue
        $script:forceRefreshGuiText.Contains('Force fresh Graph data (ignore detected-app cache)') | Should -BeTrue
    }
}
