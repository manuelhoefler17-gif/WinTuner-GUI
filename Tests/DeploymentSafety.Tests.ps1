BeforeAll {
    $guiPath = Join-Path $PSScriptRoot '..\WinTuner_GUI.ps1'
    $tokens = $null
    $parseErrors = $null
    $script:guiAst = [System.Management.Automation.Language.Parser]::ParseFile(
        $guiPath,
        [ref]$tokens,
        [ref]$parseErrors
    )

    if ($parseErrors.Count -gt 0) {
        throw ($parseErrors | ForEach-Object Message | Out-String)
    }
}

Describe 'Deployment artifact safety' {
    It 'routes destructive tenant work through isolated safety modules without GUI-thread calls' {
        $guiText = $script:guiAst.Extent.Text
        $directDeployCommands = @($script:guiAst.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.CommandAst] -and
            $node.GetCommandName() -eq 'Deploy-WtWin32App'
        }, $true))
        $directRemovalCommands = @($script:guiAst.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.CommandAst] -and
            $node.GetCommandName() -eq 'Remove-WtWin32App'
        }, $true))

        $directDeployCommands | Should -HaveCount 0
        $directRemovalCommands | Should -HaveCount 0
        $guiText | Should -Match 'Invoke-WinTunerDiscoveryDeployment'
        $guiText | Should -Match 'Invoke-WinTunerSupersededRemoval'
        $guiText | Should -Match 'Invoke-WinTunerDiscoveryScan'
        $guiText | Should -Match 'Invoke-WinTunerConnectionVerification'
        $guiText | Should -Not -Match '\[System\.Windows\.Forms\.Application\]::DoEvents'

        $deploymentModule = Get-Content (Join-Path $PSScriptRoot '..\Modules\WinTuner.DiscoveryDeployment.psm1') -Raw
        $validationOffset = $deploymentModule.IndexOf('Test-WinTunerPackageArtifact', [StringComparison]::Ordinal)
        $deploymentOffset = $deploymentModule.IndexOf('Deploy-WtWin32App', [StringComparison]::Ordinal)
        $validationOffset | Should -BeGreaterThan -1
        $deploymentOffset | Should -BeGreaterThan $validationOffset
    }
    It 'routes both update actions through the isolated update module' {
        $updateFunction = @($script:guiAst.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $node.Name -eq 'Start-WinTunerAppUpdateBatch'
        }, $true))

        $updateFunction | Should -HaveCount 1
        $updateFunction[0].Extent.Text | Should -Match 'Invoke-WinTunerAppUpdateBatch'

        $guiText = $script:guiAst.Extent.Text
        $guiText | Should -Match '(?s)\$updateSelectedButton\.Add_Click.*?Start-WinTunerAppUpdateBatch.*?-Mode Checked'
        $guiText | Should -Match '(?s)\$updateAllButton\.Add_Click.*?Start-WinTunerAppUpdateBatch.*?-Mode All'
        $guiText | Should -Not -Match 'Invoke-AppUpdateBatch'

        $discoveryStateFunction = @($script:guiAst.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $node.Name -eq 'Update-DiscoveryActionState'
        }, $true))
        $discoveryStateFunction | Should -HaveCount 1
        $discoveryStateFunction[0].Extent.Text | Should -Match 'isUpdateOperationActive'
    }
    It 'routes the main package upload through click-time validation and the upload safety module' {
        $uploadFunction = @($script:guiAst.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $node.Name -eq 'Start-WinTunerPackageUpload'
        }, $true))

        $uploadFunction | Should -HaveCount 1
        $uploadFunction[0].Extent.Text | Should -Match 'Test-WinTunerPackageArtifact'
        $uploadFunction[0].Extent.Text | Should -Match 'Invoke-WinTunerPackageUpload'
    }
}
