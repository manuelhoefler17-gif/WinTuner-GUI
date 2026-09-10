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
    It 'validates a package artifact before every direct GUI deployment call' {
        $deployCommands = @($script:guiAst.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.CommandAst] -and
            $node.GetCommandName() -eq 'Deploy-WtWin32App'
        }, $true))

        $deployCommands.Count | Should -Be 2

        foreach ($deployCommand in $deployCommands) {
            $scope = $deployCommand.Parent
            while ($scope -and $scope -isnot [System.Management.Automation.Language.ScriptBlockAst]) {
                $scope = $scope.Parent
            }

            $validationCommands = @($scope.FindAll({
                param($node)
                $node -is [System.Management.Automation.Language.CommandAst] -and
                $node.GetCommandName() -eq 'Test-WinTunerPackageArtifact'
            }, $true) | Where-Object {
                $_.Extent.StartOffset -lt $deployCommand.Extent.StartOffset
            })

            $validationCommands.Count | Should -BeGreaterThan 0 -Because (
                "deployment at line $($deployCommand.Extent.StartLineNumber) must validate its exact package first"
            )
        }
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
