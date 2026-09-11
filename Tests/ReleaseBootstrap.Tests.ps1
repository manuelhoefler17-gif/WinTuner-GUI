Describe 'Standalone release bootstrap manifest' {
    It 'includes every runtime module and worker shipped by the repository' {
        $repositoryRoot = Split-Path $PSScriptRoot -Parent
        $guiPath = Join-Path $repositoryRoot 'WinTuner_GUI.ps1'
        $tokens = $null
        $parseErrors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($guiPath, [ref]$tokens, [ref]$parseErrors)
        $parseErrors.Count | Should -Be 0

        $manifestAssignments = @($ast.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and
            $node.Left.Extent.Text -eq '$requiredReleaseFiles'
        }, $true))
        $manifestAssignments.Count | Should -Be 1

        $manifestFiles = @($manifestAssignments[0].Right.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.StringConstantExpressionAst]
        }, $true) | ForEach-Object { $_.Value.Replace('/', '\') } | Sort-Object)

        $runtimeFiles = @(
            Get-ChildItem -LiteralPath (Join-Path $repositoryRoot 'Modules') -Filter '*.psm1' -File |
                ForEach-Object { 'Modules\' + $_.Name }
            Get-ChildItem -LiteralPath (Join-Path $repositoryRoot 'Workers') -Filter '*.ps1' -File |
                ForEach-Object { 'Workers\' + $_.Name }
        ) | Sort-Object

        $manifestFiles | Should -Be $runtimeFiles
    }
}
