BeforeAll {
    $script:intuneModuleText = [System.IO.File]::ReadAllText(
        (Join-Path $PSScriptRoot '..\Modules\WinTuner.Intune.psm1')
    )
}

Describe 'Detected-app cache metadata' {
    It 'calculates a non-negative cache age' {
        $script:intuneModuleText | Should -Match 'AgeMinutes = \[math\]::Max\(0, \$ageMinutes\)'
    }

    It 'returns timestamp and rounded age for a cache hit' {
        $script:intuneModuleText | Should -Match 'RetrievedAt = \$cachedResult\.Timestamp'
        $script:intuneModuleText | Should -Match 'AgeMinutes = \[math\]::Round\(\[double\]\$cachedResult\.AgeMinutes, 1\)'
    }

    It 'returns a retrieval timestamp and zero age for fresh Graph data' {
        $script:intuneModuleText | Should -Match '\$retrievedAt = \[datetime\]::UtcNow'
        $script:intuneModuleText | Should -Match 'RetrievedAt = \$retrievedAt'
        $script:intuneModuleText | Should -Match 'AgeMinutes = 0'
    }
}
