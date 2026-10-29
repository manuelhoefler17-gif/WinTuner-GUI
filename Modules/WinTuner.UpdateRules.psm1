Set-StrictMode -Version Latest
Import-Module (Join-Path $PSScriptRoot 'WinTuner.Core.psm1') -ErrorAction Stop

function New-WinTunerUpdateRule {
    param(
        [Parameter(Mandatory)][ValidatePattern('^[A-Za-z0-9][A-Za-z0-9._-]{0,255}$')][string]$PackageId,
        [bool]$ExcludeFromAll = $false,
        [AllowEmptyString()][string]$PinnedVersion = '',
        [AllowEmptyString()][string]$Reason = '',
        [AllowEmptyString()][string]$ExpiresUtc = ''
    )
    $PinnedVersion = $PinnedVersion.Trim()
    if ($PinnedVersion -and $PinnedVersion -notmatch '^[A-Za-z0-9][A-Za-z0-9._+\-]{0,127}$') { throw 'Enter an exact package version without spaces or wildcards.' }
    if ($Reason.Length -gt 500) { throw 'The rule reason must not exceed 500 characters.' }
    $expiry = ''
    if ($ExpiresUtc) { $expiry = ([datetimeoffset]::Parse($ExpiresUtc, [Globalization.CultureInfo]::InvariantCulture)).ToUniversalTime().ToString('o') }
    [pscustomobject]@{ PackageId = $PackageId; ExcludeFromAll = $ExcludeFromAll; PinnedVersion = $PinnedVersion; Reason = $Reason.Trim(); ExpiresUtc = $expiry }
}

function Get-WinTunerUpdateRulesPath {
    param([Parameter(Mandatory)][guid]$TenantId, [string]$BasePath)
    if ($TenantId -eq [guid]::Empty) { throw 'Update rules require a verified tenant ID.' }
    if ([string]::IsNullOrWhiteSpace($BasePath)) { $BasePath = Join-Path ([Environment]::GetFolderPath('ApplicationData')) 'WinTunerGUI\UpdateRules' }
    Join-Path $BasePath ($TenantId.ToString('D') + '.json')
}

function Get-WinTunerUpdateRules {
    [CmdletBinding()]
    param([Parameter(Mandatory)][guid]$TenantId, [string]$BasePath)
    $path = Get-WinTunerUpdateRulesPath -TenantId $TenantId -BasePath $BasePath
    if (-not (Test-Path -LiteralPath $path)) { return @() }
    try {
        $document = Get-Content -LiteralPath $path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        if ($document.SchemaVersion -ne 1 -or [string]$document.TenantId -ne $TenantId.ToString('D')) { throw 'Invalid rule schema or tenant.' }
        $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        $rules = [Collections.Generic.List[object]]::new()
        foreach ($rule in @($document.Rules)) {
            if ($rule.ExcludeFromAll -isnot [bool] -or -not $seen.Add([string]$rule.PackageId)) { throw 'Invalid or duplicate rule.' }
            $expiry = if ($rule.ExpiresUtc) { ([datetimeoffset]$rule.ExpiresUtc).ToUniversalTime().ToString('o') } else { '' }
            $rules.Add((New-WinTunerUpdateRule -PackageId $rule.PackageId -ExcludeFromAll $rule.ExcludeFromAll -PinnedVersion $rule.PinnedVersion -Reason $rule.Reason -ExpiresUtc $expiry))
        }
        return $rules.ToArray()
    } catch { throw 'The local update rules could not be read or validated. Updates are blocked until the rule file is repaired.' }
}

function Set-WinTunerUpdateRule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][guid]$TenantId,
        [Parameter(Mandatory)][string]$PackageId,
        [bool]$ExcludeFromAll = $false,
        [string]$PinnedVersion = '', [string]$Reason = '', [string]$ExpiresUtc = '',
        [switch]$Remove,
        [string]$BasePath
    )
    $rule = New-WinTunerUpdateRule -PackageId $PackageId -ExcludeFromAll $ExcludeFromAll -PinnedVersion $PinnedVersion -Reason $Reason -ExpiresUtc $ExpiresUtc
    $path = Get-WinTunerUpdateRulesPath -TenantId $TenantId -BasePath $BasePath
    $lockKey = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes([IO.Path]::GetFullPath($path).ToUpperInvariant())))
    $mutex = [Threading.Mutex]::new($false, "Local\WinTuner-UpdateRules-$lockKey")
    $acquired = $false
    $temporary = "$path.$([guid]::NewGuid().ToString('N')).tmp"
    try {
        try { $acquired = $mutex.WaitOne(5000) } catch [Threading.AbandonedMutexException] { $acquired = $true }
        if (-not $acquired) { throw 'Another instance is saving update rules. Try again.' }
        $rules = @(Get-WinTunerUpdateRules -TenantId $TenantId -BasePath $BasePath | Where-Object { $_.PackageId -ine $PackageId })
        if (-not $Remove -and ($ExcludeFromAll -or $rule.PinnedVersion)) { $rules += $rule }
        $document = @{ SchemaVersion = 1; TenantId = $TenantId.ToString('D'); Rules = @($rules) }
        $null = New-Item -ItemType Directory -Path (Split-Path -Parent $path) -Force -ErrorAction Stop
        [IO.File]::WriteAllText($temporary, ($document | ConvertTo-Json -Depth 6), [Text.UTF8Encoding]::new($false))
        [IO.File]::Move($temporary, $path, $true)
    } finally {
        if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -ErrorAction SilentlyContinue }
        if ($acquired) { $mutex.ReleaseMutex() }
        $mutex.Dispose()
    }
}

function Get-WinTunerUpdateRuleDecision {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$App,
        [AllowEmptyCollection()][object[]]$Rules = @(),
        [datetimeoffset]$Now = [datetimeoffset]::UtcNow
    )
    $sourceProperty = $App.PSObject.Properties['ScannedLatestVersion']
    $target = if ($sourceProperty) { [string]$sourceProperty.Value } else { [string]$App.LatestVersion }
    $rule = @($Rules | Where-Object { $_.PackageId -ieq [string]$App.PackageId })
    if ($rule.Count -gt 1) { throw 'Multiple update rules match the same package.' }
    $description = ''
    $excluded = $false
    $pin = ''
    if ($rule.Count -eq 1) {
        if ($rule[0].ExpiresUtc -and ([datetimeoffset]$rule[0].ExpiresUtc) -le $Now) {
            $description = 'Rule expired (ignored)'
        } else {
            $excluded = [bool]$rule[0].ExcludeFromAll
            $pin = [string]$rule[0].PinnedVersion
            $parts = @()
            if ($excluded) { $parts += 'Excluded from Update All' }
            if ($pin) { $target = $pin; $parts += "Pinned to $pin" }
            if ($rule[0].Reason) { $parts += [string]$rule[0].Reason }
            if ($rule[0].ExpiresUtc) { $parts += "Until $($rule[0].ExpiresUtc)" }
            $description = $parts -join '; '
        }
    }
    $canUpdate = -not [string]::IsNullOrWhiteSpace([string]$App.PackageId) -and (Test-IsNewerVersion -Latest $target -Current ([string]$App.CurrentVersion))
    if ($pin -and -not $canUpdate) { $description += '; On hold (no upgrade to the pinned version)' }
    [pscustomobject]@{
        CanUpdateAll = $canUpdate -and -not $excluded
        CanUpdateSelected = $canUpdate
        TargetVersion = $target; PinnedVersion = $pin
        ExcludedFromAll = $excluded; Description = $description
    }
}

function Resolve-WinTunerUpdateCandidates {
    [CmdletBinding()]
    param([Parameter(Mandatory)][object[]]$Apps, [object[]]$Rules = @(), [ValidateSet('Checked','All')][string]$Mode = 'Checked')
    $allowed = [Collections.Generic.List[object]]::new()
    $skipped = [Collections.Generic.List[object]]::new()
    foreach ($app in $Apps) {
        $decision = Get-WinTunerUpdateRuleDecision -App $app -Rules $Rules
        $eligible = if ($Mode -eq 'All') { $decision.CanUpdateAll } else { $decision.CanUpdateSelected }
        if (-not $eligible) {
            $skipped.Add([pscustomobject]@{ Name = [string]$app.Name; Description = $decision.Description })
            continue
        }
        $copy = $app | Select-Object *
        $sourceProperty = $app.PSObject.Properties['ScannedLatestVersion']
        $source = if ($sourceProperty) { [string]$sourceProperty.Value } else { [string]$app.LatestVersion }
        $copy | Add-Member -NotePropertyName ScannedLatestVersion -NotePropertyValue $source -Force
        $copy | Add-Member -NotePropertyName RuleDescription -NotePropertyValue $decision.Description -Force
        $copy | Add-Member -NotePropertyName PinnedVersion -NotePropertyValue $decision.PinnedVersion -Force
        $copy.LatestVersion = $decision.TargetVersion
        $allowed.Add($copy)
    }
    [pscustomobject]@{ Apps = $allowed.ToArray(); Skipped = $skipped.ToArray() }
}

Export-ModuleMember -Function New-WinTunerUpdateRule,Get-WinTunerUpdateRules,Set-WinTunerUpdateRule,Get-WinTunerUpdateRuleDecision,Resolve-WinTunerUpdateCandidates
