Set-StrictMode -Version Latest

function Get-WinTunerHistoryFolder {
    param([Parameter(Mandatory)][guid]$TenantId, [string]$BasePath)
    if ($TenantId -eq [guid]::Empty) { throw 'A verified tenant ID is required for update history.' }
    if ([string]::IsNullOrWhiteSpace($BasePath)) {
        $BasePath = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'WinTuner\UpdateHistory'
    }
    Join-Path $BasePath $TenantId.ToString('D')
}

function Add-WinTunerUpdateHistory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][guid]$TenantId,
        [Parameter(Mandatory)][guid]$OperationId,
        [Parameter(Mandatory)][object]$Result,
        [string]$BasePath
    )

    # Persist a whitelist only: no credentials, raw HTTP errors or tokens.
    $entry = [ordered]@{
        TimestampUtc = [datetime]::UtcNow.ToString('o')
        TenantId = $TenantId.ToString('D')
        OperationId = $OperationId.ToString('D')
        Name = [string]$Result.Name
        PackageId = [string]$Result.PackageId
        GraphId = [string]$Result.GraphId
        CurrentVersion = [string]$Result.CurrentVersion
        TargetVersion = [string]$Result.TargetVersion
        EffectiveVersion = [string]$Result.EffectiveVersion
        Status = if ($Result.Succeeded) { 'Updated' } else { 'Failed' }
        ReasonCode = [string]$Result.ReasonCode
    }
    $warningProperty = $Result.PSObject.Properties['AutoUpdateSummary']
    if ($Result.Succeeded -and $warningProperty -and [string]$warningProperty.Value -like 'Warning:*') { $entry.Status = 'UpdatedWithWarning' }
    $folder = Get-WinTunerHistoryFolder -TenantId $TenantId -BasePath $BasePath
    $null = New-Item -ItemType Directory -Path $folder -Force -ErrorAction Stop
    $destination = Join-Path $folder ([guid]::NewGuid().ToString('N') + '.json')
    $temporary = "$destination.tmp"
    try {
        [IO.File]::WriteAllText($temporary, ($entry | ConvertTo-Json -Compress -Depth 4), [Text.UTF8Encoding]::new($false))
        [IO.File]::Move($temporary, $destination)
    } finally {
        if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -ErrorAction SilentlyContinue }
    }
}

function Get-WinTunerUpdateHistory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][guid]$TenantId,
        [ValidateRange(1,5000)][int]$MaxEntries = 500,
        [string]$BasePath
    )
    $folder = Get-WinTunerHistoryFolder -TenantId $TenantId -BasePath $BasePath
    if (-not (Test-Path -LiteralPath $folder -PathType Container)) { return @() }
    foreach ($file in @(Get-ChildItem -LiteralPath $folder -Filter '*.json' -File -ErrorAction Stop | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First $MaxEntries)) {
        try {
            if ($file.Attributes -band [IO.FileAttributes]::ReparsePoint) { continue }
            $entry = Get-Content -LiteralPath $file.FullName -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            if ([string]$entry.TenantId -ne $TenantId.ToString('D')) { continue }
            $entry.TimestampUtc = ([datetimeoffset]$entry.TimestampUtc).ToUniversalTime().ToString('o')
            # Re-select fields when reading as well, including before CSV export.
            $entry | Select-Object TimestampUtc,TenantId,OperationId,Name,PackageId,GraphId,CurrentVersion,TargetVersion,EffectiveVersion,Status,ReasonCode
        } catch { Write-Verbose 'An unreadable update history entry was skipped.' }
    }
}

function Export-WinTunerUpdateHistory {
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Entries, [Parameter(Mandatory)][string]$Path)
    $safeRows = foreach ($entry in $Entries) {
        $row = [ordered]@{}
        foreach ($field in 'TimestampUtc','TenantId','OperationId','Name','PackageId','GraphId','CurrentVersion','TargetVersion','EffectiveVersion','Status','ReasonCode') {
            $value = [string]$entry.$field
            if ($value -match '^\s*[=+@-]' -or $value -match '^[\t\r\n]') { $value = "'$value" }
            $row[$field] = $value
        }
        [pscustomobject]$row
    }
    @($safeRows) | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding utf8BOM -ErrorAction Stop
}

Export-ModuleMember -Function Add-WinTunerUpdateHistory,Get-WinTunerUpdateHistory,Export-WinTunerUpdateHistory
