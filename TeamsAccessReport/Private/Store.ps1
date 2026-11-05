# Report store: CSV files are the data contract between phases; checkpoint files enable resume.

function New-TarContext {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Pure in-memory constructor; does not change system state.')]
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$OutputPath,
        [int]$MaxRetries = 6, [int]$BaseDelaySeconds = 3, [int]$MaxItemsPerList = 0, [int]$MaxSubwebDepth = 5
    )
    $full = [System.IO.Path]::GetFullPath($OutputPath)
    [pscustomobject]@{
        OutputPath = $full; StateDir = (Join-Path $full 'state'); GraphBase = 'https://graph.microsoft.com/v1.0'
        TenantId = $null; TenantName = $null; TenantDomain = $null; InitialDomain = $null; AuthType = $null
        AppOnly = $false; ClientId = $null; CertificateThumbprint = $null; ProfileName = $null
        MaxRetries = $MaxRetries; BaseDelaySeconds = $BaseDelaySeconds; MaxItemsPerList = $MaxItemsPerList; MaxSubwebDepth = $MaxSubwebDepth
    }
}

function Initialize-TarStore {
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)]$Context, [switch]$Resume)
    if ($Resume) {
        if (-not (Test-Path (Join-Path $Context.OutputPath $script:TarFile.Selection))) { throw "-Resume: $($script:TarFile.Selection) is missing in $($Context.OutputPath)." }
        Assert-TarStoreSchema -Context $Context
    } elseif (Test-Path $Context.StateDir) {
        throw "$($Context.OutputPath) already contains a run. Use -Resume or a new -OutputPath."
    }
    if ($PSCmdlet.ShouldProcess($Context.StateDir, 'Create report directory')) {
        $null = New-Item -Path $Context.StateDir -ItemType Directory -Force
    }
}

function Get-TarCsvHeader {
    # Column names from the first line of an existing CSV file (our files never contain commas in column names).
    [CmdletBinding()]
    [OutputType([string[]])]
    param([Parameter(Mandatory)][string]$Path)
    $line = Get-Content -Path $Path -TotalCount 1 -Encoding utf8
    if (-not $line) { return [string[]]@() }
    [string[]]@(($line -replace '^\uFEFF', '') -split ',' | ForEach-Object { $_.Trim().Trim('"') })
}

function Assert-TarStoreSchema {
    # -Resume appends to existing files; refuse when they were written with a different column contract.
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context)
    foreach ($name in $script:TarColumns.Keys) {
        $p = Get-TarStorePath $Context $name
        if (-not (Test-Path $p)) { continue }
        $actual = Get-TarCsvHeader -Path $p
        $expected = $script:TarColumns[$name]
        if (($actual -join ',') -ne ($expected -join ',')) {
            throw "-Resume: $name in $($Context.OutputPath) was written by an incompatible version (columns differ). Start a new run without -Resume."
        }
    }
}

function Initialize-TarRecordFile {
    # Writes a header-only CSV for each named file that does not exist yet.
    # Contract: file present = the phase ran (header only = 0 rows); file absent = the phase did not run.
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][string[]]$Name)
    foreach ($n in $Name) {
        if (-not $script:TarColumns.ContainsKey($n)) { throw "No column contract for $n (add it to `$script:TarColumns)." }
        $p = Get-TarStorePath $Context $n
        if ((Test-Path $p) -or -not $PSCmdlet.ShouldProcess($p, 'Write CSV header')) { continue }
        Set-Content -Path $p -Value ((@($script:TarColumns[$n] | ForEach-Object { '"' + $_ + '"' })) -join ',') -Encoding utf8BOM
    }
}

function Get-TarStorePath {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][string]$Name)
    Join-Path $Context.OutputPath $Name
}

function Write-TarRecord {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][string]$Name, [AllowEmptyCollection()][object[]]$InputObject)
    if (-not $InputObject -or $InputObject.Count -eq 0) { return }
    $InputObject | Protect-TarCsvValue | Export-Csv -Path (Get-TarStorePath $Context $Name) -Append -NoTypeInformation -Encoding utf8BOM
}

function Read-TarRecord {
    # Emits rows to the pipeline - call as @(Read-TarRecord ...).
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][string]$Name)
    $p = Get-TarStorePath $Context $Name
    if (Test-Path $p) { Import-Csv -Path $p -Encoding utf8 }
}

function Set-TarRecord {
    # Rewrites a whole file (used for recalculation/cleanup). An empty set leaves a header-only file.
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][string]$Name, [AllowEmptyCollection()][object[]]$InputObject)
    $p = Get-TarStorePath $Context $Name
    if (-not $PSCmdlet.ShouldProcess($p, 'Rewrite')) { return }
    if (Test-Path $p) { Remove-Item $p -Force }
    if ($InputObject -and $InputObject.Count) { $InputObject | Protect-TarCsvValue | Export-Csv -Path $p -NoTypeInformation -Encoding utf8BOM }
    else { Initialize-TarRecordFile -Context $Context -Name $Name }
}

function Add-TarErrorRecord {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][string]$Stage, [string]$TeamId, [string]$ChannelId, [string]$Target, [string]$Message)
    Write-TarRecord -Context $Context -Name $script:TarFile.Errors -InputObject @([pscustomobject]@{
            Timestamp = (Get-Date).ToString('o'); Stage = $Stage; TeamId = $TeamId; ChannelId = $ChannelId; Target = $Target; Error = $Message
        })
}

function Get-TarCheckpoint {
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.HashSet[string]])]
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][string]$Name)
    $set = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $p = Join-Path $Context.StateDir $Name
    if (Test-Path $p) { foreach ($l in [System.IO.File]::ReadAllLines($p)) { if ($l.Trim()) { [void]$set.Add($l.Trim()) } } }
    Write-Output -NoEnumerate -InputObject $set
}

function Add-TarCheckpoint {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][string]$Key)
    [System.IO.File]::AppendAllLines((Join-Path $Context.StateDir $Name), [string[]]@($Key), [System.Text.UTF8Encoding]::new($false))
}

function Remove-TarCheckpoint {
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][string[]]$Key)
    $p = Join-Path $Context.StateDir $Name
    if (-not (Test-Path $p) -or -not $PSCmdlet.ShouldProcess($p, 'Remove checkpoint keys')) { return }
    $remaining = [string[]]@([System.IO.File]::ReadAllLines($p) | Where-Object { $_.Trim() -and $_.Trim() -notin $Key })
    [System.IO.File]::WriteAllLines($p, $remaining, [System.Text.UTF8Encoding]::new($false))
}
