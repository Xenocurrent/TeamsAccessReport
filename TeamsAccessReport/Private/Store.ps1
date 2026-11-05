# Rapportlager: CSV-filer är datakontraktet mellan faserna; checkpoint-filer ger återupptag.

function New-TarContext {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Ren konstruktor i minnet; ändrar inget systemtillstånd.')]
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
        AppOnly = $false; ClientId = $null; CertificateThumbprint = $null
        MaxRetries = $MaxRetries; BaseDelaySeconds = $BaseDelaySeconds; MaxItemsPerList = $MaxItemsPerList; MaxSubwebDepth = $MaxSubwebDepth
    }
}

function Initialize-TarStore {
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)]$Context, [switch]$Resume)
    if ($Resume) {
        if (-not (Test-Path (Join-Path $Context.OutputPath $script:TarFile.Selection))) { throw "-Resume: $($script:TarFile.Selection) saknas i $($Context.OutputPath)." }
    } elseif (Test-Path $Context.StateDir) {
        throw "$($Context.OutputPath) innehåller redan en körning. Använd -Resume eller en ny -OutputPath."
    }
    if ($PSCmdlet.ShouldProcess($Context.StateDir, 'Skapa rapportkatalog')) {
        $null = New-Item -Path $Context.StateDir -ItemType Directory -Force
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
    # Skickar rader i pipeline – anropa som @(Read-TarRecord ...).
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][string]$Name)
    $p = Get-TarStorePath $Context $Name
    if (Test-Path $p) { Import-Csv -Path $p -Encoding utf8 }
}

function Set-TarRecord {
    # Skriver om en hel fil (används vid omräkning/rensning). Tom mängd tar bort filen.
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][string]$Name, [AllowEmptyCollection()][object[]]$InputObject)
    $p = Get-TarStorePath $Context $Name
    if (-not $PSCmdlet.ShouldProcess($p, 'Skriv om')) { return }
    if (Test-Path $p) { Remove-Item $p -Force }
    if ($InputObject -and $InputObject.Count) { $InputObject | Protect-TarCsvValue | Export-Csv -Path $p -NoTypeInformation -Encoding utf8BOM }
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
    if (-not (Test-Path $p) -or -not $PSCmdlet.ShouldProcess($p, 'Ta bort checkpoint-nycklar')) { return }
    $remaining = [string[]]@([System.IO.File]::ReadAllLines($p) | Where-Object { $_.Trim() -and $_.Trim() -notin $Key })
    [System.IO.File]::WriteAllLines($p, $remaining, [System.Text.UTF8Encoding]::new($false))
}
