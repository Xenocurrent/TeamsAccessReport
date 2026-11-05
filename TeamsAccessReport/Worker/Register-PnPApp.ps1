#Requires -Version 7.2
<#
.SYNOPSIS
    Registers the Entra app PnP PowerShell needs for delegated (interactive) sign-in. Runs in its own pwsh process
    because PnP.PowerShell and the Microsoft Graph SDK cannot share a session (assembly conflict).
.DESCRIPTION
    Least privilege: one delegated permission, SharePoint AllSites.FullControl (needed to read role assignments).
    Register-PnPEntraIDAppForInteractiveLogin opens a browser sign-in, creates the app with redirect URI
    http://localhost and then starts the admin-consent flow. The result is written as JSON to -ResultPath:
    { ok: true, clientId: '<guid>' } or { ok: false, error: '<message>' }.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Tenant,
    [Parameter(Mandatory)][string]$ApplicationName,
    [Parameter(Mandatory)][string]$ResultPath
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$result = [ordered]@{ ok = $false; clientId = $null; error = $null }
try {
    Write-Information "Registering '$ApplicationName' in $Tenant (a browser window opens for sign-in, then for consent) ..." -InformationAction Continue
    Import-Module PnP.PowerShell -ErrorAction Stop
    $r = Register-PnPEntraIDAppForInteractiveLogin -ApplicationName $ApplicationName -Tenant $Tenant -SharePointDelegatePermissions 'AllSites.FullControl'
    # The cmdlet returns a PSObject with the property 'AzureAppId/ClientId'; fall back to any GUID-valued id property.
    $guid = [guid]::Empty
    $candidates = @($r | ForEach-Object { $_.PSObject.Properties } | Where-Object { $_.Name -match 'ClientId|AppId' } | ForEach-Object { [string]$_.Value })
    $id = @($candidates | Where-Object { [guid]::TryParse($_, [ref]$guid) }) | Select-Object -First 1
    if (-not $id) { throw "Register-PnPEntraIDAppForInteractiveLogin returned no client ID (properties: $(@($r | ForEach-Object { $_.PSObject.Properties.Name }) -join ', '))." }
    $result.ok = $true; $result.clientId = $id
    Write-Information "Done. Client ID: $id - this window can be closed." -InformationAction Continue
} catch {
    $result.error = $_.Exception.Message
    Write-Warning $result.error
} finally {
    $result | ConvertTo-Json | Set-Content -Path $ResultPath -Encoding utf8
}
