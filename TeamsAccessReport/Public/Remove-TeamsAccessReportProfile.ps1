function Remove-TeamsAccessReportProfile {
    <#
    .SYNOPSIS
        Deletes a saved connection profile. The certificate and the Entra app are not touched.
    .PARAMETER Name
        Exact profile name.
    .EXAMPLE
        Remove-TeamsAccessReportProfile -Name contoso
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param([Parameter(Mandatory, Position = 0)][ValidateNotNullOrEmpty()][string]$Name)
    $null = Get-TarProfile -Name $Name
    if ($PSCmdlet.ShouldProcess($Name, 'Remove connection profile')) {
        Write-TarProfileStore -Profiles @(Read-TarProfileStore | Where-Object { $_.Name -ne $Name }) -Confirm:$false
        Write-TarLog "Removed connection profile '$Name'."
    }
}
