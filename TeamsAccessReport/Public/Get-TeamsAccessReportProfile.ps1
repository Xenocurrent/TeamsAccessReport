function Get-TeamsAccessReportProfile {
    <#
    .SYNOPSIS
        Lists saved connection profiles.
    .PARAMETER Name
        Profile name; wildcards are allowed. Default: all profiles.
    .EXAMPLE
        Get-TeamsAccessReportProfile
    .OUTPUTS
        PSCustomObject (one per profile)
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([Parameter(Position = 0)][ValidateNotNullOrEmpty()][string]$Name = '*')
    @(Read-TarProfileStore) | Where-Object { $_.Name -like $Name }
}
