@{
    RootModule           = 'TeamsAccessReport.psm1'
    ModuleVersion        = '2.8.1'
    GUID                 = '6f0c8a51-3d7e-4b9a-9a1e-2c4f7d5b8e10'
    Author               = '<AUTHOR>'
    CompanyName          = '<COMPANY>'
    Description          = 'Access report for Microsoft Teams + SharePoint (Graph + PnP) with checkpoints, drift analysis and external associations.'
    PowerShellVersion    = '7.2'
    CompatiblePSEditions = @('Core')
    # Microsoft.Graph.Authentication is imported at run time (not RequiredModules) so that the worker process
    # never loads Graph assemblies together with PnP.
    FunctionsToExport    = @('Invoke-TeamsAccessReport', 'New-TeamsAccessReportProfile', 'Get-TeamsAccessReportProfile', 'Remove-TeamsAccessReportProfile', 'New-TeamsAccessReportPnPApp', 'New-TeamsAccessReportApp')
    CmdletsToExport      = @()
    VariablesToExport    = @()
    AliasesToExport      = @()
    PrivateData          = @{ PSData = @{ Tags = @('Teams', 'SharePoint', 'MicrosoftGraph', 'PnP', 'Audit', 'Permissions'); ReleaseNotes = 'See CHANGELOG.md' } }
}
