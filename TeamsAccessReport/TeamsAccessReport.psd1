@{
    RootModule           = 'TeamsAccessReport.psm1'
    ModuleVersion        = '2.0.0'
    GUID                 = '6f0c8a51-3d7e-4b9a-9a1e-2c4f7d5b8e10'
    Author               = '<AUTHOR>'
    CompanyName          = '<COMPANY>'
    Description          = 'Åtkomstrapport för Microsoft Teams + SharePoint (Graph + PnP) med checkpoint, avvikelseanalys och externa kopplingar.'
    PowerShellVersion    = '7.2'
    CompatiblePSEditions = @('Core')
    # Microsoft.Graph.Authentication importeras vid körning (inte RequiredModules) så att worker-processen
    # aldrig laddar Graph-assemblies tillsammans med PnP.
    FunctionsToExport    = @('Invoke-TeamsAccessReport')
    CmdletsToExport      = @()
    VariablesToExport    = @()
    AliasesToExport      = @()
    PrivateData          = @{ PSData = @{ Tags = @('Teams', 'SharePoint', 'MicrosoftGraph', 'PnP', 'Audit', 'Permissions'); ReleaseNotes = 'Se CHANGELOG.md' } }
}
