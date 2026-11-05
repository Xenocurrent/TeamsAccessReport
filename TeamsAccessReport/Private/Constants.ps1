# Constants: file names (data contract between phases), checkpoint names and finding types.
# Everything that is stored or compared references this file - no magic strings in the logic.

$script:TarFile = @{
    Teams              = 'Teams.csv'
    TeamsAccess        = 'TeamsAccess.csv'
    Channels           = 'Channels.csv'
    IncomingChannels   = 'IncomingSharedChannels.csv'
    SpAssignments      = 'SPAssignments.csv'
    SpGroupMembers     = 'SPGroupMembers.csv'
    EffectiveSpAccess  = 'EffectiveSharePointAccess.csv'
    Drift              = 'DriftReport.csv'
    ExternalAssoc      = 'ExternalAssociations.csv'
    Errors             = 'Errors.csv'
    Selection          = 'Selection.json'
    Summary            = 'Summary.json'
    Log                = 'run.jsonl'
}

$script:TarCheckpoint = @{
    Graph      = 'graph-completed.txt'
    SharePoint = 'sp-completed.txt'
}

$script:TarFinding = @{
    BroadGrant          = 'BroadGrant'
    AnonymousLink       = 'AnonymousLink'
    OrganizationLink    = 'OrganizationLink'
    SiteAdminNotInTeams = 'SiteCollectionAdminNotInTeams'
    SiteNotInTeams      = 'SiteAccessNotInTeams'
    ItemNotInTeams      = 'ItemOrListAccessNotInTeams'
    TeamsNotInSp        = 'InTeamsNotInSharePoint'
    Unresolved          = 'UnresolvedPrincipal'
}

$script:TarRetryableStatus = @(429, 500, 502, 503, 504)
# SharePoint returns role names in the site's language. These are data values, not UI text:
# 'Limited Access' (English) and its Swedish localization. Add the name for every site language you scan.
$script:TarLimitedAccessRoles = @('Limited Access', 'Begränsad åtkomst')

# How phase 2 runs. 'Process' in production (Graph/PnP assembly conflict). Tests set 'InProcess'.
$script:TarSharePointExecutionMode = 'Process'
