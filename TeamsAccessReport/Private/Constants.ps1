# Konstanter: filnamn (datakontrakt mellan faserna), checkpoint-namn och fyndtyper.
# Allt som lagras eller jämförs refererar hit – inga magiska strängar i logiken.

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
$script:TarLimitedAccessRoles = @('Limited Access', 'Begränsad åtkomst')

# Hur fas 2 körs. 'Process' i produktion (Graph/PnP-assemblykonflikt). Tester sätter 'InProcess'.
$script:TarSharePointExecutionMode = 'Process'
