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

# CSV column contract per file. Every record constructor must emit exactly these properties in this order;
# the integration tests assert it. Used to write header-only files (phase ran, 0 rows) and to detect
# output folders created by an incompatible version before -Resume appends to them.
$script:TarColumns = @{
    $script:TarFile.Teams             = @('TeamId', 'TeamName', 'Visibility', 'IsArchived', 'WebUrl', 'CreatedDateTime', 'Channels', 'Completeness', 'FailedStages')
    $script:TarFile.TeamsAccess       = @('TeamId', 'TeamName', 'Scope', 'ChannelId', 'ChannelName', 'ChannelType', 'UserId', 'DisplayName', 'Email', 'UserTenantId', 'Role', 'RawRoles', 'AccessPath', 'SourceTeamId', 'MemberType')
    $script:TarFile.Channels          = @('TeamId', 'TeamName', 'ChannelId', 'ChannelName', 'ChannelType', 'SiteUrl', 'FilesFolderUrl', 'DriveId', 'FolderItemId', 'SiteSource', 'OwnSharePointSite')
    $script:TarFile.IncomingChannels  = @('TeamId', 'TeamName', 'ChannelId', 'ChannelName', 'ChannelType', 'HostTenantId', 'CrossTenant')
    $script:TarFile.SpAssignments     = @('SiteUrl', 'WebUrl', 'Scope', 'ObjectPath', 'ListTitle', 'ItemId', 'PrincipalType', 'PrincipalTitle', 'PrincipalLogin', 'PrincipalEmail', 'Roles', 'LimitedAccessOnly', 'AssociatedGroup')
    $script:TarFile.SpGroupMembers    = @('SiteUrl', 'GroupTitle', 'MemberTitle', 'MemberLogin', 'MemberEmail', 'MemberPrincipalType')
    $script:TarFile.EffectiveSpAccess = @('SiteUrl', 'WebUrl', 'Scope', 'ObjectPath', 'ListTitle', 'ItemId', 'Roles', 'LimitedAccessOnly', 'AssociatedGroup', 'Kind', 'UserId', 'Upn', 'DisplayName', 'UserType', 'ExpansionPath')
    $script:TarFile.Drift             = @('SiteUrl', 'Site', 'SiteKind', 'Finding', 'Severity', 'Cause', 'UserId', 'Upn', 'DisplayName', 'Scope', 'ObjectPath', 'Roles', 'ExpansionPath')
    $script:TarFile.ExternalAssoc     = @('UserId', 'HostTenantId', 'HostTeamId', 'HostTeamName')
    $script:TarFile.Errors            = @('Timestamp', 'Stage', 'TeamId', 'ChannelId', 'Target', 'Error')
}

# Selection.json schema written by this version. Import accepts 1 (ChannelIds = 'ALL' | id | [ids]) and 2.
$script:TarSelectionSchemaVersion = 2

# Application permissions of the report app - single source for the scope check and for New-TeamsAccessReportApp.
# Resource app IDs are Microsoft's well-known first-party IDs; permission IDs are looked up in the tenant at runtime.
$script:TarAppPermissions = [ordered]@{
    '00000003-0000-0000-c000-000000000000' = @{ Name = 'Microsoft Graph'; Roles = @('Team.ReadBasic.All', 'Channel.ReadBasic.All', 'TeamMember.Read.All', 'ChannelMember.Read.All', 'Files.Read.All', 'GroupMember.Read.All', 'User.Read.All', 'Organization.Read.All') }
    # Reading role assignments requires Full Control (README, Permissions).
    '00000003-0000-0ff1-ce00-000000000000' = @{ Name = 'Office 365 SharePoint Online'; Roles = @('Sites.FullControl.All') }
}
# Delegated scopes the admin consents to while New-TeamsAccessReportApp runs (create app, grant app roles).
# User.Read: GET /organization (tenant name and initial domain when the tenant comes from the sign-in).
$script:TarProvisioningScopes = @('Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All', 'User.Read')

# Public Entra authority (global cloud) used for unauthenticated tenant discovery.
$script:TarAuthority = 'https://login.microsoftonline.com'
# profiles.json schema written by this version.
$script:TarProfileSchemaVersion = 1

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

# Known by-design explanations for a drift finding. Empty Cause = no known default explains it.
$script:TarCause = @{
    # Microsoft adds 'Everyone except external users' to the associated Members group of public team sites.
    PublicTeamDefault = 'PublicTeamDefault'
}

$script:TarRetryableStatus = @(429, 500, 502, 503, 504)
# SharePoint returns role names in the site's language. These are data values, not UI text:
# 'Limited Access' / 'Web-Only Limited Access' (English) and the Swedish 'Limited Access'. Add the names for every
# site language you scan (the Swedish name of 'Web-Only Limited Access' is not verified, so it is not listed).
$script:TarLimitedAccessRoles = @('Limited Access', 'Web-Only Limited Access', 'Begränsad åtkomst')

# How phase 2 runs. 'Process' in production (Graph/PnP assembly conflict). Tests set 'InProcess'.
$script:TarSharePointExecutionMode = 'Process'
