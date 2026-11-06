# Test data: deterministic Graph and SharePoint responses. The first matching route wins (ordered list).
# Fault injection is controlled via $global:TarTest (read by the mock bodies).

$script:Tenant = '11111111-1111-1111-1111-111111111111'
$script:ExtTenant = '33333333-3333-3333-3333-333333333333'
$script:G1 = 'aaaaaaaa-0000-0000-0000-000000000001'   # Finance
$script:G2 = 'aaaaaaaa-0000-0000-0000-000000000002'   # IT
$script:Sp = 'https://contoso.sharepoint.com/sites'

function Reset-TarTestState {
    $global:TarTest = @{ FailAllMembers = $false; FailSite = $null; ThrottleOnce = $true; Calls = [System.Collections.Generic.List[string]]::new() }
}

function New-GraphOk { param($Body) [pscustomobject]@{ StatusCode = 200; Body = $Body; RetryAfterSeconds = 0 } }
function New-GraphErr { param([int]$Code, [string]$Msg = 'error') [pscustomobject]@{ StatusCode = $Code; Body = @{ error = @{ code = 'x'; message = $Msg } }; RetryAfterSeconds = 0 } }
function New-User { param($Id, $Name, $Upn, $Type = 'Member') @{ '@odata.type' = '#microsoft.graph.user'; id = $Id; displayName = $Name; userPrincipalName = $Upn; userType = $Type } }
function New-Member { param($Id, $Name, $Mail, [string[]]$Roles = @(), $TenantId = $script:Tenant, $Source)
    $m = @{ '@odata.type' = '#microsoft.graph.aadUserConversationMember'; userId = $Id; displayName = $Name; email = $Mail; tenantId = $TenantId; roles = $Roles }
    if ($Source) { $m['@microsoft.graph.originalSourceMembershipUrl'] = $Source }
    $m
}

$script:SiteByChannel = @{ C1 = "$script:Sp/Finance"; C2 = "$script:Sp/Finance-Payroll"; C3 = "$script:Sp/Finance-Partner"; C4 = "$script:Sp/IT" }

$script:GraphRoutes = @(
    @{ P = '/organization'; R = { New-GraphOk @{ value = @(@{ id = $script:Tenant; displayName = 'Contoso AB'; verifiedDomains = @(@{ name = 'contoso.se'; isDefault = $true; isInitial = $false }, @{ name = 'contoso.onmicrosoft.com'; isDefault = $false; isInitial = $true }) }) } } }
    @{ P = '/teams\?\$select=id,displayName,description,visibility$'; R = { New-GraphOk @{ value = @(@{ id = $script:G1; displayName = 'Finance'; description = 'Finance department'; visibility = 'private' }); '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/teams?$skiptoken=p2' } } }
    @{ P = '/teams\?\$skiptoken=p2'; R = { New-GraphOk @{ value = @(@{ id = $script:G2; displayName = 'IT'; description = $null; visibility = 'public' }) } } }
    @{ P = '/teams/([0-9a-f-]{36})\?'; R = { param($m) $n = if ($m.Groups[1].Value -eq $script:G1) { 'Finance' } else { 'IT' }; $v = if ($m.Groups[1].Value -eq $script:G2) { 'public' } else { 'private' }; New-GraphOk @{ id = $m.Groups[1].Value; displayName = $n; visibility = $v; isArchived = ($m.Groups[1].Value -eq $script:G2); webUrl = "https://teams/$n"; createdDateTime = '2025-01-01T00:00:00Z' } } }
    @{ P = "/teams/$script:G1/members$"; R = { New-GraphOk @{ value = @((New-Member 'u1' 'Anna' 'anna@contoso.se' @('owner')), (New-Member 'u3' 'Bo' 'bo@contoso.se')) } } }
    @{ P = "/teams/$script:G2/members$"; R = { New-GraphOk @{ value = @((New-Member 'u1' 'Anna' 'anna@contoso.se' @('owner') $script:Tenant "https://graph.microsoft.com/beta/tenants/$script:Tenant/teams/$script:G1/channels/C2/members/m1")) } } }
    @{ P = "/teams/$script:G1/incomingChannels"; R = { New-GraphOk @{ value = @(@{ id = 'X1'; displayName = 'Fabrikam Partner'; membershipType = 'shared'; tenantId = $script:ExtTenant }) } } }
    @{ P = '/incomingChannels'; R = { New-GraphOk @{ value = @() } } }
    @{ P = "/teams/$script:G1/channels\?"; R = { New-GraphOk @{ value = @(@{ id = 'C1'; displayName = 'General'; membershipType = 'standard' }, @{ id = 'C2'; displayName = 'Payroll'; membershipType = 'private' }, @{ id = 'C3'; displayName = 'Partner'; membershipType = 'shared' }) } } }
    @{ P = "/teams/$script:G2/channels\?"; R = { New-GraphOk @{ value = @(@{ id = 'C4'; displayName = 'General'; membershipType = 'standard' }) } } }
    @{ P = '/channels/C4/filesFolder'; R = { New-GraphErr 404 'Folder location for this channel is not ready yet, please try again later.' } }
    @{ P = "/groups/$($script:G2)/drive/root"; R = { New-GraphOk @{ id = 'FC4'; webUrl = "$($script:SiteByChannel['C4'])/Shared Documents"; parentReference = @{ driveId = 'DC4' }; sharepointIds = @{ siteUrl = $script:SiteByChannel['C4'] } } } }
    @{ P = '/channels/(C\d)/filesFolder'; R = { param($m) New-GraphOk @{ id = "F$($m.Groups[1].Value)"; webUrl = "$($script:SiteByChannel[$m.Groups[1].Value])/Shared Documents/x"; parentReference = @{ driveId = "D$($m.Groups[1].Value)" } } } }
    @{ P = '/drives/D(C\d)/items/'; R = { param($m) New-GraphOk @{ id = "F$($m.Groups[1].Value)"; sharepointIds = @{ siteUrl = $script:SiteByChannel[$m.Groups[1].Value] } } } }
    @{ P = '/channels/C2/allMembers'; R = { if ($global:TarTest.FailAllMembers) { return New-GraphErr 403 'Forbidden' }; New-GraphOk @{ value = @((New-Member 'u1' 'Anna' 'anna@contoso.se' @('owner'))) } } }
    @{ P = '/channels/C3/allMembers'; R = {
            if ($global:TarTest.ThrottleOnce) { $global:TarTest.ThrottleOnce = $false; return New-GraphErr 503 'Service unavailable' }
            New-GraphOk @{ value = @((New-Member 'u3' 'Bo' 'bo@contoso.se'), (New-Member 'ux' '=HYPERLINK("http://evil")' 'x@fabrikam.com' @() $script:ExtTenant "https://graph.microsoft.com/beta/tenants/$script:ExtTenant/teams/HT/channels/c/members/m")) } } }
    @{ P = '/users/anna'; R = { New-GraphOk (New-User 'u1' 'Anna' 'anna@contoso.se') } }
    @{ P = '/users/bo'; R = { New-GraphOk (New-User 'u3' 'Bo' 'bo@contoso.se') } }
    @{ P = '/users/cecilia'; R = { New-GraphOk (New-User 'u9' 'Cecilia Guest' 'cecilia_ext.com#EXT#@contoso.onmicrosoft.com' 'Guest') } }
    @{ P = '/users/ghost'; R = { New-GraphErr 404 'Not found' } }
    @{ P = "/groups/$script:G1/transitiveMembers"; R = { New-GraphOk @{ value = @((New-User 'u1' 'Anna' 'anna@contoso.se'), (New-User 'u3' 'Bo' 'bo@contoso.se'), @{ '@odata.type' = '#microsoft.graph.group'; id = 'nested' }) } } }
    @{ P = "/groups/$script:G1/owners"; R = { New-GraphOk @{ value = @((New-User 'u1' 'Anna' 'anna@contoso.se')) } } }
    @{ P = "/groups/$script:G2/(transitiveMembers|owners)"; R = { New-GraphOk @{ value = @((New-User 'u1' 'Anna' 'anna@contoso.se')) } } }
    @{ P = '/users/u\d/teamwork/associatedTeams'; R = { New-GraphOk @{ value = @(@{ id = 'HT'; displayName = 'Fabrikam Team'; tenantId = $script:ExtTenant }, @{ id = $script:G1; displayName = 'Finance'; tenantId = $script:Tenant }) } } }
)

function Get-TestGraphResponse {
    param([string]$Uri)
    $global:TarTest.Calls.Add($Uri)
    foreach ($r in $script:GraphRoutes) {
        $m = [regex]::Match($Uri, $r.P)
        if ($m.Success) { return & $r.R $m }
    }
    throw "No Graph route for $Uri"
}

function J { param([string]$Json) $Json | ConvertFrom-Json }
function Get-SpRa { param($Title, $Login, $Type, [string[]]$Roles) [pscustomobject]@{ Member = [pscustomobject]@{ Title = $Title; LoginName = $Login; Email = ''; PrincipalType = $Type }; RoleDefinitionBindings = @($Roles | ForEach-Object { [pscustomobject]@{ Name = $_ } }) } }

function Get-TestSpResponse {
    param([string]$Url, $Connection)
    $site = $Connection.Url
    $name = ($site -split '/')[-1]
    switch -Regex ($Url) {
        '_api/web\?\$select=AssociatedOwnerGroup' { return [pscustomobject]@{ AssociatedOwnerGroup = [pscustomobject]@{ Title = "$name Owners" }; AssociatedMemberGroup = [pscustomobject]@{ Title = "$name Members" }; AssociatedVisitorGroup = $null } }
        'siteusers\?' { if ($name -eq 'Finance') { return J '{"value":[{"Title":"Global Admin","LoginName":"i:0#.f|membership|ghost@contoso.se","Email":"ghost@contoso.se","PrincipalType":1}]}' } return J '{"value":[]}' }
        'sitegroups\?' { return J "{`"value`":[{`"Id`":3,`"Title`":`"$name Members`"},{`"Id`":4,`"Title`":`"$name Owners`"},{`"Id`":9,`"Title`":`"SharingLinks.abc.Flexible.def`"}]}" }
        'getbyid\(3\)/users' {
            switch ($name) {
                'Finance' { return J "{`"value`":[{`"Title`":`"Finance Members`",`"LoginName`":`"c:0o.c|federateddirectoryclaimprovider|$script:G1`",`"Email`":`"`",`"PrincipalType`":4},{`"Title`":`"Cecilia`",`"LoginName`":`"i:0#.f|membership|cecilia_ext.com#ext#@contoso.onmicrosoft.com`",`"Email`":`"c@ext.com`",`"PrincipalType`":1}]}" }
                # Public team: Microsoft adds 'Everyone except external users' to the associated Members group by default.
                'IT' { return J "{`"value`":[{`"Title`":`"IT Members`",`"LoginName`":`"c:0o.c|federateddirectoryclaimprovider|$script:G2`",`"Email`":`"`",`"PrincipalType`":4},{`"Title`":`"Everyone except external users`",`"LoginName`":`"c:0-.f|rolemanager|spo-grid-all-users/111`",`"Email`":`"`",`"PrincipalType`":4}]}" }
                'Finance-Payroll' { return J '{"value":[{"Title":"Anna","LoginName":"i:0#.f|membership|anna@contoso.se","Email":"anna@contoso.se","PrincipalType":1},{"Title":"Bo","LoginName":"i:0#.f|membership|bo@contoso.se","Email":"bo@contoso.se","PrincipalType":1}]}' }
                default { return J '{"value":[{"Title":"Bo","LoginName":"i:0#.f|membership|bo@contoso.se","Email":"bo@contoso.se","PrincipalType":1}]}' }
            }
        }
        'getbyid\(4\)/users' { if ($name -in 'Finance', 'IT') { $g = if ($name -eq 'IT') { $script:G2 } else { $script:G1 }; return J "{`"value`":[{`"Title`":`"Owners`",`"LoginName`":`"c:0o.c|federateddirectoryclaimprovider|${g}_o`",`"Email`":`"`",`"PrincipalType`":4}]}" } return J '{"value":[{"Title":"Anna","LoginName":"i:0#.f|membership|anna@contoso.se","Email":"anna@contoso.se","PrincipalType":1}]}' }
        'getbyid\(9\)/users' { return J '{"value":[{"Title":"Anna","LoginName":"i:0#.f|membership|anna@contoso.se","Email":"anna@contoso.se","PrincipalType":1}]}' }
        '/Archive/_api/web\?' { return J '{"Url":"x","HasUniqueRoleAssignments":false}' }
        '_api/web\?' { return J '{"Url":"x","HasUniqueRoleAssignments":true}' }
        '_api/web/roleassignments' {
            $ras = @((Get-SpRa "$name Members" "$name Members" 8 'Edit'), (Get-SpRa "$name Owners" "$name Owners" 8 'Full Control'), (Get-SpRa 'SharingLinks.abc.Flexible.def' 'SharingLinks.abc.Flexible.def' 8 'Limited Access'))
            if ($name -eq 'Finance') { $ras += Get-SpRa 'Everyone except external users' 'c:0-.f|rolemanager|spo-grid-all-users/111' 4 'Read' }
            return [pscustomobject]@{ value = $ras }
        }
        '/Archive/_api/web/lists\?' { return J '{"value":[]}' }
        '_api/web/lists\?' { $c = if ($name -eq 'Finance') { 3 } else { 0 }; return J "{`"value`":[{`"Id`":`"L1`",`"Title`":`"Documents`",`"Hidden`":false,`"ItemCount`":$c,`"HasUniqueRoleAssignments`":false,`"RootFolder`":{`"ServerRelativeUrl`":`"/sites/$name/Shared Documents`"}}]}" }
        'items\?\$select=Id,FileRef' {
            if ($Url -notmatch 'skiptoken') { return [pscustomobject]@{ value = @([pscustomobject]@{ Id = 1; FileRef = '/a.docx'; HasUniqueRoleAssignments = $false }, [pscustomobject]@{ Id = 2; FileRef = '/secret.xlsx'; HasUniqueRoleAssignments = $true }); 'odata.nextLink' = "$site/_api/web/lists(guid'L1')/items?`$select=Id,FileRef&`$skiptoken=2" } }
            return J '{"value":[{"Id":3,"FileRef":"/c.docx","HasUniqueRoleAssignments":false}]}'
        }
        'items\(2\)/roleassignments' { return [pscustomobject]@{ value = @((Get-SpRa 'SharingLinks.abc.AnonymousView.xyz' 'SharingLinks.abc.AnonymousView.xyz' 8 'Read'), (Get-SpRa 'Unknown' 'i:0#.f|membership|ghost@contoso.se' 1 'Contribute')) } }
        '/Archive/_api/web/webs' { return J '{"value":[]}' }
        '_api/web/webs' { if ($name -eq 'Finance') { return J "{`"value`":[{`"Url`":`"$site/Archive`"}]}" } return J '{"value":[]}' }
    }
    throw "No SharePoint route for $Url"
}
