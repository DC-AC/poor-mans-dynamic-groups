<#
    Poor-man's dynamic group sync — managed identity permission setup
    ------------------------------------------------------------------
    REQUIRES: Global Administrator (or Privileged Role Administrator).

    Application Administrator is NOT sufficient. That role is explicitly barred
    from consenting to Microsoft Graph application permissions and from
    assigning directory roles, so both halves below fail with
    403 Authorization_RequestDenied when run by an app admin.

    What this grants the managed identity, and why each is needed:

      1. Groups Administrator, scoped to an administrative unit containing
         ONLY the target group. Lets the runbook add/remove members of that
         one group. Chosen over the tenant-wide GroupMember.ReadWrite.All
         app permission, which would allow membership writes on every group
         in the tenant.

      2. User.Read.All (Graph app role, tenant-wide). The runbook enumerates
         every non-guest user in the tenant to build the target set, so this
         one cannot be scoped down.

    Tenant : fa35d6ce-bd5a-4ca6-86c7-08e18dfa8674
    MI     : 5486ddf3-e5b5-4d99-9631-4bfee1666d97  (Automation Account system-assigned MI)
    Group  : e447e1f4-dde4-492d-9e8b-ea7391ad874d
#>

$ErrorActionPreference = 'Stop'

$TenantId    = 'fa35d6ce-bd5a-4ca6-86c7-08e18dfa8674'
$MiObjectId  = '5486ddf3-e5b5-4d99-9631-4bfee1666d97'
$GroupId     = 'e447e1f4-dde4-492d-9e8b-ea7391ad874d'
$AuName      = 'AU-DynamicGroupSync'

Connect-MgGraph -TenantId $TenantId -Scopes @(
    'AdministrativeUnit.ReadWrite.All'
    'RoleManagement.ReadWrite.Directory'
    'AppRoleAssignment.ReadWrite.All'
    'Application.Read.All'
    'Group.Read.All'
)

# Fail early with a clear message rather than a 403 five steps in.
$ctx = Get-MgContext
Write-Host "Signed in as $($ctx.Account) on tenant $($ctx.TenantId)" -ForegroundColor Cyan

#region 0 -- Sanity-check the target group -----------------------------------

$group = Get-MgGroup -GroupId $GroupId -Property 'id,displayName,groupTypes,membershipRule'
Write-Host "Target group: $($group.DisplayName)" -ForegroundColor Cyan

if ($group.GroupTypes -contains 'DynamicMembership') {
    throw "Group '$($group.DisplayName)' has dynamic membership. Entra owns its membership and the runbook's writes will be rejected or reverted. Convert it toassigned membership first."
}

#endregion

#region 1 -- Administrative unit containing only the target group ------------

$au = Get-MgDirectoryAdministrativeUnit -Filter "displayName eq '$AuName'" -ErrorAction SilentlyContinue |
        Select-Object -First 1

if (-not $au) {
    $au = New-MgDirectoryAdministrativeUnit -BodyParameter @{
        displayName = $AuName
        description = 'Scope boundary for the poor-mans-dynamic-groups automation MI. Contains only the group it syncs.'
    }
    Write-Host "Created AU '$AuName' ($($au.Id))" -ForegroundColor Green
} else {
    Write-Host "AU '$AuName' already exists ($($au.Id))" -ForegroundColor Yellow
}

$auMembers = Get-MgDirectoryAdministrativeUnitMember -AdministrativeUnitId $au.Id -All
if ($auMembers.Id -notcontains $GroupId) {
    New-MgDirectoryAdministrativeUnitMemberByRef -AdministrativeUnitId $au.Id -BodyParameter @{
        '@odata.id' = "https://graph.microsoft.com/v1.0/groups/$GroupId"
    }
    Write-Host "Added group to AU" -ForegroundColor Green
} else {
    Write-Host "Group already in AU" -ForegroundColor Yellow
}

#endregion

#region 2 -- Groups Administrator, scoped to that AU -------------------------

$roleDef = Get-MgRoleManagementDirectoryRoleDefinition -Filter "displayName eq 'Groups Administrator'"
if (-not $roleDef) { throw "Could not resolve the 'Groups Administrator' role definition." }

$scope = "/administrativeUnits/$($au.Id)"

$existing = Get-MgRoleManagementDirectoryRoleAssignment `
                -Filter "principalId eq '$MiObjectId' and roleDefinitionId eq '$($roleDef.Id)'" -All |
            Where-Object { $_.DirectoryScopeId -eq $scope }

if (-not $existing) {
    New-MgRoleManagementDirectoryRoleAssignment -BodyParameter @{
        principalId      = $MiObjectId
        roleDefinitionId = $roleDef.Id
        directoryScopeId = $scope
    } | Out-Null
    Write-Host "Assigned Groups Administrator scoped to $scope" -ForegroundColor Green
} else {
    Write-Host "Scoped Groups Administrator assignment already present" -ForegroundColor Yellow
}

#endregion

#region 3 -- User.Read.All app role (tenant-wide, unavoidable) ---------------

$graphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"

# NOTE: the property is AppRoles (plural). $graphSp.AppRole silently yields
# nothing, and the empty AppRoleId then fails as "Cannot convert the literal ''
# to the expected type 'Edm.Guid'".
$appRole = $graphSp.AppRoles |
    Where-Object { $_.Value -eq 'User.Read.All' -and $_.AllowedMemberTypes -contains 'Application' }
if (-not $appRole) { throw "App role 'User.Read.All' not found on the Graph service principal." }

$haveRole = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $MiObjectId -All |
    Where-Object { $_.AppRoleId -eq $appRole.Id -and $_.ResourceId -eq $graphSp.Id }

if (-not $haveRole) {
    New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $MiObjectId `
        -PrincipalId $MiObjectId -ResourceId $graphSp.Id -AppRoleId $appRole.Id | Out-Null
    Write-Host "Granted User.Read.All" -ForegroundColor Green
} else {
    Write-Host "User.Read.All already granted" -ForegroundColor Yellow
}

#endregion

#region 4 -- Verify ----------------------------------------------------------

Write-Host "`n--- Result ---" -ForegroundColor Cyan

Write-Host "`nGraph app roles on the MI:"
Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $MiObjectId -All |
    ForEach-Object {
        $r = $graphSp.AppRoles | Where-Object Id -eq $_.AppRoleId
        [pscustomobject]@{ Permission = $r.Value; Resource = $_.ResourceDisplayName }
    } | Format-Table -AutoSize

Write-Host "Directory role assignments on the MI:"
Get-MgRoleManagementDirectoryRoleAssignment -Filter "principalId eq '$MiObjectId'" -All |
    ForEach-Object {
        $d = Get-MgRoleManagementDirectoryRoleDefinition -UnifiedRoleDefinitionId $_.RoleDefinitionId
        [pscustomobject]@{ Role = $d.DisplayName; Scope = $_.DirectoryScopeId }
    } | Format-Table -AutoSize

Write-Host "AU '$AuName' contains:"
Get-MgDirectoryAdministrativeUnitMember -AdministrativeUnitId $au.Id -All |
    ForEach-Object { [pscustomobject]@{ Id = $_.Id; Type = $_.AdditionalProperties['@odata.type'] } } |
    Format-Table -AutoSize

Write-Host "Done. Allow a few minutes before the first runbook execution -- app role" -ForegroundColor Cyan
Write-Host "and directory role changes do not appear in already-issued MI tokens." -ForegroundColor Cyan

#endregion
PS /Users/joey/dcac/poor-mans-dynamic-groups>  