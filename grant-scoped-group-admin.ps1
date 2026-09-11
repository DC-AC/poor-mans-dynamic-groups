<#
.SYNOPSIS
    Grants the sync managed identity the Groups Administrator role scoped to a
    single group, via an administrative unit containing only that group.

.DESCRIPTION
    The alternative is the tenant-wide GroupMember.ReadWrite.All Graph
    application permission, which lets the managed identity rewrite membership
    on every group in the tenant. Entra has no way to scope a Graph app role to
    one resource, so the scoping has to come from the directory-role side
    instead: an administrative unit is the only scope boundary that group
    membership writes respect.

    The result is a managed identity that can manage exactly one group and is
    not a group admin anywhere else in the tenant.

    This script does NOT grant User.Read.All. The runbook enumerates every
    non-guest user in the tenant to build its target set, and that read cannot
    be scoped to an AU, so it stays a separate tenant-wide app-role grant. See
    the README.

.NOTES
    REQUIRES: Global Administrator or Privileged Role Administrator.

    Application Administrator is NOT sufficient. That role is barred from
    assigning directory roles and from consenting to Microsoft Graph
    application permissions; both fail with 403 Authorization_RequestDenied.
    A PIM-eligible role must be activated first, and an existing Graph token
    does not pick up a newly activated role -- Disconnect-MgGraph and
    reconnect.

    Every step is idempotent, so a re-run after a partial failure is safe.

    Scoped role assignments do not appear in tokens already issued to the
    managed identity. Allow a few minutes before the first runbook run.

.EXAMPLE
    ./grant-scoped-group-admin.ps1 `
        -GroupId    'e447e1f4-dde4-492d-9e8b-ea7391ad874d' `
        -MiObjectId '5486ddf3-e5b5-4d99-9631-4bfee1666d97' -WhatIf

.EXAMPLE
    ./grant-scoped-group-admin.ps1 `
        -GroupId    'e447e1f4-dde4-492d-9e8b-ea7391ad874d' `
        -MiObjectId '5486ddf3-e5b5-4d99-9631-4bfee1666d97'
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    # Object ID of the group the managed identity will be allowed to manage.
    [Parameter(Mandatory)]
    [ValidatePattern('^[0-9a-fA-F-]{36}$')]
    [string] $GroupId,

    # Object ID of the Automation Account's system-assigned managed identity.
    # This is the service principal object ID, not the client/app ID.
    [Parameter(Mandatory)]
    [ValidatePattern('^[0-9a-fA-F-]{36}$')]
    [string] $MiObjectId,

    # Administrative unit to create or reuse as the scope boundary.
    [Parameter()]
    [string] $AdministrativeUnitName = 'AU-DynamicGroupSync'
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# Built-in Groups Administrator. For built-in roles the role definition id is
# the same as the well-known template id, so this is stable across tenants and
# avoids a displayName lookup that shifts with directory language.
$GroupsAdminRoleId = 'fdd7a751-b60b-444a-984c-02652fe8fa1c'

#region ------------------------------------------------------- Connect

$requiredScopes = @(
    'AdministrativeUnit.ReadWrite.All'
    'RoleManagement.ReadWrite.Directory'
    'Group.Read.All'
)

$context = Get-MgContext
if (-not $context -or @($requiredScopes | Where-Object { $_ -notin $context.Scopes }).Count -gt 0) {
    Connect-MgGraph -Scopes $requiredScopes | Out-Null
    $context = Get-MgContext
}

Write-Output "Signed in as $($context.Account) on tenant $($context.TenantId)."

#endregion

#region -------------------------------------------- Validate the target

# Fail here rather than after creating an AU we would not want to keep.
$group = Get-MgGroup -GroupId $GroupId -Property 'id,displayName,groupTypes'

if ($group.GroupTypes -contains 'DynamicMembership') {
    throw "Group '$($group.DisplayName)' has a dynamic membership rule. Entra owns its membership and the runbook's writes would be reverted. Convert it to assigned membership first."
}

Write-Output "Target group: $($group.DisplayName) ($GroupId)"

#endregion

#region ------------------------------------------ Administrative unit

# Escape embedded quotes so a name like "Joey's AU" does not break the filter.
$nameFilter = $AdministrativeUnitName -replace "'", "''"
$au = @(Get-MgDirectoryAdministrativeUnit -Filter "displayName eq '$nameFilter'" -All) |
        Select-Object -First 1

if (-not $au) {
    if ($PSCmdlet.ShouldProcess($AdministrativeUnitName, 'Create administrative unit')) {
        $au = New-MgDirectoryAdministrativeUnit -BodyParameter @{
            displayName = $AdministrativeUnitName
            description = 'Scope boundary for the poor-mans-dynamic-groups managed identity. Contains only the group it syncs.'
        }
        Write-Output "Created administrative unit '$AdministrativeUnitName' ($($au.Id))."
    } else {
        # -WhatIf: no AU exists to report scope against, so stop here rather
        # than print a misleading plan built on a null id.
        Write-Output "WhatIf: would create administrative unit '$AdministrativeUnitName', add group $GroupId to it, and assign Groups Administrator over it to $MiObjectId."
        return
    }
} else {
    Write-Output "Reusing existing administrative unit '$AdministrativeUnitName' ($($au.Id))."
}

$memberIds = @(Get-MgDirectoryAdministrativeUnitMember -AdministrativeUnitId $au.Id -All |
                ForEach-Object { $_.Id })

if ($memberIds -notcontains $GroupId) {
    if ($PSCmdlet.ShouldProcess("group $GroupId", "Add to administrative unit '$AdministrativeUnitName'")) {
        New-MgDirectoryAdministrativeUnitMemberByRef -AdministrativeUnitId $au.Id -BodyParameter @{
            '@odata.id' = "https://graph.microsoft.com/v1.0/groups/$GroupId"
        }
        Write-Output 'Added the group to the administrative unit.'
    }
} else {
    Write-Output 'Group is already in the administrative unit.'
}

# The AU is the scope boundary, so anything else in it also becomes writable
# by the managed identity. Worth surfacing rather than discovering later.
$strays = @($memberIds | Where-Object { $_ -ne $GroupId })
if ($strays.Count -gt 0) {
    Write-Warning "Administrative unit '$AdministrativeUnitName' also contains $($strays.Count) other object(s). The managed identity will be able to manage those too: $($strays -join ', ')"
}

#endregion

#region ----------------------------------- Scoped role assignment

$scope = "/administrativeUnits/$($au.Id)"

$existing = @(Get-MgRoleManagementDirectoryRoleAssignment `
                -Filter "principalId eq '$MiObjectId' and roleDefinitionId eq '$GroupsAdminRoleId'" -All |
              Where-Object { $_.DirectoryScopeId -eq $scope })

if ($existing.Count -eq 0) {
    if ($PSCmdlet.ShouldProcess("managed identity $MiObjectId", "Assign Groups Administrator scoped to $scope")) {
        New-MgRoleManagementDirectoryRoleAssignment -BodyParameter @{
            principalId      = $MiObjectId
            roleDefinitionId = $GroupsAdminRoleId
            directoryScopeId = $scope
        } | Out-Null
        Write-Output "Assigned Groups Administrator scoped to $scope."
    }
} else {
    Write-Output 'Scoped Groups Administrator assignment is already in place.'
}

#endregion

#region ------------------------------------------------------- Verify

Write-Output "`nDirectory role assignments on the managed identity:"

Get-MgRoleManagementDirectoryRoleAssignment -Filter "principalId eq '$MiObjectId'" -All |
    ForEach-Object {
        $def = Get-MgRoleManagementDirectoryRoleDefinition -UnifiedRoleDefinitionId $_.RoleDefinitionId
        [pscustomobject]@{
            Role  = $def.DisplayName
            Scope = $_.DirectoryScopeId
        }
    } | Format-Table -AutoSize

Write-Output @"
A Scope of '/' means tenant-wide. The assignment added here should read
'$scope'.

Still required, and not granted by this script:
  User.Read.All  (Microsoft Graph application permission, tenant-wide)

Then dry-run the sync before scheduling it:
  ./runbook.ps1 -GroupId '$GroupId' -WhatIf
"@

#endregion
