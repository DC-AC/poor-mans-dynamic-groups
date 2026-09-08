<#
.SYNOPSIS
    Poor-man's dynamic group. Syncs all non-guest users in the tenant into a
    static security group, using a set difference so that the first run
    (empty group) and every later run use the same code path.

.DESCRIPTION
    Designed for Azure Automation with a system-assigned managed identity.
    Requires only the Az.Accounts module; all Graph work is raw REST, which
    avoids the Microsoft.Graph SDK module sprawl in Automation Accounts.

    Managed identity needs these Microsoft Graph application permissions:
        User.Read.All              - to enumerate tenant users
        GroupMember.ReadWrite.All  - to add/remove group members

    Group ownership alone is NOT sufficient: it gives no rights to enumerate
    users, and owner-implied writes are inconsistent for app-only callers.

.NOTES
    Run with -WhatIf first. Seriously.

    Two dry-run modes, deliberately kept separate:
        -WhatIf      walks the whole code path and prints every write it
                     would make, in the order it would make them.
        -ReportOnly  stops right after the diff and prints the object ids
                     on both sides. Easier to wire into an Automation
                     schedule, since it is a plain switch parameter.

    ConfirmImpact is left at the default (Medium) on purpose. Raising it to
    High would make -Confirm prompt, and an interactive prompt hangs an
    Automation job until the run times out.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter()]
    [string] $GroupId = '01f57884-ff6d-40ef-aad1-ea4813246e99',

    # Report intended changes without writing anything.
    [Parameter()]
    [switch] $ReportOnly,

    # Include accounts where accountEnabled = false.
    [Parameter()]
    [switch] $IncludeDisabledUsers,

    # Safety valve: abort if a single run would remove more than this
    # percentage of current members. Guards against a bad user query
    # emptying the group. Set to 100 to disable.
    [Parameter()]
    [ValidateRange(1, 100)]
    [int] $MaxRemovalPercent = 20,

    # Abort if the tenant user query returns fewer than this many users.
    # Another guard against a partial/failed enumeration.
    [Parameter()]
    [int] $MinExpectedUsers = 1
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$GraphBase = 'https://graph.microsoft.com/v1.0'

#region ---------------------------------------------------------- Auth

Write-Output 'Authenticating with managed identity...'

# Both of these declare SupportsShouldProcess, so -WhatIf on the script would
# otherwise skip the sign-in and leave Get-AzAccessToken with no context.
# Authenticating is a read; force it to happen even in a -WhatIf run.
Disable-AzContextAutosave -Scope Process -WhatIf:$false | Out-Null
$null = Connect-AzAccount -Identity -WhatIf:$false

$tokenResponse = Get-AzAccessToken -ResourceUrl 'https://graph.microsoft.com'

# Az.Accounts 5.x returns a SecureString; earlier versions return plain text.
if ($tokenResponse.Token -is [System.Security.SecureString]) {
    $accessToken = [System.Net.NetworkCredential]::new('', $tokenResponse.Token).Password
}
else {
    $accessToken = $tokenResponse.Token
}

$Headers = @{
    Authorization    = "Bearer $accessToken"
    'Content-Type'   = 'application/json'
    # Required for the 'ne' filter and $count on /users.
    ConsistencyLevel = 'eventual'
}

#endregion

#region ------------------------------------------------------- Helpers

function Invoke-GraphRequest {
    <#
        Thin wrapper adding retry on throttling (429) and transient 5xx.
    #>
    param(
        [Parameter(Mandatory)] [string] $Uri,
        [string] $Method = 'GET',
        $Body,
        [int] $MaxAttempts = 5
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            $params = @{
                Uri     = $Uri
                Method  = $Method
                Headers = $Headers
            }
            if ($null -ne $Body) {
                $params.Body = ($Body | ConvertTo-Json -Depth 10 -Compress)
            }
            return Invoke-RestMethod @params
        }
        catch {
            $status = $null
            if ($_.Exception.PSObject.Properties.Name -contains 'Response' -and $_.Exception.Response) {
                $status = [int] $_.Exception.Response.StatusCode
            }

            $retryable = ($status -eq 429) -or ($status -ge 500 -and $status -le 599)

            if (-not $retryable -or $attempt -eq $MaxAttempts) {
                throw "Graph $Method $Uri failed (HTTP $status): $($_.Exception.Message)"
            }

            $delay = [Math]::Pow(2, $attempt)
            Write-Warning "HTTP $status on attempt $attempt. Retrying in $delay s..."
            Start-Sleep -Seconds $delay
        }
    }
}

function Get-GraphCollection {
    <#
        Follows @odata.nextLink and returns every page as one flat list.
    #>
    param([Parameter(Mandatory)] [string] $Uri)

    $results = [System.Collections.Generic.List[object]]::new()
    $next = $Uri

    while ($next) {
        $page = Invoke-GraphRequest -Uri $next

        if ($page.PSObject.Properties.Name -contains 'value') {
            foreach ($item in $page.value) { $results.Add($item) }
        }

        $next = if ($page.PSObject.Properties.Name -contains '@odata.nextLink') {
            $page.'@odata.nextLink'
        } else {
            $null
        }
    }

    return $results
}

function New-IdSet {
    <#
        Builds a case-insensitive HashSet of ids from a collection that may be
        empty or null. Passing a null collection to the HashSet constructor
        throws, so items are added one at a time instead.
    #>
    param($Items)

    $set = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )

    foreach ($item in @($Items)) {
        if ($null -ne $item -and $null -ne $item.id) {
            [void] $set.Add([string] $item.id)
        }
    }

    return , $set
}

function Add-GroupMembers {
    <#
        Adds up to 20 members per call using members@odata.bind on the group.
        Far cheaper than one POST per user.

        Declares SupportsShouldProcess so that -WhatIf on the script reaches
        this write. Invoke-RestMethod has no ShouldProcess of its own, so the
        PATCH has to be gated explicitly rather than suppressed for us.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)] [string] $GroupId,
        [Parameter(Mandatory)] [string[]] $UserIds
    )

    $added = 0
    for ($i = 0; $i -lt $UserIds.Count; $i += 20) {
        $chunk = $UserIds[$i..([Math]::Min($i + 19, $UserIds.Count - 1))]

        $action = "Add $($chunk.Count) member(s): $($chunk -join ', ')"
        if (-not $PSCmdlet.ShouldProcess("group $GroupId", $action)) { continue }

        $body = @{
            'members@odata.bind' = @(
                $chunk | ForEach-Object { "https://graph.microsoft.com/v1.0/directoryObjects/$_" }
            )
        }

        Invoke-GraphRequest -Uri "$GraphBase/groups/$GroupId" -Method 'PATCH' -Body $body | Out-Null
        $added += $chunk.Count
        Write-Output "  added $added / $($UserIds.Count)"
    }
}

function Remove-GroupMembers {
    <#
        Removals have no bulk endpoint, so batch 20 DELETEs via /$batch.

        Gated by ShouldProcess for the same reason as Add-GroupMembers.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)] [string] $GroupId,
        [Parameter(Mandatory)] [string[]] $UserIds
    )

    $removed = 0
    for ($i = 0; $i -lt $UserIds.Count; $i += 20) {
        $chunk = $UserIds[$i..([Math]::Min($i + 19, $UserIds.Count - 1))]

        $action = "Remove $($chunk.Count) member(s): $($chunk -join ', ')"
        if (-not $PSCmdlet.ShouldProcess("group $GroupId", $action)) { continue }

        $requests = @()
        $n = 0
        foreach ($id in $chunk) {
            $n++
            $requests += @{
                id     = "$n"
                method = 'DELETE'
                url    = "/groups/$GroupId/members/$id/`$ref"
            }
        }

        $response = Invoke-GraphRequest -Uri "$GraphBase/`$batch" -Method 'POST' -Body @{ requests = $requests }

        foreach ($r in $response.responses) {
            if ($r.status -ge 400) {
                Write-Warning "  removal failed (id $($r.id), HTTP $($r.status)): $($r.body.error.message)"
            }
        }

        $removed += $chunk.Count
        Write-Output "  processed $removed / $($UserIds.Count)"
    }
}

#endregion

#region ---------------------------------------------- Gather both sets

Write-Output "Reading current members of $GroupId..."

# Direct members only. Filtering to microsoft.graph.user means nested groups
# and service principals in the group are left alone rather than "removed".
$currentMembers = Get-GraphCollection -Uri "$GraphBase/groups/$GroupId/members/microsoft.graph.user?`$select=id,userPrincipalName&`$top=999"

$currentIds = New-IdSet -Items $currentMembers

Write-Output "  $($currentIds.Count) user members currently in group."

Write-Output 'Reading tenant users...'

# 'ne Guest' rather than 'eq Member' so that accounts with a null userType
# (which some synced accounts have) are still included.
$userFilter = "userType ne 'Guest'"
if (-not $IncludeDisabledUsers) {
    $userFilter += ' and accountEnabled eq true'
}

$targetUsers = Get-GraphCollection -Uri "$GraphBase/users?`$filter=$([uri]::EscapeDataString($userFilter))&`$select=id,userPrincipalName,userType&`$count=true&`$top=999"

$targetIds = New-IdSet -Items $targetUsers

Write-Output "  $($targetIds.Count) non-guest users found."

if ($targetIds.Count -lt $MinExpectedUsers) {
    throw "Only $($targetIds.Count) users returned, below the -MinExpectedUsers floor of $MinExpectedUsers. Aborting rather than risk emptying the group."
}

#endregion

#region ------------------------------------------------------- The diff

# @() keeps these as arrays when the result is empty. A bare [string[]] cast
# of an empty pipeline yields $null, which then breaks .Count under StrictMode.
$toAdd = @($targetIds | Where-Object { -not $currentIds.Contains($_) })
$toRemove = @($currentIds | Where-Object { -not $targetIds.Contains($_) })

Write-Output ''
Write-Output "To add:    $($toAdd.Count)"
Write-Output "To remove: $($toRemove.Count)"
Write-Output ''

if ($toRemove.Count -gt 0 -and $currentIds.Count -gt 0) {
    $removalPercent = [Math]::Round(($toRemove.Count / $currentIds.Count) * 100, 1)
    if ($removalPercent -gt $MaxRemovalPercent) {
        throw "This run would remove $($toRemove.Count) of $($currentIds.Count) members ($removalPercent%), exceeding -MaxRemovalPercent of $MaxRemovalPercent. Aborting. Re-run with a higher threshold if this is expected."
    }
}

if ($ReportOnly) {
    Write-Output '-ReportOnly specified. No changes written.'
    if ($toAdd.Count)    { Write-Output "Would add: $($toAdd -join ', ')" }
    if ($toRemove.Count) { Write-Output "Would remove: $($toRemove -join ', ')" }
    return
}

#endregion

#region ---------------------------------------------------- Apply

if ($toAdd.Count -gt 0) {
    Write-Output 'Adding members...'
    Add-GroupMembers -GroupId $GroupId -UserIds $toAdd
}

if ($toRemove.Count -gt 0) {
    Write-Output 'Removing members...'
    Remove-GroupMembers -GroupId $GroupId -UserIds $toRemove
}

Write-Output ''
if ($WhatIfPreference) {
    Write-Output "-WhatIf run complete. Nothing was written. The group would contain $($targetIds.Count) users."
}
else {
    Write-Output "Sync complete. Group should now contain $($targetIds.Count) users."
}

#endregion