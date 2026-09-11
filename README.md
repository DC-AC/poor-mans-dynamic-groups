# Poor Man's Dynamic Groups

Poor Man's Dynamic Groups syncs every non-guest user in an Entra ID tenant into
a **static** security group. Azure Automation runs it on a schedule. It gives
you dynamic-group behaviour without Entra ID P1 licensing. Dynamic membership
rules need P1. Static group membership does not.

Everything lives in [runbook.ps1](runbook.ps1).

---

## How it works

Each run is a full reconciliation, not an incremental update:

1. Authenticate to Microsoft Graph. Use the Automation Account's
   system-assigned managed identity.
2. Read the group's **current** direct user members.
3. Read the **target** set: all non-guest users in the tenant.
4. Compute the set difference in both directions (`toAdd`, `toRemove`).
5. Run the safety checks. Abort the whole run if a check fails.
6. Apply the adds. Then apply the removals.

The first run against an empty group and the five-hundredth run against a
fully-populated group take the identical code path. The set difference is the
reason. The script has no separate "initial seed" mode to get wrong.

The script is idempotent. A run with nothing to change writes nothing.

### Why raw REST instead of the Graph SDK

The script imports only `Az.Accounts`. Every Graph call uses
`Invoke-RestMethod` against `https://graph.microsoft.com/v1.0`. The script
avoids the `Microsoft.Graph.*` module set on purpose. That module set is large,
slow to install, and prone to version conflicts between submodules.

---

## Prerequisites

### Azure Automation Account

- Enable a **system-assigned managed identity** on the Automation Account.
- Import the `Az.Accounts` module. You need nothing else.
- The PowerShell 7.x runtime is the recommended one.

### Microsoft Graph application permissions

Grant these two **application** (app-only) permissions to the Automation
Account's managed identity. Grant admin consent for them.

| Permission | Why |
|---|---|
| `User.Read.All` | Enumerate the tenant users to build the target set |
| `GroupMember.ReadWrite.All` | Add and remove group members |

> **Do not make the managed identity only an *owner* of the group.** Group
> ownership gives no rights to enumerate the tenant users. Owner-implied write
> access is also inconsistent for app-only callers. Grant the two application
> permissions above.

Entra ID has no portal UI to grant Graph app roles to a managed identity. Run
the block below from a machine that has `Microsoft.Graph` installed.

> **You need the Privileged Role Administrator role or the Global
> Administrator role to run this block.** Application Administrator and Cloud
> Application Administrator are *not* enough. Entra ID bars both roles from
> consenting to Microsoft Graph application permissions. Such consent is a path
> to self-granting `RoleManagement.ReadWrite.Directory`. The block fails with
> `403 Authorization_RequestDenied: Insufficient privileges` when an
> Application Administrator runs it.
>
> Run `Disconnect-MgGraph` after you activate the role through PIM. Then
> connect again. An existing token does not pick up a newly activated role.

```powershell
Connect-MgGraph -Scopes 'AppRoleAssignment.ReadWrite.All','Application.Read.All'

$miObjectId = '<managed identity object id>'   # the Automation Account's MI
$graphSp    = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"

foreach ($role in 'User.Read.All','GroupMember.ReadWrite.All') {
    $appRole = $graphSp.AppRoles | Where-Object { $_.Value -eq $role -and $_.AllowedMemberTypes -contains 'Application' }
    if (-not $appRole) { throw "App role '$role' not found on the Graph service principal" }
    New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $miObjectId `
        -PrincipalId $miObjectId -ResourceId $graphSp.Id -AppRoleId $appRole.Id
}
```

### The target group

The target group must use **static (assigned) membership**. A dynamic
membership rule makes Entra ID the owner of the membership. Entra ID then
rejects or reverts these writes.

---

## Parameters

| Parameter | Type | Default | Purpose |
|---|---|---|---|
| `-GroupId` | string | `01f57884-ff6d-40ef-aad1-ea4813246e99` | The object ID of the group to sync. **Change this default, or always pass the parameter explicitly.** |
| `-WhatIf` | switch | off | A standard PowerShell dry run. It walks the full apply path. It prints each write it would make, and makes none of them. |
| `-ReportOnly` | switch | off | Stops right after the diff. Prints the object IDs on both sides. Writes nothing. |
| `-IncludeDisabledUsers` | switch | off | By default, only `accountEnabled eq true` users are targets. A disabled account therefore leaves the group on the next run. Set this switch to keep disabled accounts as members. |
| `-MaxRemovalPercent` | int (1–100) | `20` | Aborts a single run that would remove more than this share of the current members. |
| `-MinExpectedUsers` | int | `1` | Aborts the run if the tenant user query returns fewer users than this number. |

### Two dry-run modes

Both modes write nothing. They differ in how far they get and what they print.

`-WhatIf` is standard `ShouldProcess` support. It runs the whole script,
including the apply path. It prints one line for each batched write it would
have issued:

```
Adding members...
What if: Performing the operation "Add 1 member(s): u11" on target "group <id>".
Removing members...
What if: Performing the operation "Remove 1 member(s): u10" on target "group <id>".
```

`-ReportOnly` returns straight after the diff and the safety checks. It prints
the full add and remove ID lists in one go. It is a plain switch parameter, so
it is the easier mode to drive from an Automation schedule.

Two implementation notes, for when you edit the script:

- **The script forces the authentication to run under `-WhatIf`.**
  `Connect-AzAccount` and `Disable-AzContextAutosave` both declare
  `SupportsShouldProcess`. A `-WhatIf` run would otherwise skip the sign-in and
  then fail on `Get-AzAccessToken`. The script calls both with `-WhatIf:$false`.
- **The script gates every Graph write explicitly.** `Invoke-RestMethod` has no
  `ShouldProcess` of its own, so `-WhatIf` cannot suppress a REST call for you.
  `Add-GroupMembers` and `Remove-GroupMembers` each check
  `$PSCmdlet.ShouldProcess(...)` once per 20-item chunk. Any new write must do
  the same. A write that skips the check fires during a dry run.

`ConfirmImpact` stays at the default (Medium) on purpose. `-Confirm` therefore
never prompts in an unattended run. An interactive prompt hangs an Automation
job until the run times out.

---

## The safety guards

Both guards `throw`. The throw aborts the run before any write. The guards stop
a malformed or partially failed user query from wiping the group.

**`-MinExpectedUsers`** runs immediately after the tenant query, before the
diff. The run aborts if Graph returns fewer users than this floor. The default
of `1` catches only a total enumeration failure. **Set this parameter to a
realistic number for your tenant**, for example `450` in a 500-user tenant. The
default gives close to no protection.

**`-MaxRemovalPercent`** runs after the diff. The run aborts with the exact
counts if the removals would exceed this percentage of the current members. The
script skips the guard when the group is empty (`$currentIds.Count -gt 0`). The
guard therefore never blocks a first run that seeds an empty group. Set the
value to `100` to turn the guard off.

A guard can fire on a legitimate change, such as a large planned offboarding.
Re-run once with a raised threshold. Then put the threshold back.

---

## Behaviour details worth knowing

**The script touches only direct *user* members.** The query for the current
members uses `/members/microsoft.graph.user`. Nested groups, service
principals, and devices in the group stay invisible to the diff. The script
never removes them.

**The script excludes guests with `userType ne 'Guest'`, not
`userType eq 'Member'`.** Some directory-synced accounts have a null
`userType`. A `ne 'Guest'` filter includes those accounts. An `eq 'Member'`
filter drops them silently. This negation filter is the reason the request sends
`ConsistencyLevel: eventual` and `$count=true`. The `ne` operator on `/users`
needs Graph advanced query support.

**Paging.** `Get-GraphCollection` follows `@odata.nextLink` to the end. It
returns one flat list. The page size is `$top=999`.

**Batching.** The adds use `members@odata.bind` on a `PATCH` to the group, 20
users per call. The removals have no bulk endpoint. They go through `/$batch`
as 20 `DELETE` requests per call.

**Retries.** `Invoke-GraphRequest` retries HTTP 429 and 5xx up to 5 attempts,
with exponential backoff (2, 4, 8, 16 seconds). Anything else throws
immediately.

**Add failures are fatal. Removal failures are not.** A `PATCH` chunk is
all-or-nothing. One bad ID fails the whole 20-user chunk, and the script stops.
The script logs individual failures inside a removal `$batch` as warnings. The
run then continues.

---

## Running it

Always do a dry run first:

```powershell
.\runbook.ps1 -GroupId '<group-object-id>' -WhatIf
.\runbook.ps1 -GroupId '<group-object-id>' -ReportOnly
```

For a real run, tune the guards for the tenant:

```powershell
.\runbook.ps1 -GroupId '<group-object-id>' -MinExpectedUsers 450 -MaxRemovalPercent 10
```

### As a scheduled Automation runbook

1. Create a PowerShell 7.x runbook in the Automation Account.
2. Paste `runbook.ps1` into the runbook.
3. Publish the runbook.
4. Attach a schedule. Hourly or daily is typical. The script is a poll, so
   group membership lags reality by up to one interval.
5. Set the schedule parameters. `GROUPID` is the minimum. Add
   `MINEXPECTEDUSERS` and `MAXREMOVALPERCENT` if you override the defaults.

All progress goes to `Write-Output`, so it lands in the job's Output stream.
Per-item removal failures and retry notices go to `Write-Warning`.

---

## Troubleshooting

| Symptom | Cause |
|---|---|
| `Authorization_RequestDenied` on `/users` | The `User.Read.All` app role is missing, or nobody granted admin consent. Ownership of the group is not a substitute. |
| `Authorization_RequestDenied` on the group PATCH | The `GroupMember.ReadWrite.All` app role is missing. |
| `Request_UnsupportedQuery`, or another advanced query error | The users request lost the `ConsistencyLevel: eventual` header or `$count=true`. The `ne` filter needs both. |
| `One or more added object references already exist` | A user in the add chunk is already a member. The group changed between the read and the write. Re-run the script. The next diff is correct. |
| The run aborts on `-MaxRemovalPercent` | This abort is intended. Check the reported counts against the `-ReportOnly` output before you raise the threshold. |
| Group membership does not stick | The group has a dynamic membership rule. Convert the group to assigned membership. |
| `Get-AzAccessToken` token issues | The script handles this. It unwraps the `SecureString` that Az.Accounts 5.x returns, and falls back to plain text on older versions. |

---

## Limitations

- **Polling, not eventing.** Membership is only as fresh as the schedule
  interval. A new hire is not in the group until the next run.
- **One group, one rule.** The script hardcodes the target set to "non-guest
  users". Edit `$userFilter` for anything more selective.
- **No transcript of *who* changed.** The output reports counts. Under
  `-ReportOnly` it also reports object IDs, but never UPNs. The queries already
  select `userPrincipalName`, so the logging is easy to extend.
- **Large tenants.** Every run enumerates every user. This is fine into the
  tens of thousands. Past that, consider Graph delta queries.
