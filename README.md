# Poor Man's Dynamic Groups

Syncs every non-guest user in an Entra ID tenant into a **static** security group,
on a schedule, from an Azure Automation runbook. It exists so you can get
dynamic-group behaviour without Entra ID P1 licensing (dynamic membership rules
are a P1 feature; static group membership is not).

Everything lives in [runbook.ps1](runbook.ps1).

---

## How it works

Each run is a full reconciliation, not an incremental update:

1. Authenticate to Microsoft Graph using the Automation Account's
   system-assigned managed identity.
2. Read the group's **current** direct user members.
3. Read the **target** set — all non-guest users in the tenant.
4. Compute the set difference in both directions (`toAdd`, `toRemove`).
5. Run safety checks. Abort the whole run if either trips.
6. Apply adds, then removals.

Because it is a set difference, the first run against an empty group and the
five-hundredth run against a fully-populated group take the identical code path.
There is no "initial seed" mode to get wrong.

The script is idempotent: a run with nothing to change writes nothing.

### Why raw REST instead of the Graph SDK

Only `Az.Accounts` is imported. All Graph calls are `Invoke-RestMethod` against
`https://graph.microsoft.com/v1.0`. This deliberately avoids importing the
`Microsoft.Graph.*` module set into an Automation Account, which is slow to
install, large, and prone to version conflicts between submodules.

---

## Prerequisites

### Azure Automation Account

- A **system-assigned managed identity** enabled on the Automation Account.
- The `Az.Accounts` module imported. Nothing else is required.
- PowerShell 7.x runtime recommended.

### Microsoft Graph application permissions

Grant these to the Automation Account's managed identity as **application**
(app-only) permissions, with admin consent:

| Permission | Why |
|---|---|
| `User.Read.All` | Enumerate tenant users to build the target set |
| `GroupMember.ReadWrite.All` | Add and remove group members |

> **Making the managed identity an *owner* of the group is not enough.**
> Group ownership grants no rights to enumerate tenant users, and
> owner-implied write access is inconsistent for app-only callers. Grant the
> two application permissions above.

Granting Graph app roles to a managed identity has no portal UI. From a machine
with `Microsoft.Graph` installed:

```powershell
Connect-MgGraph -Scopes 'AppRoleAssignment.ReadWrite.All','Application.Read.All'

$miObjectId = '<managed identity object id>'   # the Automation Account's MI
$graphSp    = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"

foreach ($role in 'User.Read.All','GroupMember.ReadWrite.All') {
    $appRole = $graphSp.AppRole | Where-Object { $_.Value -eq $role -and $_.AllowedMemberTypes -contains 'Application' }
    New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $miObjectId `
        -PrincipalId $miObjectId -ResourceId $graphSp.Id -AppRoleId $appRole.Id
}
```

### The target group

Must be a **static / assigned-membership** group. If the group already has a
dynamic membership rule, Entra owns its membership and these writes will be
rejected or reverted.

---

## Parameters

| Parameter | Type | Default | Purpose |
|---|---|---|---|
| `-GroupId` | string | `01f57884-ff6d-40ef-aad1-ea4813246e99` | Object ID of the group to sync. **Change this default, or always pass it explicitly.** |
| `-WhatIf` | switch | off | Standard PowerShell dry run. Walks the full apply path and prints each write it would make, without making it. |
| `-ReportOnly` | switch | off | Stops right after the diff, prints the object ids on both sides, writes nothing. |
| `-IncludeDisabledUsers` | switch | off | By default only `accountEnabled eq true` users are targets, so disabling an account removes it from the group on the next run. Set this to keep disabled accounts as members. |
| `-MaxRemovalPercent` | int (1–100) | `20` | Abort if a single run would remove more than this share of current members. |
| `-MinExpectedUsers` | int | `1` | Abort if the tenant user query returns fewer than this many users. |

### Two dry-run modes

Both write nothing; they differ in how far they get and what they print.

`-WhatIf` is standard `ShouldProcess` support. It runs the whole script,
including the apply path, and prints one line per batched write it would have
issued:

```
Adding members...
What if: Performing the operation "Add 1 member(s): u11" on target "group <id>".
Removing members...
What if: Performing the operation "Remove 1 member(s): u10" on target "group <id>".
```

`-ReportOnly` returns immediately after the diff and safety checks, printing the
full add/remove id lists in one go. It is the easier one to drive from an
Automation schedule, being an ordinary switch parameter.

Two implementation notes, if you edit the script:

- **Authentication is forced to run under `-WhatIf`.** `Connect-AzAccount` and
  `Disable-AzContextAutosave` both declare `SupportsShouldProcess`, so a
  `-WhatIf` run would otherwise skip the sign-in and then fail on
  `Get-AzAccessToken`. Both are called with `-WhatIf:$false`.
- **Every Graph write is gated explicitly.** `Invoke-RestMethod` has no
  `ShouldProcess` of its own, so `-WhatIf` cannot suppress a REST call for you.
  `Add-GroupMembers` and `Remove-GroupMembers` each check
  `$PSCmdlet.ShouldProcess(...)` per 20-item chunk. Any new write must do the
  same or it will fire during a dry run.

`ConfirmImpact` is deliberately left at the default (Medium), so `-Confirm`
never prompts unattended — an interactive prompt would hang an Automation job
until the run times out.

---

## The safety guards

Both guards `throw`, which aborts the run before any write. They exist to stop a
malformed or partially-failed user query from wiping the group.

**`-MinExpectedUsers`** — checked immediately after the tenant query, before the
diff. If Graph returns fewer users than this floor, the run aborts. The default
of `1` only catches a total enumeration failure. **Set this to a realistic
number for your tenant** (e.g. `450` for a 500-user tenant); the default is
close to no protection at all.

**`-MaxRemovalPercent`** — checked after the diff. If removals would exceed this
percentage of current members, the run aborts with the exact counts. Note it is
skipped when the group is currently empty (`$currentIds.Count -gt 0`), so a
first run seeding an empty group is never blocked. Set to `100` to disable.

If a guard fires legitimately — a large planned offboarding, say — re-run once
with a raised threshold, then put it back.

---

## Behaviour details worth knowing

**Only direct *user* members are touched.** The current-members query is scoped
to `/members/microsoft.graph.user`, so nested groups, service principals, and
devices in the group are invisible to the diff and are never removed.

**Guests are excluded via `userType ne 'Guest'`, not `userType eq 'Member'`.**
Some directory-synced accounts have a null `userType`; a `ne 'Guest'` filter
includes them, an `eq 'Member'` filter would silently drop them. This negation
filter is why the request sends `ConsistencyLevel: eventual` and `$count=true` —
Graph advanced query support is mandatory for `ne` on `/users`.

**Paging.** `Get-GraphCollection` follows `@odata.nextLink` to the end and
returns one flat list. Page size is `$top=999`.

**Batching.** Adds use `members@odata.bind` on a `PATCH` to the group, 20 users
per call. Removals have no bulk endpoint, so they go through `/$batch` as 20
`DELETE` requests per call.

**Retries.** `Invoke-GraphRequest` retries HTTP 429 and 5xx up to 5 attempts
with exponential backoff (2, 4, 8, 16 seconds). Anything else throws immediately.

**Add failures are fatal; removal failures are not.** A `PATCH` chunk is
all-or-nothing — one bad ID fails the whole 20-user chunk and the script stops.
Individual failures inside a removal `$batch`, by contrast, are logged as
warnings and the run continues.

---

## Running it

Dry run first, always:

```powershell
.\runbook.ps1 -GroupId '<group-object-id>' -WhatIf
.\runbook.ps1 -GroupId '<group-object-id>' -ReportOnly
```

Real run, with guards tuned for the tenant:

```powershell
.\runbook.ps1 -GroupId '<group-object-id>' -MinExpectedUsers 450 -MaxRemovalPercent 10
```

### As a scheduled Automation runbook

1. Create a PowerShell 7.x runbook in the Automation Account, paste in
   `runbook.ps1`, publish it.
2. Attach a schedule (hourly or daily is typical — this is a poll, so group
   membership lags reality by up to one interval).
3. Set schedule parameters: at minimum `GROUPID`, plus `MINEXPECTEDUSERS` and
   `MAXREMOVALPERCENT` if you are overriding the defaults.

All progress goes to `Write-Output`, so it lands in the job's Output stream;
per-item removal failures and retry notices go to `Write-Warning`.

---

## Troubleshooting

| Symptom | Cause |
|---|---|
| `Authorization_RequestDenied` on `/users` | Missing `User.Read.All` app role, or admin consent not granted. Ownership of the group does not substitute. |
| `Authorization_RequestDenied` on the group PATCH | Missing `GroupMember.ReadWrite.All`. |
| `Request_UnsupportedQuery` / advanced query error | The `ConsistencyLevel: eventual` header or `$count=true` was lost from the users request. Both are required by the `ne` filter. |
| `One or more added object references already exist` | A user in the add chunk is already a member — the group changed between the read and the write. Re-run; the next diff will be correct. |
| Run aborts on `-MaxRemovalPercent` | Intended. Check the reported counts against `-ReportOnly` output before raising the threshold. |
| Group membership does not stick | The group has a dynamic membership rule. Convert it to assigned membership. |
| `Get-AzAccessToken` token issues | Handled — the script unwraps the `SecureString` returned by Az.Accounts 5.x and falls back to plain text on older versions. |

---

## Limitations

- **Polling, not eventing.** Membership is only as fresh as the schedule
  interval. A new hire is not in the group until the next run.
- **One group, one rule.** The target set is hardcoded to "non-guest users".
  Anything more selective means editing `$userFilter`.
- **No transcript of *who* changed.** Output reports counts and, under
  `-ReportOnly`, object IDs — not UPNs. The queries already select
  `userPrincipalName`, so extending the logging is straightforward.
- **Large tenants.** Every run enumerates every user. Fine into the tens of
  thousands; past that, consider Graph delta queries.
