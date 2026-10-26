# ExchangeNamedProperties

PowerShell module for inventorying, searching, assessing, and removing MAPI named properties in Exchange Server on-premises mailboxes.

## Scope

- Exchange Server on-premises only.
- Windows PowerShell 5.1.
- EWS Managed API for mailbox item access.
- The Store inventory method uses Exchange's `Get-StoreQuery` diagnostic function to read the mailbox named-property mapping table.
- No Microsoft Graph, Exchange Online, OAuth, or direct database diagnostic client is used by the EWS workflows.
- Search and removal operate only on supplied, typed extended-property definitions. Store mappings do not include the MAPI value type or item occurrence count.

`Get-MailboxExtendedProperty` is Store-only and returns the mapping rows for the successfully resolved mailbox store. EWS is used for item-level search and removal with caller-supplied, typed property definitions; it does not discover unknown named-property definitions.

## Requirements

- Windows PowerShell 5.1.
- Exchange Management Shell cmdlets for Store inventory and mailbox alias/GUID resolution. On import, if the required cmdlets are missing, the module tries to load `Bin\RemoteExchange.ps1` from the Exchange installation path and runs `Connect-ExchangeServer -Auto`.
- Microsoft.Exchange.WebServices.dll for EWS search/removal. The DLL is not bundled. See [lib/README.md](lib/README.md).
- An Exchange Server on-premises EWS endpoint enabled for the mailbox.

When EMS cmdlets are already available, import displays a detection notice and the module version/command summary. If they are missing, the module looks for the Exchange install path in `$exinstallpath`, `$env:ExchangeInstallPath`, and the Exchange v15 registry setup key, then attempts to load EMS. A successful automatic connection displays a loading notice and suppresses the summary; if EMS cannot be loaded, a dependency warning is shown and import completes. To suppress the automatic EMS load attempt (for example, in build scripts), use `Import-Module ExchangeNamedProperties.psd1 -ArgumentList @{ SkipExchangeShell = $true }`.

## Authentication and Permissions

EWS uses the current Windows security context by default through `UseDefaultCredentials`. An explicit `PSCredential` can be supplied to `Connect-EWSSession` or a mailbox cmdlet; it is used through EWS `WebCredentials`. This is not OAuth and does not imply HTTP Basic authentication is enabled on the server.

For the current user's mailbox, the Windows identity must have mailbox access and EWS must be enabled. Access to another mailbox requires the appropriate Exchange mailbox permissions. EWS impersonation is opt-in with `-Impersonate` and requires the Exchange role/permissions configured by the administrator; the module does not elevate or grant permissions.

For Exchange Server on-premises, the account that runs the module needs the `ApplicationImpersonation` management role to use `-Impersonate`. An Exchange administrator can assign it in the Exchange Management Shell. Set `$account` to the identity that will run the module:

```powershell
$account = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
New-ManagementRoleAssignment -Role ApplicationImpersonation -User $account
```

This unscoped assignment can allow impersonation across the organization. In production, have the Exchange administrator restrict the assignment to the intended recipient scope and follow your organization's approval process. Then opt in when connecting:

```powershell
Connect-EWSSession -EwsUrl ex3 -UseDefaultCredentials -Impersonate
```

The Store method uses the current Exchange Management Shell security context. Depending on the Exchange installation and assigned permissions, the shell running the Store method may need to be elevated. The module will not attempt to elevate itself.

## Connection Configuration

Connection settings may be applied for the current module session:

```powershell
Connect-EWSSession -Identity user@contoso.com -UseDefaultCredentials
```

Without an explicit `EwsUrl`, the EWS Managed API uses Autodiscover. Only HTTPS redirection targets are accepted. To test with a known endpoint, supply it locally through an environment variable rather than embedding it in scripts or documentation:

```powershell
Connect-EWSSession -EwsUrl $env:EXCHANGE_EWS_URL -UseDefaultCredentials
```

`EwsUrl` also accepts a bare host or host/path shorthand. A bare host defaults to the standard EWS endpoint, while a host/path is prefixed with HTTPS:

```powershell
Connect-EWSSession -EwsUrl ex3 -UseDefaultCredentials
# Uses https://ex3/EWS/Exchange.asmx

Connect-EWSSession -EwsUrl ex3/EWS/Exchange.asmx -UseDefaultCredentials
# Uses https://ex3/EWS/Exchange.asmx
```

Explicit HTTP URLs remain rejected; use HTTPS.

Disconnect clears module-session settings and cached services:

```powershell
Disconnect-EWSSession
```

`Get-EWSSession` reports the configured endpoint, authentication method and identity (never the credential secret), impersonation settings, retry default, trace settings, and cached EWS services. `-RetryCount` on `Connect-EWSSession` sets the default for Search and Remove; a command-level `-RetryCount` overrides it:

```powershell
Connect-EWSSession -EwsUrl ex3 -UseDefaultCredentials -RetryCount 5 -TraceEnabled -TraceFilePath C:\Logs\ews-trace.log
Get-EWSSession
```

To access the live EWS Managed API service after connecting with `-Identity` or creating a cached service through Search:

```powershell
$service = (Get-EWSSession).Session
```

Select a specific cached mailbox service by SMTP address, alias, or GUID:

```powershell
$service = (Get-EWSSession -Identity user@contoso.com).Session
```

With `-Identity`, the report includes only that mailbox's cached service and endpoint. An uncached mailbox returns an error rather than creating a new connection.

`Session` is null when no service is cached or multiple mailbox services are cached. For multiple mailboxes, select the appropriate `MailboxIdentity` from `(Get-EWSSession).CachedServices` and use that entry's `Session` property. These are live references: changes affect subsequent module operations, and the objects include authentication configuration. Do not serialize or log the full report; select only diagnostic fields when exporting it.

EWS does not expose Exchange Management Role assignments; use Exchange Management Shell to inspect RBAC assignments.

The former `Connect-ExchangeNamedProperties` and `Disconnect-ExchangeNamedProperties` names remain as compatibility aliases. The EWS DLL can be supplied through `-EwsManagedApiPath`, placed in the current working directory, module directory, module `lib` directory, or installed in a standard EWS Managed API location. Certificate validation is enabled by default. Safe connection traces do not include SOAP payloads or credentials.

For controlled testing only, `-TrustAll` explicitly accepts certificate errors for the authority in the supplied HTTPS `EwsUrl`. It cannot be used with Autodiscover, emits a warning, and remains active for that authority until `Disconnect-EWSSession`. This modifies the process-wide .NET callback, but the callback is host-and-port scoped and delegates every other authority to the previous validation policy. Do not use it in production; install a valid certificate matching the EWS endpoint instead.

## Store Inventory

```powershell
$properties = Get-MailboxExtendedProperty -Identity user@contoso.com
```

Store mapping output includes the property-set GUID and either the property name or numeric dispid when available. `NamedPropCount` is the raw mapping-row count from the Store query, repeated on each result; it can include malformed rows that are skipped from individual output. `Test-MailboxExtendedProperty` uses that raw count for the total quota comparison. By default, Store mappings do not include the MAPI value type. Use `Get-MailboxExtendedProperty -ResolvePropertyType` to opt into per-mapping Store message probes; this can issue multiple mailbox-wide diagnostic queries. A detected type is reported as `PropertyType` plus its MAPI hex `PropertyTypeCode` (for example, `String` and `001f`). A mapping with no observed message value remains unresolved and is not assigned a guessed type. This identifies observed usage, not every type theoretically associated with the mapping.

Store queries reuse the mounted server and worker process for one inventory invocation when the installed `Get-StoreQuery` supports `-Server` and `-ProcessId`. This avoids a database-status lookup for every type probe. A stale RPC endpoint triggers one refresh and retry per query; older diagnostic commands retain database-mode querying. Worker information is not cached between invocations. Type probes remain serial, so resolving every mapping can still be expensive, especially for properties not observed on messages.

With `-ResolvePropertyType`, the initial progress shows only `Resolving type for selected Store mapping X of Y`. After the first mapping completes, updates include average processed mappings per second (`type/sec`) and estimated time remaining. The ETA shows `Calculating...` until at least 10 mappings, 10% of the selection (rounded up), and 30 seconds have been observed; it then uses `MM:SS`. This avoids extrapolating a full run from the first few fast batches, but later slowdowns can still increase the estimate. Small or fast runs may finish before a numeric ETA is shown; completion always reports `00:00`. The timer starts before batch probing and excludes mailbox/database lookup and mapping retrieval. Updates occur after the first mapping, every 10 mappings, and the final mapping. Unresolved mappings still count as processed; the estimate is approximate because probe costs vary. Minutes do not wrap at one hour, and the estimate is `--:--` if throughput cannot yet be measured.

`-ResolvePropertyType` automatically groups probes in batches of up to 5 mappings; no batch-size parameter is required or exposed. Smaller groups may reduce query-width failures but can increase diagnostic calls; they do not prevent all fallback. Probes follow the same candidate-type order as individual probing, beginning with Unicode (`001f`). Each batch selects one matching message and resolves only properties actually observed on that message, then repeats for remaining candidates. Only candidates with a successful zero-row result advance to the next type; ambiguous and failed probes retain individual fallback, so an earlier candidate is not silently skipped. Singleton groups also use individual fallback. Without `-ResolvePropertyType`, inventory retrieval does not run type probes.

Successful zero-row batch results are reused during fallback only within the current invocation, database, and mailbox partition. Missing fields, `DBNull`, and the Store `NULL` sentinel in a returned row do not establish mailbox-wide absence. Empty strings, zero, and false are valid observations. A failed group larger than two is split into smaller groups while other queued groups remain eligible. After three failed batch calls across the invocation, batching is disabled and unresolved mappings use individual probes. Failures never establish absence. There is no cross-invocation cache; observations and negative results describe the queries' execution times, not an atomic mailbox snapshot.

Add `-Verbose` to see notices when batch queries fail, smaller groups are attempted, batching is disabled for the invocation, or unresolved mappings enter individual fallback. These notices do not change the normal progress display or returned property objects.

Batch probing runs before the per-mapping output loop and uses the same `Resolving type for selected Store mapping X of Y` status, speed, and ETA as individual probing. The completed count advances as batch queries observe property types, then continues during individual fallback without counting batch observations twice or resetting the count during output. Batch completions can advance several mappings at once; updates occur when the count crosses a 10-mapping boundary, after the first completion, and at completion. Until any mapping completes, only the initial plain status is shown. The synchronous Store request cannot provide a progress heartbeat while it is waiting; this progress reporting does not change query execution time or add a timeout.

Batching can reduce diagnostic round trips but wider predicates may still require substantial Store work. For example, five mappings absent from all 21 candidate probes need 21 successful batch queries instead of 105 individual queries, excluding metadata queries and retries. This is a query-count comparison, not an elapsed-time guarantee. Batching does not assume one matching message contains every property, prove that a mapping has only one possible type, or add multivalued-type support. Type resolution remains opt-in through `-ResolvePropertyType`.

For name-based cleanup, filter mappings before resolving types. `-PropertyName` uses PowerShell wildcard matching and does not change the mailbox-wide `NamedPropCount` value:

```powershell
$properties = Get-MailboxExtendedProperty `
	-Identity user@contoso.com `
	-PropertyName 'x-sogeti-*' `
	-ResolvePropertyType
```

## Search

To search item values, EWS needs a typed property definition. You can pass a typed definition directly:

```powershell
$definition = [pscustomobject]@{
PropertySetId = '00020329-0000-0000-c000-000000000046'
PropertyName = 'x-example-property'
PropertyType = 'String'
}

$matches = Search-MailboxExtendedProperty `
-Identity user@contoso.com `
-MailboxExtendedProperty $definition `
-PageSize 100
```

Search pages the primary mailbox's visible folders recursively by default. It excludes archive, recoverable-items folders, and associated items unless explicitly requested. Results contain item/folder identity, requested metadata, the property value, and a UTC timestamp. Item bodies and attachments are not requested. `-ResultSize` limits total matches. Individual folder errors are returned as structured error objects and processing continues.

Pipeline use:

```powershell
$definitions | Search-MailboxExtendedProperty -Identity user@contoso.com
$properties | Search-MailboxExtendedProperty
```

If `MailboxExtendedProperty` is omitted and `Identity` is supplied, Search loads the mailbox's Store mappings. If `Identity` is omitted, it is inferred from the input object's `PrimarySmtpAddress` or `MailboxIdentity`. `PropertyName` wildcards expand against Store mappings. Store does not provide a value `PropertyType`, so untyped mappings return structured `PropertyTypeRequired` results rather than item values. Add a verified type before EWS value search.

## Risk Analysis

```powershell
Test-MailboxExtendedProperty -Identity user@contoso.com -Threshold 1.0 -SelectFirst 10
```

Store inventory results can also be piped directly; `Identity` is inferred from their `MailboxIdentity` field:

```powershell
$properties = Get-MailboxExtendedProperty -Identity user@contoso.com
$properties | Test-MailboxExtendedProperty
```

To assess whether an on-premises mailbox fits the Exchange Online named-property quotas before migration, use the Exchange Online reference profile:

```powershell
$properties | Test-MailboxExtendedProperty -ExchangeOnlineQuotas
```

This mode compares against the published Exchange Online limits rather than the source mailbox's `NamedPropertiesCountQuota`, and labels results as a migration-target assessment. It does not query `Get-MailboxStatistics` for the source quota.

`Test-MailboxExtendedProperty` prints a colored console report for the total, PublicStrings, InternetHeaders, per-GUID namespace, and global first-five-character prefix quota checks, followed by a summary. Namespace and prefix checks are sorted by count and limited by `SelectFirst`. Use `-PassThru` to also return structured risk result objects and the normalized input definitions for further processing; those objects include `PercentageUsed` (0-100), `PercentOfTotalQuota`, and the integer `ThresholdValue`.

The risk checks follow the latest CSS-Exchange script's five quota scenarios. When `Get-MailboxStatistics` returns `NamedPropertiesCountQuota`, that runtime total is used and the derived limits are 60% for Internet Headers, 20% for Public Strings, 20% per namespace GUID, and 10% for each global first-five-character property-name prefix. The namespace list includes the special namespaces as well as custom GUIDs. Prefix quota logic in CSS-Exchange applies to `StringProperty` definitions; Store mappings do not provide property types, so Store-derived prefix counts are candidates and have `Status = Unknown` rather than claiming exact string-only usage.

If the on-premises Exchange shell does not return `NamedPropertiesCountQuota`, the Microsoft [named-property quota table](https://learn.microsoft.com/en-us/troubleshoot/exchange/outlook-issues/named-property-quota-exceeded#cause) is shown only as an Exchange Online reference: 16,384 total, 9,830 Internet Headers, 3,277 Public Strings, 3,277 per namespace, and 1,638 per global prefix. Reference-only percentage results have `Status = Unknown`; configure version-specific Exchange Server limits in `Data/NamedPropertyLimits.json` to enable server-specific statuses. `Threshold` defaults to 1.0 (100%) and `SelectFirst` defaults to 10. Incomplete inventories cannot produce a mailbox-wide `Pass`.

## Safe Removal

Removal affects only the named property on an item. It never deletes the item. `-WhatIf` and `-Confirm` are supported; default conflict resolution is `AutoResolve`. Use `-Verify` to rebind and check that the property is absent.

```powershell
$matches | Remove-MailboxExtendedProperty -Identity user@contoso.com -WhatIf
$matches | Remove-MailboxExtendedProperty -Identity user@contoso.com -Confirm -Verify
$matches | Remove-MailboxExtendedProperty -WhatIf
```

When `Identity` is omitted, Remove infers it from each piped search result. It still requires `ItemId` and a selected typed property definition; inventory rows from `Get-MailboxExtendedProperty` alone cannot identify an item to modify. The command validates mailbox ownership and returns one structured result per input. Conflicts, access failures, validation failures, and transient server-busy responses are reported separately; a property already absent from its item is reported as successful.

All piped inputs are processed. Each input is validated and confirmed as it arrives; approved removals are then sent per mailbox in groups of `-BatchSize` (default 100): one batched EWS bind, one batched update for the changed items, and with `-Verify` one batched verification bind. Several properties on the same item are removed with a single update. A progress bar is updated after each batch, for example `Batch 2 of 5 | Processed 200 of 450 | Removed 180 | Already absent 15 | Skipped 0 | Failed 5`. Items are bound fresh, so the `ChangeKey` on search results is not used; `-ConflictResolution` applies to the bound item version.

## Cleanup Workflow

The requested workflow shape is:

```powershell
$properties = Get-MailboxExtendedProperty `
	-Identity user@contoso.com `
	-PropertyName 'x-sogeti-*' `
	-ResolvePropertyType

$matches = $properties |
Search-MailboxExtendedProperty -Identity user@contoso.com

$matches |
Remove-MailboxExtendedProperty -Identity user@contoso.com -WhatIf

$matches |
Remove-MailboxExtendedProperty -Identity user@contoso.com -Confirm
```

Store does not supply `PropertyType`; `-ResolvePropertyType` probes only the selected name mappings and reports an observed type where found. Unresolved mappings remain untyped and cannot be passed to EWS Search or Remove. The workflow is complete only for the selected Store mapping rows and the items/folders scanned; no EWS operation proves item-wide property-map completeness.

## Troubleshooting

- **EWS DLL not found or load failure:** pass `-EwsManagedApiPath` and verify the assembly is compatible with Windows PowerShell 5.1.
- **Autodiscover failure:** verify mailbox SMTP resolution, DNS, EWS virtual directory availability, and HTTPS redirection.
- **Access denied / EWS disabled:** verify mailbox EWS access, current Windows identity, and Exchange permissions.
- **Store provider unavailable:** run in the Exchange Management Shell. `Connect-EWSSession` checks that `Get-Mailbox`, `Get-MailboxDatabase`, and `Get-ExchangeDiagnosticInfo` are available. If `Get-StoreQuery` is not loaded, it loads `Scripts\ManagedStoreDiagnosticFunctions.ps1` from `$exinstallpath` (or `$env:ExchangeInstallPath`).
- **Store method permission error:** depending on local configuration, run the shell elevated or ask an Exchange administrator to grant the required rights.
- **Search/removal rejects type:** Store inventory does not discover value types. Supply a correct EWS-supported `PropertyType` from an authoritative source.
- **Throttling/server busy:** the module retries bounded transient responses; reduce page size or retry count if the server remains busy.

## Security Considerations

- Do not put credentials, EWS endpoints, or mailbox content in committed configuration.
- The user-supplied EWS test endpoint is local test configuration only; provide it through `EXCHANGE_EWS_URL` when testing and do not publish its literal value.
- The module does not bypass certificate validation by default. `-TrustAll` is an explicit, insecure test-only exception for the configured EWS authority; disconnect immediately after use.
- Review every removal result. Use `-WhatIf` first and use `-Verify` for high-risk cleanup.
- Store diagnostic access can require elevated shell permissions. The module does not self-elevate.

## Validation

Pester is not used. Run `./build.ps1` in Windows PowerShell 5.1 to validate the module and execute `Tests/Run-Assertions.ps1`. Live EWS checks are opt-in and require administrator-provided environment configuration; the assertion suite does not require an Exchange server.
