# Changelog

Notable changes to ExchangeMEPTools are recorded here.

## [Unreleased]

### Added

- Renamed the module to `ExchangeMEPTools` (Mailbox Extended Property tools). Import `ExchangeMEPTools.psd1`; public command names and existing result-object type names are unchanged. The former `ExchangeNamedProperties` and `ExchangeMailboxExtendedPropertyCleaner` import filenames have been retired.
- Initial Windows PowerShell 5.1 module manifest and loader.
- Session-scoped EWS connection configuration with current Windows credentials by default, optional PSCredential, and HTTPS-only Autodiscover redirection validation.
- Store named-property mapping provider based on Exchange mailbox metadata and `Get-StoreQuery`.
- Store-only named-property inventory; EWS is reserved for item-level operations with known, typed property definitions.
- EWS Managed API installation guidance; the DLL is not bundled.
- Explicit testing-only `-TrustAll` support scoped to the authority in an explicit HTTPS EWS URL and restored by disconnect.
- Store diagnostic-script discovery from `$exinstallpath` or the `ExchangeInstallPath` environment variable, with Exchange Management Shell prerequisite checks.
- Store mapping results include the raw `NamedPropCount`, which is used for total named-property risk calculations.
- Risk analysis follows the CSS-Exchange five-scenario quota model, uses `NamedPropertiesCountQuota` when returned by on-premises `Get-MailboxStatistics`, and marks Store prefix candidates unknown when property types are unavailable.
- `Test-MailboxExtendedProperty -ExchangeOnlineQuotas` compares a mailbox against the Exchange Online quota profile for migration-readiness assessment without querying the current server quota.
- Store inventory reuses worker-process routing within each invocation to avoid repeated database-status lookups, with one refresh/retry for a stale RPC endpoint and database-mode fallback for older diagnostics commands.
- Automatic type-probe batching in groups of up to 5 mappings when `-ResolvePropertyType` is requested, reducing query width while retaining individual fallback.
- Batching follows the existing candidate-type order, with invocation-scoped reuse of successful empty probes, conservative individual fallback, and bounded group-splitting recovery disabled after three failed batch calls. The public result shape remains unchanged.
- Type-resolution progress displays average processed mappings per second and estimated minutes/seconds remaining, including batch-probe time.
- Numeric ETA is withheld during warm-up until at least 10 mappings, 10% of the selection, and 30 seconds have been observed, reducing optimistic estimates from early fast batches without concealing later slowdowns.
- Batch and individual resolution share the regular mapping-count, speed and ETA progress display, with increasing completion counts and no double counting of batch observations during output.
- Verbose-only notices report failed batch recovery, invocation-wide batching disablement, and individual type-resolution fallback.
- Removed the public `TypeProbeBatchSize` parameter in favor of automatic five-mapping batches; existing scripts must omit that argument. Inventory without type resolution is unchanged.
- Removed `ProfileTypeResolution` and its investigation-only sample cap, timing diagnostics, counters, and alternate query execution path. Regular progress speed and ETA remain available.
- `Remove-MailboxExtendedProperty` processes all inputs instead of returning `BatchSizeExceeded` beyond `-BatchSize`. `-BatchSize` now sets how many removals are sent in one batched EWS bind/update request (EWS `BindToItems`/`UpdateItems`, saved without sending message or meeting updates), and per-item EWS errors are mapped back to their inputs.
- `Remove-MailboxExtendedProperty` shows a per-batch progress bar with processed, removed, already absent, skipped, and failed counts.
- `Remove-MailboxExtendedProperty` now uses the EWS session configured through `Connect-EWSSession` and no longer accepts per-command endpoint, credential, or service configuration parameters.
- `Remove-MailboxExtendedProperty` now accepts search results through `-MailboxExtendedProperty`; the redundant `-MessagesWithExtendedProperty` parameter has been removed.
- Fixed removal of search results that carry a `ChangeKey`; the EWS Managed API has no public `ItemId(uniqueId, changeKey)` constructor, so items are now bound by unique ID.
- Module import now checks all required Exchange Management Shell cmdlets, displays a detection notice and command summary when EMS is already loaded, and attempts to load `Bin\RemoteExchange.ps1` and connect with `Connect-ExchangeServer -Auto` when needed. Successful automatic connections display a loading notice and suppress the summary; failures warn without aborting import. `Import-Module -ArgumentList @{ SkipExchangeShell = $true }` bypasses the attempt.
- `Get-ExchangeInstallPath` now falls back to the Exchange v15 registry `MsiInstallPath` when neither `$exinstallpath` nor `$env:ExchangeInstallPath` is set.

### Notes

- The Store method may require an elevated Exchange Management Shell, depending on local Exchange permissions and installation configuration. The module does not self-elevate.
- Store mapping rows do not supply property value types or item occurrence counts; those values remain unknown unless another provider establishes them.
- The supplied EWS endpoint is test configuration only and is not included in the module or this changelog.
- `-TrustAll` disables certificate validation for its configured EWS authority while enabled. Do not use it in production; prefer a correctly issued certificate.
