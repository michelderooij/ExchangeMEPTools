# ExchangeMEPTools

![PowerShell Gallery Downloads](https://img.shields.io/powershellgallery/dt/ExchangeMEPTools)
![GitHub Release](https://img.shields.io/github/v/release/michelderooij/ExchangeMEPTools)
![GitHub forks](https://img.shields.io/github/forks/michelderooij/ExchangeMEPTools?style=lat-square)
![GitHub forks](https://img.shields.io/github/stars/michelderooij/ExchangeMEPTools?style=flat-square)

ExchangeMEPTools is a PowerShell module for managing MAPI named properties in Exchange Server on-premises mailboxes. Use it to:

- Inventory a mailbox's named-property mappings.
- Assess named-property quota usage.
- Search mailbox items for known named properties.
- Remove a named property from selected items.

Inventory reads the Exchange Store mapping table. Search and removal use EWS and require a typed property definition; Store inventory does not provide the property value type.

## Requirements

- Windows PowerShell 5.1 and Exchange Management Shell access.
- An on-premises Exchange mailbox with EWS enabled.
- The EWS Managed API for searching and removing properties [[https://eightwone.com/2020/10/05/ews-webservices-managed-api/](learn more)]

Import the module:

```powershell
Import-Module .\ExchangeMEPTools.psd1
```

## Examples

### Inventory and assess a mailbox

```powershell
$properties = Get-MailboxExtendedProperty -Identity user@contoso.com
$properties | Format-Table

$properties | Test-MailboxExtendedProperty
```

To select mappings by name before inventory or type resolution, use `-PropertyName` with a PowerShell wildcard, such as `-PropertyName 'x-example-*'`. `-ResolvePropertyType` can probe selected mappings for an observed type; it may take time, and some mappings may remain unresolved.

An inventory filtered with `-PropertyName` is partial. When passed to `Test-MailboxExtendedProperty`, all assessment checks are reported as incomplete. Use an unfiltered inventory for a complete assessment.

### Search for a known property

Provide a verified property-set ID, property name, and value type:

```powershell
Connect-EWSSession -Identity user@contoso.com -UseDefaultCredentials

$definition = [pscustomobject]@{
    PropertySetId = '00020329-0000-0000-c000-000000000046'
    PropertyName  = 'x-example-property'
    PropertyType  = 'String'
}

$matches = Search-MailboxExtendedProperty `
    -Identity user@contoso.com `
    -MailboxExtendedProperty $definition
```

Replace the example property details with the definition you intend to search for. EWS needs the correct type; inventory alone cannot supply it.

### Remove a property from matching items

Review the preview, then run the confirmed removal:

```powershell
$matches | Remove-MailboxExtendedProperty -WhatIf
$matches | Remove-MailboxExtendedProperty -Confirm -Verify
```

Removal deletes only the named property from an item, not the item itself. Review the matches and use `-WhatIf` before making changes.

## FAQ

- **Store inventory is unavailable:** Run from Exchange Management Shell and check that the account has the required Exchange permissions. Store diagnostic access may require an elevated shell.
- **The EWS DLL cannot be found or loaded:** See [lib/README.md](lib/README.md), or provide its location with `-EwsManagedApiPath`.
- **EWS cannot connect:** Check that EWS is enabled, the mailbox identity and permissions are correct, and Autodiscover or the HTTPS endpoint is available.
- **Search or removal requires a property type:** Store mappings do not include value types. Supply a verified, EWS-supported `PropertyType`.
- **A mapping has no resolved type:** `-ResolvePropertyType` reports observed types only; it does not guess when no value is found.

## Learn more

For more information and usage examples, visit [https://eightwone.com](https://eightwone.com/2026/11/19/exchange-named-property-exhaustion/)).

## Credits

Thanks to David Barrett (MSFT) for examples of using `ManagedStoreDiagnosticFunctions` to collect named property information. See [PowerShell-EWS-Scripts](https://github.com/David-Barrett-MS/PowerShell-EWS-Scripts).
