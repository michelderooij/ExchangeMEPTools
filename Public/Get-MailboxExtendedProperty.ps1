function Get-MailboxExtendedProperty {
    <#
    .SYNOPSIS
    Lists named-property definitions from the mailbox's Exchange Store mapping.

    .DESCRIPTION
    Queries the authoritative Exchange Store named-property mapping for the selected mailbox. Use PropertyName to filter mappings before optional type resolution. Item-level EWS search is provided separately by Search-MailboxExtendedProperty and requires typed property definitions.

    .PARAMETER Identity
    Mailbox SMTP address, alias, GUID, or another identifier resolvable by Get-Mailbox.

    .PARAMETER PropertyName
    Optional PowerShell wildcard pattern to select named properties. When combined with ResolvePropertyType, only matching mappings are probed.

    .PARAMETER ResolvePropertyType
    Probe Store message rows for observed MAPI property types in automatic batches of up to 5 mappings, with individual fallback. This can issue multiple mailbox-wide diagnostic queries per selected mapping; unresolved mappings remain untyped.

    .EXAMPLE
    Get-MailboxExtendedProperty -Identity user@contoso.com
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('MailboxIdentity', 'PrimarySmtpAddress')]
        [string]$Identity,

        [ValidateNotNullOrEmpty()]
        [string]$PropertyName,

        [switch]$ResolvePropertyType
    )

    process {
        Get-StoreNamedProperty -Identity $Identity -PropertyName $PropertyName -ResolvePropertyType:$ResolvePropertyType
    }
}