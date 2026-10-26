function Search-MailboxExtendedProperty {
    <#
    .SYNOPSIS
    Searches mailbox items for explicitly defined MAPI extended properties.

    .DESCRIPTION
    Traverses visible folders and uses EWS to retrieve only the requested item fields and supplied extended-property definitions. Identity can be inferred from Get-MailboxExtendedProperty output. If no definitions are supplied, Store mappings are loaded. Property-name wildcards are expanded against Store mappings. EWS still requires a known PropertyType; Store mappings without a type are reported, not guessed.

    .PARAMETER Identity
    Optional mailbox SMTP address, alias, GUID, or another identifier resolvable by Get-Mailbox. If omitted, it is inferred from input property objects.

    .PARAMETER MailboxExtendedProperty
    Optional named-property definitions from Get-MailboxExtendedProperty, or a property-name wildcard. Typed definitions are required for EWS value searches.

    .PARAMETER Folder
    Optional folder path or display name to search within the traversed hierarchy.

    .PARAMETER IncludeSubfolders
    Include recursively nested folders. Enabled by default; specify $false to search only top-level folders.

    .PARAMETER Archive
    Include the archive mailbox hierarchy.

    .PARAMETER PageSize
    Number of items requested per EWS page.

    .PARAMETER ResultSize
    Maximum total matches returned. Omit for no client-side result limit.

    .PARAMETER SearchDumpster
    Include the recoverable-items hierarchy.

    .PARAMETER IncludeAssociatedItems
    Search associated items in addition to ordinary items.

    .EXAMPLE
    $properties | Search-MailboxExtendedProperty -Identity user@contoso.com

    .EXAMPLE
    Search-MailboxExtendedProperty -Identity user@contoso.com -MailboxExtendedProperty $definition -PageSize 100
    #>
    [CmdletBinding()]
    param(
        [Alias('MailboxIdentity', 'PrimarySmtpAddress')]
        [string]$Identity,

        [Parameter(ValueFromPipeline = $true)]
        [Alias('PropertyDefinition')]
        [psobject[]]$MailboxExtendedProperty,

        [string]$Folder,

        [bool]$IncludeSubfolders = $true,

        [switch]$Archive,

        [ValidateRange(1, 1000)]
        [int]$PageSize = 100,

        [ValidateRange(1, 2147483647)]
        [int]$ResultSize,

        [switch]$SearchDumpster,

        [switch]$IncludeAssociatedItems,

        [ValidateRange(0, 10)]
        [int]$RetryCount,

        [ValidateRange(1, 600)]
        [int]$RequestTimeoutSeconds,

        [string]$EwsUrl,

        [string]$EwsManagedApiPath,

        [System.Management.Automation.PSCredential]$Credential,

        [switch]$UseDefaultCredentials,

        [string]$ExchangeVersion,

        [switch]$TraceEnabled,

        [string]$TraceFilePath,

        [switch]$Impersonate,

        [switch]$TrustAll
    )

    begin {
        $definitions = New-Object 'System.Collections.Generic.List[object]'
        $identityCandidates = New-Object 'System.Collections.Generic.List[string]'
        $searchTimestamp = [DateTime]::UtcNow.ToString('o', [Globalization.CultureInfo]::InvariantCulture)
        $matchCount = 0
        $seenMatches = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        if (-not $PSBoundParameters.ContainsKey('RetryCount')) {
            $RetryCount = if ($script:ConnectionOptions -and $script:ConnectionOptions.PSObject.Properties['RetryCount']) { [int]$script:ConnectionOptions.RetryCount } else { 3 }
        }
    }

    process {
        foreach ($definition in @($MailboxExtendedProperty)) {
            if ($null -ne $definition) {
                $definitions.Add($definition)
                $primaryAddressProperty = $definition.PSObject.Properties['PrimarySmtpAddress']
                $mailboxIdentityProperty = $definition.PSObject.Properties['MailboxIdentity']
                $candidateIdentity = if ($primaryAddressProperty -and $primaryAddressProperty.Value) {
                    [string]$primaryAddressProperty.Value
                }
                elseif ($mailboxIdentityProperty -and $mailboxIdentityProperty.Value) {
                    [string]$mailboxIdentityProperty.Value
                }
                else {
                    $null
                }
                if ($candidateIdentity -and -not $identityCandidates.Contains($candidateIdentity)) {
                    $identityCandidates.Add($candidateIdentity)
                }
            }
        }
    }

    end {
        $targetIdentity = $Identity
        if (-not $targetIdentity -and $identityCandidates.Count -gt 0) {
            $uniqueIdentityCandidates = @($identityCandidates | Select-Object -Unique)
            if ($uniqueIdentityCandidates.Count -gt 1) {
                throw "Input contains multiple mailbox identities: $($uniqueIdentityCandidates -join ', '). Search one mailbox at a time or specify Identity."
            }
            $targetIdentity = $uniqueIdentityCandidates[0]
        }
        if (-not $targetIdentity) {
            Write-Warning 'Identity is missing. Supply Identity or pipe Get-MailboxExtendedProperty results that include MailboxIdentity or PrimarySmtpAddress.'
            return
        }

        if ($definitions.Count -eq 0) {
            $definitions.AddRange(@(Get-MailboxExtendedProperty -Identity $targetIdentity))
        }
        if ($definitions.Count -eq 0) {
            Write-Warning 'No named-property definitions were provided or returned by Get-MailboxExtendedProperty.'
            return
        }

        $exactDefinitions = New-Object 'System.Collections.Generic.List[object]'
        $wildcardDefinitions = New-Object 'System.Collections.Generic.List[object]'
        foreach ($definition in $definitions) {
            if ($definition -is [string]) {
                if ([System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($definition)) {
                    $wildcardDefinitions.Add([pscustomobject]@{ Pattern = [string]$definition; PropertySetId = $null; PropertyType = $null; EwsMapiPropertyType = $null })
                }
                else {
                    Write-Warning "Property name '$definition' is not a definition object. Pipe Get-MailboxExtendedProperty output or provide a typed definition."
                }
                continue
            }

            $propertyNameProperty = $definition.PSObject.Properties['PropertyName']
            $propertyName = if ($propertyNameProperty) { [string]$propertyNameProperty.Value } else { $null }
            if ($propertyName -and [System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($propertyName)) {
                $setIdProperty = $definition.PSObject.Properties['PropertySetId']
                if (-not $setIdProperty) { $setIdProperty = $definition.PSObject.Properties['PropertyNamespace'] }
                $typeProperty = $definition.PSObject.Properties['PropertyType']
                $ewsTypeProperty = $definition.PSObject.Properties['EwsMapiPropertyType']
                $wildcardDefinitions.Add([pscustomobject]@{
                        Pattern             = $propertyName
                        PropertySetId       = if ($setIdProperty) { [string]$setIdProperty.Value } else { $null }
                        PropertyType        = if ($typeProperty) { [string]$typeProperty.Value } else { $null }
                        EwsMapiPropertyType = if ($ewsTypeProperty) { [string]$ewsTypeProperty.Value } else { $null }
                    })
            }
            else {
                $exactDefinitions.Add($definition)
            }
        }

        if ($wildcardDefinitions.Count -gt 0) {
            $storeDefinitions = @(Get-MailboxExtendedProperty -Identity $targetIdentity)
            foreach ($wildcardDefinition in $wildcardDefinitions) {
                $wildcardCandidates = @($storeDefinitions | Where-Object {
                        $_.PropertyName -and $_.PropertyName -like $wildcardDefinition.Pattern -and
                        (-not $wildcardDefinition.PropertySetId -or $_.PropertySetId -eq $wildcardDefinition.PropertySetId)
                    })
                foreach ($candidateDefinition in $wildcardCandidates) {
                    $expandedDefinition = [ordered]@{}
                    foreach ($property in $candidateDefinition.PSObject.Properties) {
                        $expandedDefinition[$property.Name] = $property.Value
                    }
                    if ($wildcardDefinition.PropertyType) {
                        $expandedDefinition['PropertyType'] = $wildcardDefinition.PropertyType
                    }
                    if ($wildcardDefinition.EwsMapiPropertyType) {
                        $expandedDefinition['EwsMapiPropertyType'] = $wildcardDefinition.EwsMapiPropertyType
                    }
                    $expandedObject = [pscustomobject]$expandedDefinition
                    if ($wildcardDefinition.PropertyType) {
                        $exactDefinitions.Add((ConvertTo-NormalizedNamedProperty -Definition $expandedObject -DiscoverySource 'Store'))
                    }
                    else {
                        if ($candidateDefinition.PSObject.TypeNames -contains 'ExchangeNamedProperties.NamedPropertyDefinition') {
                            $expandedObject.PSObject.TypeNames.Insert(0, 'ExchangeNamedProperties.NamedPropertyDefinition')
                        }
                        $exactDefinitions.Add($expandedObject)
                    }
                }
                if ($wildcardCandidates.Count -eq 0) {
                    Write-Warning "Wildcard '$($wildcardDefinition.Pattern)' matched no Store property names for the target mailbox."
                }
            }
        }

        $typedDefinitions = New-Object 'System.Collections.Generic.List[object]'
        $untypedDefinitions = New-Object 'System.Collections.Generic.List[object]'
        foreach ($definition in $exactDefinitions) {
            $normalized = if ($definition.PSObject.TypeNames -contains 'ExchangeNamedProperties.NamedPropertyDefinition') {
                $definition
            }
            else {
                ConvertTo-NormalizedNamedProperty -Definition $definition
            }
            $propertyTypeProperty = $normalized.PSObject.Properties['PropertyType']
            if (-not $propertyTypeProperty -or [string]::IsNullOrWhiteSpace([string]$propertyTypeProperty.Value)) {
                $untypedDefinitions.Add($normalized)
                continue
            }
            $typedDefinitions.Add([pscustomobject]@{ Normalized = $normalized; EwsDefinition = $null })
        }

        if ($untypedDefinitions.Count -gt 0) {
            Write-Warning "Skipped $($untypedDefinitions.Count) named-property definition(s) without PropertyType. Store inventory does not identify value types; supply Get-MailboxExtendedProperty results enriched with a verified PropertyType before EWS searching."
        }
        if ($typedDefinitions.Count -eq 0) {
            foreach ($untypedDefinition in $untypedDefinitions) {
                [pscustomobject]@{
                    PSTypeName         = 'ExchangeNamedProperties.SearchError'
                    MailboxIdentity    = $targetIdentity
                    PrimarySmtpAddress = $targetIdentity
                    PropertySetId      = $untypedDefinition.PropertySetId
                    PropertyId         = if ($untypedDefinition.PSObject.Properties['PropertyId']) { $untypedDefinition.PropertyId } elseif ($untypedDefinition.PSObject.Properties['PropDispId']) { $untypedDefinition.PropDispId } else { $null }
                    PropertyName       = $untypedDefinition.PropertyName
                    PropertyType       = if ($untypedDefinition.PSObject.Properties['PropertyType']) { $untypedDefinition.PropertyType } else { $null }
                    FolderPath         = $null
                    ErrorCode          = 'PropertyTypeRequired'
                    ErrorMessage       = 'This Store definition has no PropertyType. EWS cannot retrieve its item values without a verified type.'
                    SearchTimestamp    = $searchTimestamp
                }
            }
            return
        }

        $connectionArguments = @{ Identity = $targetIdentity }
        foreach ($connectionParameter in @('EwsUrl', 'EwsManagedApiPath', 'Credential', 'ExchangeVersion', 'RequestTimeoutSeconds', 'TraceFilePath')) {
            if ($PSBoundParameters.ContainsKey($connectionParameter)) {
                $connectionArguments[$connectionParameter] = $PSBoundParameters[$connectionParameter]
            }
        }
        foreach ($switchParameter in @('UseDefaultCredentials', 'TraceEnabled', 'Impersonate', 'TrustAll')) {
            if ($PSBoundParameters.ContainsKey($switchParameter)) {
                $connectionArguments[$switchParameter] = $PSBoundParameters[$switchParameter]
            }
        }

        $mailbox = Resolve-MailboxIdentity -Identity $targetIdentity
        foreach ($untypedDefinition in $untypedDefinitions) {
            [pscustomobject]@{
                PSTypeName         = 'ExchangeNamedProperties.SearchError'
                MailboxIdentity    = $mailbox.Identity
                PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                PropertySetId      = $untypedDefinition.PropertySetId
                PropertyId         = if ($untypedDefinition.PSObject.Properties['PropertyId']) { $untypedDefinition.PropertyId } elseif ($untypedDefinition.PSObject.Properties['PropDispId']) { $untypedDefinition.PropDispId } else { $null }
                PropertyName       = $untypedDefinition.PropertyName
                PropertyType       = if ($untypedDefinition.PSObject.Properties['PropertyType']) { $untypedDefinition.PropertyType } else { $null }
                FolderPath         = $null
                ErrorCode          = 'PropertyTypeRequired'
                ErrorMessage       = 'This Store definition has no PropertyType. EWS cannot retrieve its item values without a verified type.'
                SearchTimestamp    = $searchTimestamp
            }
        }
        $service = Get-EwsService @connectionArguments

        foreach ($definitionEntry in $typedDefinitions) {
            $definitionEntry.EwsDefinition = ConvertTo-EwsExtendedPropertyDefinition -Definition $definitionEntry.Normalized
        }

        try {
            $folders = @(Get-EwsFolderHierarchy -Service $service -PrimarySmtpAddress $mailbox.PrimarySmtpAddress -PageSize ([Math]::Min($PageSize, 100)) -IncludeSubfolders:$IncludeSubfolders -IncludeArchive:$Archive -SearchDumpster:$SearchDumpster -RetryCount $RetryCount)
        }
        catch {
            [pscustomobject]@{
                PSTypeName         = 'ExchangeNamedProperties.SearchError'
                MailboxIdentity    = $mailbox.Identity
                PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                FolderPath         = $null
                ErrorCode          = $_.Exception.GetType().Name
                ErrorMessage       = $_.Exception.Message
                SearchTimestamp    = $searchTimestamp
            }
            return
        }

        $totalFolders = $folders.Count
        $folderIndex = 0
        foreach ($folderEntry in $folders) {
            $folderIndex++
            if ($folderEntry.PSTypeNames -contains 'ExchangeNamedProperties.SearchError') {
                Add-Member -InputObject $folderEntry -MemberType NoteProperty -Name MailboxIdentity -Value $mailbox.Identity -Force
                Add-Member -InputObject $folderEntry -MemberType NoteProperty -Name PrimarySmtpAddress -Value $mailbox.PrimarySmtpAddress -Force
                Add-Member -InputObject $folderEntry -MemberType NoteProperty -Name SearchTimestamp -Value $searchTimestamp -Force
                $folderEntry
                continue
            }
            if ($Folder -and $folderEntry.FolderPath -notlike $Folder -and $folderEntry.Folder.DisplayName -ne $Folder) {
                continue
            }

            Write-Progress -Activity 'Searching mailbox extended properties' -Status $folderEntry.FolderPath -PercentComplete $(if ($totalFolders -gt 0) { [int](100 * $folderIndex / $totalFolders) } else { 100 })
            $propertySet = New-Object -TypeName Microsoft.Exchange.WebServices.Data.PropertySet -ArgumentList ([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
            foreach ($itemProperty in @(
                    [Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject,
                    [Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass,
                    [Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeCreated,
                    [Microsoft.Exchange.WebServices.Data.ItemSchema]::LastModifiedTime,
                    [Microsoft.Exchange.WebServices.Data.ItemSchema]::ParentFolderId,
                    [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::DateTimeReceived,
                    [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::InternetMessageId
                )) {
                $propertySet.Add($itemProperty)
            }
            foreach ($definitionEntry in $typedDefinitions) {
                $propertySet.Add($definitionEntry.EwsDefinition)
            }

            $traversals = @([Microsoft.Exchange.WebServices.Data.ItemTraversal]::Shallow)
            if ($IncludeAssociatedItems) {
                $traversals += [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Associated
            }

            foreach ($traversal in $traversals) {
                $offset = 0
                $mailSearchFilter = New-Object -TypeName 'Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring' -ArgumentList @(
                    [Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass,
                    'IPM.Note',
                    [Microsoft.Exchange.WebServices.Data.ContainmentMode]::Prefixed,
                    [Microsoft.Exchange.WebServices.Data.ComparisonMode]::IgnoreCase
                )
                do {
                    if ($ResultSize -and $matchCount -ge $ResultSize) {
                        break
                    }

                    $itemView = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList @($PageSize, $offset)
                    $itemView.PropertySet = $propertySet
                    $itemView.Traversal = $traversal
                    $itemView.OrderBy.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeCreated, [Microsoft.Exchange.WebServices.Data.SortDirection]::Ascending)

                    try {
                        $itemPage = Invoke-EwsRequestWithRetry -RetryCount $RetryCount -ScriptBlock {
                            Write-Output -NoEnumerate ($service.FindItems($folderEntry.FolderId, $mailSearchFilter, $itemView))
                        }
                    }
                    catch {
                        [pscustomobject]@{
                            PSTypeName         = 'ExchangeNamedProperties.SearchError'
                            MailboxIdentity    = $mailbox.Identity
                            PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                            FolderPath         = $folderEntry.FolderPath
                            ErrorCode          = $_.Exception.GetType().Name
                            ErrorMessage       = $_.Exception.Message
                            SearchTimestamp    = $searchTimestamp
                        }
                        break
                    }

                    foreach ($item in $itemPage.Items) {
                        if ($ResultSize -and $matchCount -ge $ResultSize) {
                            break
                        }
                        foreach ($definitionEntry in $typedDefinitions) {
                            if ($ResultSize -and $matchCount -ge $ResultSize) {
                                break
                            }
                            $propertyValue = $null
                            if (-not $item.TryGetProperty($definitionEntry.EwsDefinition, [ref]$propertyValue)) {
                                continue
                            }

                            $matchKey = '{0}|{1}|{2}' -f $item.Id.UniqueId, $definitionEntry.Normalized.Identity, $traversal
                            if (-not $seenMatches.Add($matchKey)) {
                                continue
                            }

                            $entryId = $null
                            $propertyTagProperty = $definitionEntry.Normalized.PSObject.Properties['PropertyTag']
                            if ($propertyTagProperty -and $null -ne $propertyTagProperty.Value -and [uint32]$propertyTagProperty.Value -eq [uint32]0x0FFF0102) {
                                $entryId = ConvertFrom-EwsExtendedProperty -Value $propertyValue
                            }

                            [pscustomobject]@{
                                PSTypeName         = 'ExchangeNamedProperties.SearchResult'
                                MailboxIdentity    = $mailbox.Identity
                                PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                                ItemId             = $item.Id.UniqueId
                                ChangeKey          = $item.Id.ChangeKey
                                EntryId            = $entryId
                                ParentFolderId     = if ($item.ParentFolderId) { $item.ParentFolderId.UniqueId } else { $null }
                                FolderPath         = $folderEntry.FolderPath
                                ItemClass          = $item.ItemClass
                                Subject            = $item.Subject
                                InternetMessageId  = $item.InternetMessageId
                                DateTimeReceived   = if ($item.DateTimeReceived) { $item.DateTimeReceived.ToUniversalTime().ToString('o', [Globalization.CultureInfo]::InvariantCulture) } else { $null }
                                DateTimeCreated    = if ($item.DateTimeCreated) { $item.DateTimeCreated.ToUniversalTime().ToString('o', [Globalization.CultureInfo]::InvariantCulture) } else { $null }
                                LastModifiedTime   = if ($item.LastModifiedTime) { $item.LastModifiedTime.ToUniversalTime().ToString('o', [Globalization.CultureInfo]::InvariantCulture) } else { $null }
                                PropertySetId      = $definitionEntry.Normalized.PropertySetId
                                PropertyId         = if ($definitionEntry.Normalized.PSObject.Properties['PropertyId']) { $definitionEntry.Normalized.PropertyId } elseif ($definitionEntry.Normalized.PSObject.Properties['PropDispId']) { $definitionEntry.Normalized.PropDispId } else { $null }
                                PropertyName       = $definitionEntry.Normalized.PropertyName
                                PropertyType       = $definitionEntry.Normalized.PropertyType
                                PropertyValue      = ConvertFrom-EwsExtendedProperty -Value $propertyValue
                                HasProperty        = $true
                                SearchTimestamp    = $searchTimestamp
                            }
                            $matchCount++
                            if ($ResultSize -and $matchCount -ge $ResultSize) {
                                break
                            }
                        }
                    }

                    $offset += $itemPage.Items.Count
                    $moreItems = $itemPage.MoreAvailable
                } while ($moreItems -and $itemPage.Items.Count -gt 0 -and (-not $ResultSize -or $matchCount -lt $ResultSize))
            }
        }
        Write-Progress -Activity 'Searching mailbox extended properties' -Completed
    }
}