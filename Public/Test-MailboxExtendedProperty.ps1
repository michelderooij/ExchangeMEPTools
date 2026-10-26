function Test-MailboxExtendedProperty {
    <#
    .SYNOPSIS
    Reports named-property inventory and risk indicators for a mailbox.

    .DESCRIPTION
    Analyzes Store mappings or supplied provider results. Use -ExchangeOnlineQuotas to compare an on-premises mailbox with Exchange Online quotas as a migration-target assessment. Incomplete inventories cannot produce a definitive mailbox-wide Pass.

    .PARAMETER Identity
    Mailbox SMTP address, alias, GUID, or another identifier resolvable by Get-Mailbox. When omitted, it is inferred from piped MailboxIdentity or PrimarySmtpAddress values; piped input must represent one mailbox.

    .PARAMETER MailboxExtendedProperty
    Optional provider results or named-property definitions to analyze.

    .PARAMETER Threshold
    Quota threshold proportion from 0 through 1. Defaults to 1.0 (100%). PercentageUsed results are reported from 0 through 100.

    .PARAMETER SelectFirst
    Number of top namespace and name-prefix groups to include in the report.

    .PARAMETER ExchangeOnlineQuotas
    Compare the inventory with Exchange Online quota limits to assess whether it fits the migration target. This skips the current Exchange Server quota lookup.

    .PARAMETER PassThru
    Return structured risk result objects and the normalized input definitions after the console report.

    .EXAMPLE
    Test-MailboxExtendedProperty -Identity user@contoso.com

    .EXAMPLE
    $properties | Test-MailboxExtendedProperty -Identity user@contoso.com -Threshold 0.9

    .EXAMPLE
    $properties | Test-MailboxExtendedProperty -ExchangeOnlineQuotas
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [Alias('MailboxIdentity', 'PrimarySmtpAddress')]
        [string]$Identity,

        [Parameter(ValueFromPipeline = $true)]
        [Alias('PropertyDefinition')]
        [psobject]$MailboxExtendedProperty,

        [ValidateRange(0.0, 1.0)]
        [double]$Threshold = 1.0,

        [ValidateRange(0, 1000)]
        [int]$SelectFirst = 10,

        [switch]$ExchangeOnlineQuotas,

        [switch]$PassThru
    )

    begin {
        $inputDefinitions = New-Object 'System.Collections.Generic.List[object]'
        $inputMailboxIdentities = New-Object 'System.Collections.Generic.List[string]'
        $prefixLength = 5
    }

    process {
        if ($null -ne $MailboxExtendedProperty) {
            $inputDefinitions.Add($MailboxExtendedProperty)
            foreach ($identityPropertyName in @('MailboxIdentity', 'PrimarySmtpAddress')) {
                $identityProperty = $MailboxExtendedProperty.PSObject.Properties[$identityPropertyName]
                if ($identityProperty -and -not [string]::IsNullOrWhiteSpace([string]$identityProperty.Value)) {
                    $inputMailboxIdentities.Add([string]$identityProperty.Value)
                    break
                }
            }
        }
    }

    end {
        if ([string]::IsNullOrWhiteSpace($Identity)) {
            $uniqueMailboxIdentities = @($inputMailboxIdentities | Select-Object -Unique)
            if ($uniqueMailboxIdentities.Count -eq 1) {
                $Identity = $uniqueMailboxIdentities[0]
            }
            elseif ($uniqueMailboxIdentities.Count -gt 1) {
                throw 'Pipeline input contains multiple mailbox identities. Specify -Identity to analyze one mailbox at a time.'
            }
            else {
                throw 'Specify -Identity or pipe input that includes MailboxIdentity or PrimarySmtpAddress.'
            }
        }

        if ($inputDefinitions.Count -gt 0) {
            $definitions = @($inputDefinitions.ToArray())
        }
        else {
            $definitions = @(Get-MailboxExtendedProperty -Identity $Identity -WarningAction SilentlyContinue)
        }

        $limits = Get-NamedPropertyLimits
        $serverLimitByTest = @{}
        foreach ($limitEntry in $limits.ExchangeServerLimits) {
            if (-not $ExchangeOnlineQuotas -and $limitEntry.Name) {
                $serverLimitByTest[[string]$limitEntry.Name] = $limitEntry
            }
        }
        $referenceLimitByTest = @{}
        foreach ($limitEntry in $limits.ExchangeOnlineReferenceLimits) {
            if ($limitEntry.Name) {
                $referenceLimitByTest[[string]$limitEntry.Name] = $limitEntry
            }
        }
        $serverLimitSource = $limits.ExchangeServerSource
        if (-not $ExchangeOnlineQuotas -and (Get-Command -Name 'Get-MailboxStatistics' -ErrorAction SilentlyContinue)) {
            try {
                $mailboxStatistics = Get-MailboxStatistics -Identity $Identity -ErrorAction Stop
                $quotaProperty = $mailboxStatistics.PSObject.Properties['NamedPropertiesCountQuota']
                $runtimeQuota = 0.0
                if ($quotaProperty -and [double]::TryParse([string]$quotaProperty.Value, [ref]$runtimeQuota) -and $runtimeQuota -gt 0) {
                    $serverLimitSource = 'Get-MailboxStatistics.NamedPropertiesCountQuota'
                    $derivedServerLimits = @(
                        [pscustomobject]@{ Name = 'TotalNamedProperties'; Scope = 'All namespaces in mailbox'; Limit = [int]$runtimeQuota; PercentOfTotalQuota = 100; Source = $serverLimitSource }
                        [pscustomobject]@{ Name = 'InternetHeadersNamespaceUsage'; Scope = 'InternetHeaders namespace'; Limit = [int](0.6 * $runtimeQuota); PercentOfTotalQuota = 60; Source = $serverLimitSource }
                        [pscustomobject]@{ Name = 'PublicStringsNamespaceUsage'; Scope = 'PublicStrings namespace'; Limit = [int](0.2 * $runtimeQuota); PercentOfTotalQuota = 20; Source = $serverLimitSource }
                        [pscustomobject]@{ Name = 'NamespaceUsage'; Scope = 'Each distinct namespace GUID'; Limit = [int](0.2 * $runtimeQuota); PercentOfTotalQuota = 20; Source = $serverLimitSource }
                        [pscustomobject]@{ Name = 'RepeatedPropertyNamePrefix'; Scope = 'Each first-five-character property-name prefix'; Limit = [int](0.1 * $runtimeQuota); PercentOfTotalQuota = $null; Source = $serverLimitSource }
                    )
                    foreach ($derivedLimit in $derivedServerLimits) {
                        $serverLimitByTest[$derivedLimit.Name] = $derivedLimit
                    }
                }
            }
            catch {
                Write-Warning "Could not read NamedPropertiesCountQuota from Get-MailboxStatistics for '$Identity': $($_.Exception.Message). Quota status will use configured/reference limits."
            }
        }

        $sourceValues = @($definitions | ForEach-Object {
                $sourceProperty = $_.PSObject.Properties['DiscoverySource']
                $isStoreDefinition = ($_.PSObject.TypeNames -contains 'ExchangeNamedProperties.NamedPropertyDefinition') -and
                $_.PSObject.Properties['NamedPropCount'] -and $_.PSObject.Properties['PropNumber']
                if ($sourceProperty -and $sourceProperty.Value) { [string]$sourceProperty.Value }
                elseif ($isStoreDefinition) { 'Store' }
                else { 'Pipeline' }
            } | Select-Object -Unique)
        $completenessValues = @($definitions | ForEach-Object {
                $completenessProperty = $_.PSObject.Properties['DiscoveryCompleteness']
                $isStoreDefinition = ($_.PSObject.TypeNames -contains 'ExchangeNamedProperties.NamedPropertyDefinition') -and
                $_.PSObject.Properties['NamedPropCount'] -and $_.PSObject.Properties['PropNumber']
                if ($completenessProperty -and $completenessProperty.Value) { [string]$completenessProperty.Value }
                elseif ($isStoreDefinition) { 'Authoritative' }
                else { 'SuppliedDefinitions' }
            } | Select-Object -Unique)
        $isCompleteInventory = $completenessValues.Count -eq 1 -and $completenessValues[0] -eq 'Authoritative'
        $dataSource = $sourceValues -join ', '
        $completeness = $completenessValues -join ', '
        $mailbox = Resolve-MailboxIdentity -Identity $Identity
        $results = New-Object 'System.Collections.Generic.List[object]'

        $normalizedDefinitions = New-Object 'System.Collections.Generic.List[object]'
        $invalidCount = 0
        foreach ($definition in $definitions) {
            if ($definition.PSObject.TypeNames -contains 'ExchangeNamedProperties.NamedPropertyDefinition') {
                $normalizedDefinitions.Add($definition)
                $propertyId = if ($definition.PSObject.Properties['PropertyId']) { $definition.PropertyId } elseif ($definition.PSObject.Properties['PropDispId']) { $definition.PropDispId } else { $null }
                if (-not $definition.PropertySetId -or (-not $definition.PropertyName -and $null -eq $propertyId)) {
                    $invalidCount++
                }
            }
            else {
                $validation = Test-NamedPropertyDefinition -Definition $definition
                if ($validation.IsValid) {
                    $normalizedDefinitions.Add($validation.Definition)
                }
                else {
                    $invalidCount++
                }
            }
        }

        $namedPropertyCount = $normalizedDefinitions.Count
        $storeCountValues = @($definitions | Where-Object {
                $sourceProperty = $_.PSObject.Properties['DiscoverySource']
                $countProperty = $_.PSObject.Properties['NamedPropCount']
                $isStoreDefinition = ($_.PSObject.TypeNames -contains 'ExchangeNamedProperties.NamedPropertyDefinition') -and
                $_.PSObject.Properties['PropNumber']
                (($sourceProperty -and $sourceProperty.Value -eq 'Store') -or $isStoreDefinition) -and
                $countProperty -and $null -ne $countProperty.Value
            } | ForEach-Object { [int]$_.NamedPropCount } | Select-Object -Unique)
        if ($storeCountValues.Count -eq 1) {
            $namedPropertyCount = $storeCountValues[0]
        }
        elseif ($storeCountValues.Count -gt 1) {
            Write-Warning 'Store definitions contain inconsistent NamedPropCount values. Using the count of valid returned definitions for the total.'
        }

        $keys = @($normalizedDefinitions | ForEach-Object {
                if ($_.PropertyName) { '{0}|Name:{1}' -f $_.PropertySetId, $_.PropertyName.ToLowerInvariant() }
                elseif ($_.PSObject.Properties['PropertyId'] -and $null -ne $_.PropertyId) { '{0}|Id:{1}' -f $_.PropertySetId, $_.PropertyId }
                elseif ($_.PSObject.Properties['PropDispId'] -and $null -ne $_.PropDispId) { '{0}|Id:{1}' -f $_.PropertySetId, $_.PropDispId }
                else { 'Tag:{0}' -f $_.PropertyTag }
            })
        $duplicateCount = @($keys | Group-Object | Where-Object Count -GT 1 | ForEach-Object { $_.Count - 1 } | Measure-Object -Sum).Sum
        if ($null -eq $duplicateCount) { $duplicateCount = 0 }

        $typedDefinitions = @($normalizedDefinitions | Where-Object {
                $typeProperty = $_.PSObject.Properties['PropertyType']
                $typeProperty -and -not [string]::IsNullOrWhiteSpace([string]$typeProperty.Value)
            })
        $typeConflicts = @($typedDefinitions | Group-Object {
                if ($_.PropertyName) { '{0}|Name:{1}' -f $_.PropertySetId, $_.PropertyName.ToLowerInvariant() }
                elseif ($_.PSObject.Properties['PropertyId'] -and $null -ne $_.PropertyId) { '{0}|Id:{1}' -f $_.PropertySetId, $_.PropertyId }
                else { '{0}|Id:{1}' -f $_.PropertySetId, $_.PropDispId }
            } | Where-Object { @($_.Group.PropertyType | Select-Object -Unique).Count -gt 1 }).Count

        $publicStringsGuid = '00020329-0000-0000-c000-000000000046'
        $internetHeadersGuid = '00020386-0000-0000-c000-000000000046'
        $publicStringsCount = @($normalizedDefinitions | Where-Object { $_.PropertySetId -eq $publicStringsGuid }).Count
        $internetHeadersCount = @($normalizedDefinitions | Where-Object { $_.PropertySetId -eq $internetHeadersGuid }).Count
        $namespaceCounts = @($normalizedDefinitions | Where-Object { $_.PropertySetId } | Group-Object PropertySetId | Sort-Object Count -Descending | Select-Object -First $SelectFirst)

        $prefixCandidates = New-Object 'System.Collections.Generic.List[object]'
        foreach ($definition in ($normalizedDefinitions | Where-Object PropertyName)) {
            $propertyName = [string]$definition.PropertyName
            $propertyPrefixLength = [Math]::Min($prefixLength, $propertyName.Length)
            $prefix = $propertyName.Substring(0, $propertyPrefixLength)
            $typeProperty = $definition.PSObject.Properties['PropertyType']
            $typeIsUnknown = -not $typeProperty -or [string]::IsNullOrWhiteSpace([string]$typeProperty.Value)
            if ($typeIsUnknown -or [string]$typeProperty.Value -in @('String', 'StringProperty')) {
                $prefixCandidates.Add([pscustomobject]@{ Prefix = $prefix; PropertyName = $propertyName; TypeIsUnknown = $typeIsUnknown })
            }
        }
        $prefixGroups = @($prefixCandidates | Group-Object Prefix | Sort-Object Count -Descending | Select-Object -First $SelectFirst)

        $totalServerLimit = $serverLimitByTest['TotalNamedProperties']
        $totalReferenceLimit = $referenceLimitByTest['TotalNamedProperties']
        $totalLimit = if ($ExchangeOnlineQuotas) { $totalReferenceLimit } elseif ($totalServerLimit) { $totalServerLimit } else { $totalReferenceLimit }
        $totalLimitIsApplicable = $null -ne $totalServerLimit -or ($ExchangeOnlineQuotas -and $null -ne $totalReferenceLimit)
        $totalStatus = 'Unknown'
        $totalLimitValue = $null
        $totalThresholdValue = $null
        $totalPercentage = $null
        $totalFractionUsed = $null
        $recommendation = 'No Exchange Server named-property limit is configured; obtain a version-specific documented or administrator-approved limit.'
        if ($totalLimit -and $null -ne $totalLimit.Limit -and [double]$totalLimit.Limit -gt 0) {
            $totalLimitValue = [double]$totalLimit.Limit
            $totalThresholdValue = [int]($totalLimitValue * $Threshold)
            $totalFractionUsed = [double]$namedPropertyCount / $totalLimitValue
            $totalPercentage = $totalFractionUsed * 100
            if ($totalLimitIsApplicable) {
                $totalStatus = if ($namedPropertyCount -ge $totalLimitValue) { 'Critical' } elseif ($namedPropertyCount -ge $totalThresholdValue) { 'Warning' } elseif ($isCompleteInventory) { 'Pass' } else { 'Incomplete' }
                $recommendation = if ($ExchangeOnlineQuotas -and $totalStatus -eq 'Critical') { 'The inventory exceeds the Exchange Online migration-target total quota.' } elseif ($ExchangeOnlineQuotas -and $totalStatus -eq 'Warning') { 'Review the inventory against the Exchange Online migration-target warning threshold.' } elseif ($ExchangeOnlineQuotas) { 'Compared against the Exchange Online migration-target total quota.' } elseif ($totalStatus -eq 'Critical') { 'Review and reduce named-property use; verify the configured limit applies to this Exchange build.' } elseif ($totalStatus -eq 'Warning') { 'Review the inventory against the configured Exchange Server warning threshold.' } elseif ($isCompleteInventory) { 'Continue monitoring with a complete Store inventory.' } else { 'The configured limit is not a substitute for a complete mailbox inventory.' }
            }
            else {
                $totalStatus = 'Unknown'
                $recommendation = 'Percentage is a comparison with the Exchange Online reference only; applicability to this Exchange Server version is unknown.'
            }
        }
        elseif (-not $isCompleteInventory) {
            $totalStatus = 'Incomplete'
            $recommendation = 'The inventory is not authoritative for the mailbox-wide mapping; use the Store provider before assessing quota usage.'
        }

        $results.Add([pscustomobject]@{
                PSTypeName            = 'ExchangeNamedProperties.TestResult'
                Identity              = $Identity
                PrimarySmtpAddress    = $mailbox.PrimarySmtpAddress
                Test                  = 'TotalNamedProperties'
                Scope                 = 'Mailbox named-property mapping'
                CurrentValue          = $namedPropertyCount
                NamedPropCount        = $namedPropertyCount
                Limit                 = $totalLimitValue
                PercentOfTotalQuota   = if ($totalLimit -and $totalLimit.PSObject.Properties['PercentOfTotalQuota']) { $totalLimit.PercentOfTotalQuota } else { $null }
                Threshold             = $Threshold
                ThresholdValue        = $totalThresholdValue
                PercentageUsed        = $totalPercentage
                Status                = $totalStatus
                Severity              = if ($totalStatus -eq 'Critical') { 'High' } elseif ($totalStatus -eq 'Warning') { 'Medium' } else { 'Informational' }
                IsCompleteInventory   = $isCompleteInventory
                DiscoveryCompleteness = $completeness
                Recommendation        = $recommendation
                DataSource            = if ($ExchangeOnlineQuotas) { "$dataSource; $($limits.ExchangeOnlineReferenceSource) (migration target)" } elseif ($totalServerLimit) { "$dataSource; $serverLimitSource" } elseif ($totalReferenceLimit) { "$dataSource; $($limits.ExchangeOnlineReferenceSource)" } else { $dataSource }
            })

        $analysisTests = New-Object 'System.Collections.Generic.List[object]'
        $analysisTests.Add([pscustomobject]@{ Name = 'PublicStringsNamespaceUsage'; Value = $publicStringsCount; Recommendation = 'Review properties in the Public Strings namespace and remove unused mappings only after item-level validation.' })
        $analysisTests.Add([pscustomobject]@{ Name = 'InternetHeadersNamespaceUsage'; Value = $internetHeadersCount; Recommendation = 'Review custom Internet Headers namespace use and avoid creating unnecessary header names.' })
        $analysisTests.Add([pscustomobject]@{ Name = 'InvalidDefinitions'; Value = $invalidCount; Recommendation = 'Correct or remove invalid property definitions before searching or removal.' })
        $analysisTests.Add([pscustomobject]@{ Name = 'DuplicateDefinitions'; Value = [int]$duplicateCount; Recommendation = 'Deduplicate definitions so each named property is evaluated once.' })
        $analysisTests.Add([pscustomobject]@{ Name = 'ConflictingPropertyTypes'; Value = $typeConflicts; Recommendation = 'Resolve conflicting property type definitions before issuing EWS requests.' })

        foreach ($testDefinition in $analysisTests) {
            $serverLimit = $serverLimitByTest[$testDefinition.Name]
            $referenceLimit = $referenceLimitByTest[$testDefinition.Name]
            $applicableLimit = if ($ExchangeOnlineQuotas) { $referenceLimit } elseif ($serverLimit) { $serverLimit } else { $referenceLimit }
            $limitIsApplicable = $null -ne $serverLimit -or ($ExchangeOnlineQuotas -and $null -ne $referenceLimit)
            $fractionUsed = if ($applicableLimit -and [double]$applicableLimit.Limit -gt 0) { [double]$testDefinition.Value / [double]$applicableLimit.Limit } else { $null }
            $percentageUsed = if ($null -ne $fractionUsed) { $fractionUsed * 100 } else { $null }
            $thresholdValue = if ($applicableLimit -and [double]$applicableLimit.Limit -gt 0) { [int]([double]$applicableLimit.Limit * $Threshold) } else { $null }
            if ($testDefinition.Name -in @('InvalidDefinitions', 'DuplicateDefinitions', 'ConflictingPropertyTypes')) {
                $status = if ($testDefinition.Value -gt 0) { 'Warning' } elseif ($isCompleteInventory) { 'Pass' } else { 'Incomplete' }
            }
            elseif ($limitIsApplicable) {
                $status = if ([double]$testDefinition.Value -ge [double]$applicableLimit.Limit) { 'Critical' } elseif ([double]$testDefinition.Value -ge $thresholdValue) { 'Warning' } elseif ($isCompleteInventory) { 'Pass' } else { 'Incomplete' }
            }
            elseif ($referenceLimit) {
                $status = 'Unknown'
            }
            else {
                $status = if ($isCompleteInventory) { 'Pass' } else { 'Incomplete' }
            }
            $results.Add([pscustomobject]@{
                    PSTypeName            = 'ExchangeNamedProperties.TestResult'
                    Identity              = $Identity
                    PrimarySmtpAddress    = $mailbox.PrimarySmtpAddress
                    Test                  = $testDefinition.Name
                    Scope                 = if ($applicableLimit -and $applicableLimit.Scope) { [string]$applicableLimit.Scope } else { 'Returned definitions' }
                    CurrentValue          = $testDefinition.Value
                    Limit                 = if ($applicableLimit) { [double]$applicableLimit.Limit } else { $null }
                    PercentOfTotalQuota   = if ($applicableLimit -and $applicableLimit.PSObject.Properties['PercentOfTotalQuota']) { $applicableLimit.PercentOfTotalQuota } else { $null }
                    Threshold             = if ($applicableLimit) { $Threshold } else { $null }
                    ThresholdValue        = $thresholdValue
                    PercentageUsed        = $percentageUsed
                    Status                = $status
                    Severity              = if ($status -eq 'Warning') { 'Medium' } elseif ($status -in @('Incomplete', 'Unknown')) { 'Informational' } else { 'Low' }
                    IsCompleteInventory   = $isCompleteInventory
                    DiscoveryCompleteness = $completeness
                    Recommendation        = if ($status -eq 'Warning') { $testDefinition.Recommendation } elseif ($status -eq 'Critical' -and $ExchangeOnlineQuotas) { 'This definition count exceeds the Exchange Online migration-target quota.' } elseif ($ExchangeOnlineQuotas -and $referenceLimit) { 'Compared against the Exchange Online migration-target quota.' } elseif ($referenceLimit) { 'Percentage is an Exchange Online reference comparison only; applicability to this Exchange Server version is unknown.' } elseif ($status -eq 'Incomplete') { 'No issue was identified within the returned input scope; the inventory is incomplete.' } else { 'No issue was identified within the reported input scope.' }
                    DataSource            = if ($ExchangeOnlineQuotas) { "$dataSource; $($limits.ExchangeOnlineReferenceSource) (migration target)" } elseif ($serverLimit) { "$dataSource; $serverLimitSource" } elseif ($referenceLimit) { "$dataSource; $($limits.ExchangeOnlineReferenceSource)" } else { $dataSource }
                })
        }

        foreach ($namespaceGroup in $namespaceCounts) {
            $serverLimit = $serverLimitByTest['NamespaceUsage']
            $referenceLimit = $referenceLimitByTest['NamespaceUsage']
            $applicableLimit = if ($ExchangeOnlineQuotas) { $referenceLimit } elseif ($serverLimit) { $serverLimit } else { $referenceLimit }
            $limitIsApplicable = $null -ne $serverLimit -or ($ExchangeOnlineQuotas -and $null -ne $referenceLimit)
            $fractionUsed = if ($applicableLimit -and [double]$applicableLimit.Limit -gt 0) { [double]$namespaceGroup.Count / [double]$applicableLimit.Limit } else { $null }
            $percentageUsed = if ($null -ne $fractionUsed) { $fractionUsed * 100 } else { $null }
            $thresholdValue = if ($applicableLimit -and [double]$applicableLimit.Limit -gt 0) { [int]([double]$applicableLimit.Limit * $Threshold) } else { $null }
            $results.Add([pscustomobject]@{
                    PSTypeName            = 'ExchangeNamedProperties.TestResult'
                    Identity              = $Identity
                    PrimarySmtpAddress    = $mailbox.PrimarySmtpAddress
                    Test                  = 'NamespaceUsage'
                    Scope                 = $namespaceGroup.Name
                    CurrentValue          = $namespaceGroup.Count
                    Limit                 = if ($applicableLimit) { [double]$applicableLimit.Limit } else { $null }
                    PercentOfTotalQuota   = if ($applicableLimit -and $applicableLimit.PSObject.Properties['PercentOfTotalQuota']) { $applicableLimit.PercentOfTotalQuota } else { $null }
                    Threshold             = if ($applicableLimit) { $Threshold } else { $null }
                    ThresholdValue        = $thresholdValue
                    PercentageUsed        = $percentageUsed
                    Status                = if ($limitIsApplicable -and $namespaceGroup.Count -ge [double]$applicableLimit.Limit) { 'Critical' } elseif ($limitIsApplicable -and $namespaceGroup.Count -ge $thresholdValue) { 'Warning' } elseif ($limitIsApplicable -and $isCompleteInventory) { 'Pass' } elseif ($limitIsApplicable) { 'Incomplete' } else { 'Unknown' }
                    Severity              = 'Informational'
                    IsCompleteInventory   = $isCompleteInventory
                    DiscoveryCompleteness = $completeness
                    Recommendation        = if ($ExchangeOnlineQuotas) { 'Compared against the Exchange Online migration-target per-namespace quota.' } elseif ($referenceLimit -and -not $serverLimit) { 'Percentage is an Exchange Online per-namespace reference comparison; applicability to this Exchange Server version is unknown.' } else { 'Review use of this property set in the context of the mailbox workload.' }
                    DataSource            = if ($ExchangeOnlineQuotas) { "$dataSource; $($limits.ExchangeOnlineReferenceSource) (migration target)" } elseif ($serverLimit) { "$dataSource; $serverLimitSource" } elseif ($referenceLimit) { "$dataSource; $($limits.ExchangeOnlineReferenceSource)" } else { $dataSource }
                })
        }

        foreach ($prefixGroup in $prefixGroups) {
            $serverLimit = $serverLimitByTest['RepeatedPropertyNamePrefix']
            $referenceLimit = $referenceLimitByTest['RepeatedPropertyNamePrefix']
            $applicableLimit = if ($ExchangeOnlineQuotas) { $referenceLimit } elseif ($serverLimit) { $serverLimit } else { $referenceLimit }
            $limitIsApplicable = $null -ne $serverLimit -or ($ExchangeOnlineQuotas -and $null -ne $referenceLimit)
            $prefixUnknownTypeCount = @($prefixGroup.Group | Where-Object TypeIsUnknown).Count
            $fractionUsed = if ($applicableLimit -and [double]$applicableLimit.Limit -gt 0) { [double]$prefixGroup.Count / [double]$applicableLimit.Limit } else { $null }
            $percentageUsed = if ($null -ne $fractionUsed) { $fractionUsed * 100 } else { $null }
            $thresholdValue = if ($applicableLimit -and [double]$applicableLimit.Limit -gt 0) { [int]([double]$applicableLimit.Limit * $Threshold) } else { $null }
            $results.Add([pscustomobject]@{
                    PSTypeName            = 'ExchangeNamedProperties.TestResult'
                    Identity              = $Identity
                    PrimarySmtpAddress    = $mailbox.PrimarySmtpAddress
                    Test                  = 'RepeatedPropertyNamePrefix'
                    Scope                 = $prefixGroup.Name
                    PropertySetId         = $null
                    PropertyNamePrefix    = $prefixGroup.Name
                    UnknownTypeCount      = $prefixUnknownTypeCount
                    CurrentValue          = $prefixGroup.Count
                    Limit                 = if ($applicableLimit) { [double]$applicableLimit.Limit } else { $null }
                    PercentOfTotalQuota   = if ($applicableLimit -and $applicableLimit.PSObject.Properties['PercentOfTotalQuota']) { $applicableLimit.PercentOfTotalQuota } else { $null }
                    Threshold             = if ($applicableLimit) { $Threshold } else { $null }
                    ThresholdValue        = $thresholdValue
                    PercentageUsed        = $percentageUsed
                    Status                = if ($prefixUnknownTypeCount -gt 0) { 'Unknown' } elseif ($limitIsApplicable -and $prefixGroup.Count -ge [double]$applicableLimit.Limit) { 'Critical' } elseif ($limitIsApplicable -and $prefixGroup.Count -ge $thresholdValue) { 'Warning' } elseif ($limitIsApplicable -and $isCompleteInventory) { 'Pass' } elseif ($limitIsApplicable) { 'Incomplete' } elseif ($referenceLimit) { 'Unknown' } elseif ($prefixGroup.Count -gt 1) { 'Warning' } elseif ($isCompleteInventory) { 'Pass' } else { 'Incomplete' }
                    Severity              = if ($prefixUnknownTypeCount -gt 0 -or (-not $limitIsApplicable -and $referenceLimit)) { 'Informational' } elseif ($fractionUsed -ge 1) { 'High' } elseif ($fractionUsed -ge $Threshold) { 'Medium' } elseif ($prefixGroup.Count -gt 1) { 'Low' } else { 'Informational' }
                    IsCompleteInventory   = $isCompleteInventory
                    DiscoveryCompleteness = $completeness
                    Recommendation        = if ($prefixUnknownTypeCount -gt 0) { 'Store does not provide property types. This is a candidate count of names with this first-five-character prefix; exact StringProperty quota usage is unknown.' } elseif ($ExchangeOnlineQuotas) { 'Compared against the Exchange Online migration-target prefix quota.' } elseif ($referenceLimit -and -not $serverLimit) { 'This is an Exchange Online reference comparison; applicability to this Exchange Server version is unknown.' } else { 'Review whether repeated name prefixes are expected and consistently governed.' }
                    DataSource            = if ($ExchangeOnlineQuotas) { "$dataSource; $($limits.ExchangeOnlineReferenceSource) (migration target)" } elseif ($serverLimit) { "$dataSource; $serverLimitSource" } elseif ($referenceLimit) { "$dataSource; $($limits.ExchangeOnlineReferenceSource)" } else { $dataSource }
                })
        }

        $definitionsWithObservation = @($definitions | Where-Object {
                $_.PSObject.Properties['ItemCount'] -and $null -ne $_.ItemCount
            })
        $observationStatus = if ($definitionsWithObservation.Count -gt 0) { 'Known' } else { 'Unknown' }
        $suppliedDefinitions = @($definitions | Where-Object {
                $sourceProperty = $_.PSObject.Properties['DiscoverySource']
                $isStoreDefinition = ($_.PSObject.TypeNames -contains 'ExchangeNamedProperties.NamedPropertyDefinition') -and
                $_.PSObject.Properties['NamedPropCount'] -and $_.PSObject.Properties['PropNumber']
                -not $isStoreDefinition -and
                (-not $sourceProperty -or $sourceProperty.Value -match '(?i)supplieddefinitions|pipeline|input')
            })
        $suppliedNotObservedCount = if ($observationStatus -eq 'Known') {
            @($suppliedDefinitions | Where-Object { $null -ne $_.ItemCount -and [int]$_.ItemCount -eq 0 }).Count
        }
        else {
            $null
        }
        $suppliedObservationStatus = if ($observationStatus -eq 'Unknown') {
            'Unknown'
        }
        elseif ($suppliedNotObservedCount -gt 0) {
            'Warning'
        }
        elseif ($isCompleteInventory) {
            'Pass'
        }
        else {
            'Incomplete'
        }
        $suppliedObservationRecommendation = if ($observationStatus -eq 'Unknown') {
            'No scan observation data was supplied; an input definition cannot be classified as observed or not observed.'
        }
        elseif ($suppliedNotObservedCount -gt 0) {
            'These input definitions were not observed in the reported scan scope; this does not prove that the property is unused elsewhere in the mailbox.'
        }
        else {
            'Interpret observation state only within the reported scan scope.'
        }
        $results.Add([pscustomobject]@{
                PSTypeName            = 'ExchangeNamedProperties.TestResult'
                Identity              = $Identity
                PrimarySmtpAddress    = $mailbox.PrimarySmtpAddress
                Test                  = 'SuppliedPropertiesNotObserved'
                Scope                 = 'Supplied definitions compared with EWS scan observations'
                CurrentValue          = $suppliedNotObservedCount
                Limit                 = $null
                PercentOfTotalQuota   = $null
                Threshold             = $null
                ThresholdValue        = $null
                PercentageUsed        = $null
                Status                = $suppliedObservationStatus
                Severity              = if ($suppliedObservationStatus -eq 'Warning') { 'Low' } else { 'Informational' }
                IsCompleteInventory   = $false
                DiscoveryCompleteness = $completeness
                Recommendation        = $suppliedObservationRecommendation
                DataSource            = if ($observationStatus -eq 'Known') { $dataSource } else { 'No item scan' }
            })
        $results.Add([pscustomobject]@{
                PSTypeName            = 'ExchangeNamedProperties.TestResult'
                Identity              = $Identity
                PrimarySmtpAddress    = $mailbox.PrimarySmtpAddress
                Test                  = 'ObservedProperties'
                Scope                 = 'EWS item scan'
                CurrentValue          = if ($observationStatus -eq 'Known') { @($definitionsWithObservation | Where-Object { $_.ItemCount -gt 0 }).Count } else { $null }
                Limit                 = $null
                PercentOfTotalQuota   = $null
                Threshold             = $null
                ThresholdValue        = $null
                PercentageUsed        = $null
                Status                = if ($observationStatus -eq 'Known') { 'Pass' } else { 'Unknown' }
                Severity              = 'Informational'
                IsCompleteInventory   = $false
                DiscoveryCompleteness = $completeness
                Recommendation        = if ($observationStatus -eq 'Known') { 'Interpret observations only within the scan scope reported by the source provider.' } else { 'No EWS item observation data was supplied; run an explicitly scoped item search before interpreting item usage.' }
                DataSource            = if ($observationStatus -eq 'Known') { $dataSource } else { 'No item scan' }
            })

        $totalResult = @($results | Where-Object Test -EQ 'TotalNamedProperties')[0]
        $publicStringsResult = @($results | Where-Object Test -EQ 'PublicStringsNamespaceUsage')[0]
        $internetHeadersResult = @($results | Where-Object Test -EQ 'InternetHeadersNamespaceUsage')[0]
        $namespaceResults = @($results | Where-Object Test -EQ 'NamespaceUsage')
        $prefixResults = @($results | Where-Object Test -EQ 'RepeatedPropertyNamePrefix')
        $namespaceLimit = if ($serverLimitByTest['NamespaceUsage']) { $serverLimitByTest['NamespaceUsage'].Limit } else { $referenceLimitByTest['NamespaceUsage'].Limit }
        $prefixLimit = if ($serverLimitByTest['RepeatedPropertyNamePrefix']) { $serverLimitByTest['RepeatedPropertyNamePrefix'].Limit } else { $referenceLimitByTest['RepeatedPropertyNamePrefix'].Limit }
        $hasKnownExceededQuota = $false
        $hasReferenceExceededQuota = $false
        $hasCandidateThresholdExceeded = $false
        $hasUnknownQuota = $false

        if ($ExchangeOnlineQuotas) {
            Write-Host 'Used Exchange Online quotas as a migration-target assessment.' -ForegroundColor Cyan
        }
        elseif (-not $totalServerLimit -and $totalReferenceLimit) {
            Write-Host 'Quota values are Exchange Online references; applicability to this Exchange Server version is unknown.' -ForegroundColor Yellow
        }

        Write-Host "Checking the mailbox $Identity for having exceeded the threshold of $([int]($Threshold * 100))% of any named properties quota." -ForegroundColor Blue
        Write-Host "The total named properties quota is $($totalResult.Limit)." -ForegroundColor Gray
        Write-Host "The PublicStrings namespace named properties quota is $($publicStringsResult.Limit)." -ForegroundColor Gray
        Write-Host "The InternetHeaders namespace named properties quota is $($internetHeadersResult.Limit)." -ForegroundColor Gray
        Write-Host "The any namespace named properties quota is $namespaceLimit." -ForegroundColor Gray
        Write-Host "The same name prefix named properties quota is $prefixLimit." -ForegroundColor Gray

        $scenarioRows = @(
            [pscustomobject]@{ Heading = 'total named properties quota'; Result = $totalResult; Count = $namedPropertyCount }
            [pscustomobject]@{ Heading = 'PublicStrings namespace named properties quota'; Result = $publicStringsResult; Count = $publicStringsCount }
            [pscustomobject]@{ Heading = 'InternetHeaders namespace named properties quota'; Result = $internetHeadersResult; Count = $internetHeadersCount }
        )
        foreach ($scenario in $scenarioRows) {
            Write-Host "Checking if mailbox has exceeded the threshold of $($scenario.Heading)." -ForegroundColor Blue
            if ($null -eq $scenario.Result.Limit -or $null -eq $scenario.Result.ThresholdValue) {
                Write-Host "The $($scenario.Heading) is unknown for this Exchange Server version." -ForegroundColor Yellow
                $hasUnknownQuota = $true
                continue
            }

            $isExceeded = [double]$scenario.Count -ge [double]$scenario.Result.ThresholdValue
            if ($scenario.Result.Status -eq 'Incomplete') {
                Write-Host "The inventory is incomplete; $($scenario.Count) named properties were returned, so mailbox-wide quota status is unknown." -ForegroundColor Yellow
                $hasUnknownQuota = $true
            }
            elseif ($scenario.Result.Status -in @('Warning', 'Critical', 'Pass')) {
                if ($isExceeded) {
                    Write-Host "The mailbox has exceeded the threshold with $($scenario.Count) named properties." -ForegroundColor Red
                    $hasKnownExceededQuota = $true
                }
                else {
                    Write-Host "The mailbox is under quota with $($scenario.Count) named properties." -ForegroundColor Green
                }
            }
            elseif ($isExceeded) {
                Write-Host "The mailbox has exceeded the reference threshold with $($scenario.Count) named properties; on-premises applicability is unknown." -ForegroundColor Yellow
                $hasReferenceExceededQuota = $true
                $hasUnknownQuota = $true
            }
            else {
                Write-Host "The mailbox is under the reference quota with $($scenario.Count) named properties; on-premises applicability is unknown." -ForegroundColor Yellow
                $hasUnknownQuota = $true
            }
        }

        Write-Host 'Checking if mailbox has exceeded the threshold of any other named namespace quota.' -ForegroundColor Blue
        foreach ($namespaceResult in $namespaceResults) {
            $isExceeded = $null -ne $namespaceResult.ThresholdValue -and [double]$namespaceResult.CurrentValue -ge [double]$namespaceResult.ThresholdValue
            if ($namespaceResult.Status -eq 'Incomplete') {
                Write-Host "The $($namespaceResult.Scope) namespace has $($namespaceResult.CurrentValue) returned named properties; the inventory is incomplete, so mailbox-wide quota status is unknown." -ForegroundColor Yellow
                $hasUnknownQuota = $true
            }
            elseif ($namespaceResult.Status -in @('Warning', 'Critical', 'Pass')) {
                if ($isExceeded) {
                    Write-Host "The $($namespaceResult.Scope) namespace has exceeded quota with $($namespaceResult.CurrentValue) named properties. The quota is $($namespaceResult.Limit)." -ForegroundColor Red
                    $hasKnownExceededQuota = $true
                }
                else {
                    Write-Host "The $($namespaceResult.Scope) namespace is under quota with $($namespaceResult.CurrentValue) named properties." -ForegroundColor Green
                }
            }
            elseif ($isExceeded) {
                Write-Host "The $($namespaceResult.Scope) namespace exceeds the reference quota with $($namespaceResult.CurrentValue) named properties. The quota is $($namespaceResult.Limit); on-premises applicability is unknown." -ForegroundColor Yellow
                $hasReferenceExceededQuota = $true
                $hasUnknownQuota = $true
            }
            else {
                Write-Host "The $($namespaceResult.Scope) namespace is under the Exchange Online reference quota with $($namespaceResult.CurrentValue) named properties. The quota is $($namespaceResult.Limit); on-premises applicability is unknown." -ForegroundColor Yellow
                $hasUnknownQuota = $true
            }
        }

        Write-Host 'Checking if mailbox has exceeded the threshold of named properties with the same name prefix quota.' -ForegroundColor Blue
        $hasUnknownPrefixType = $false
        foreach ($prefixResult in $prefixResults) {
            $isExceeded = $null -ne $prefixResult.ThresholdValue -and [double]$prefixResult.CurrentValue -ge [double]$prefixResult.ThresholdValue
            if ($prefixResult.UnknownTypeCount -gt 0) {
                $prefixQuotaText = if ($null -ne $prefixResult.Limit) { "The quota is $($prefixResult.Limit)." } else { 'The applicable quota is unknown.' }
                $hasApplicableQuota = $null -ne $serverLimitByTest['RepeatedPropertyNamePrefix'] -or ($ExchangeOnlineQuotas -and $null -ne $referenceLimitByTest['RepeatedPropertyNamePrefix'])
                $isUnderApplicableQuota = $hasApplicableQuota -and [double]$prefixResult.CurrentValue -lt [double]$prefixResult.Limit
                if ($isUnderApplicableQuota -and -not $isExceeded) {
                    Write-Host "The $($prefixResult.PropertyNamePrefix) prefix is under quota with $($prefixResult.CurrentValue) named properties." -ForegroundColor Green
                }
                elseif ($isExceeded) {
                    Write-Host "The $($prefixResult.PropertyNamePrefix) prefix has $($prefixResult.CurrentValue) named properties. The quota is $($prefixResult.ThresholdValue)." -ForegroundColor Yellow
                    $hasCandidateThresholdExceeded = $true
                }
                elseif ($hasApplicableQuota) {
                    Write-Host "The $($prefixResult.PropertyNamePrefix) prefix has $($prefixResult.CurrentValue) named properties. Its warning threshold has been reached." -ForegroundColor Yellow
                }
                else {
                    Write-Host "The $($prefixResult.PropertyNamePrefix) prefix has $($prefixResult.CurrentValue) named properties. $prefixQuotaText On-premises applicability is unknown." -ForegroundColor Yellow
                }
                $hasUnknownPrefixType = $true
                $hasUnknownQuota = $true
            }
            elseif ($prefixResult.Status -eq 'Incomplete') {
                Write-Host "The $($prefixResult.PropertyNamePrefix) prefix has $($prefixResult.CurrentValue) returned named properties; the inventory is incomplete, so mailbox-wide quota status is unknown." -ForegroundColor Yellow
                $hasUnknownQuota = $true
            }
            elseif ($prefixResult.Status -in @('Warning', 'Critical', 'Pass')) {
                if ($isExceeded) {
                    Write-Host "The $($prefixResult.PropertyNamePrefix) prefix has exceeded quota with $($prefixResult.CurrentValue) named properties. The quota is $($prefixResult.Limit)." -ForegroundColor Red
                    $hasKnownExceededQuota = $true
                }
                else {
                    Write-Host "The $($prefixResult.PropertyNamePrefix) prefix is under quota with $($prefixResult.CurrentValue) named properties." -ForegroundColor Green
                }
            }
            elseif ($isExceeded) {
                Write-Host "The $($prefixResult.PropertyNamePrefix) prefix exceeds the reference quota with $($prefixResult.CurrentValue) named properties. The quota is $($prefixResult.Limit); on-premises applicability is unknown." -ForegroundColor Yellow
                $hasReferenceExceededQuota = $true
                $hasUnknownQuota = $true
            }
            else {
                Write-Host "The $($prefixResult.PropertyNamePrefix) prefix is under the Exchange Online reference quota with $($prefixResult.CurrentValue) named properties. The quota is $($prefixResult.Limit); on-premises applicability is unknown." -ForegroundColor Yellow
                $hasUnknownQuota = $true
            }
        }
        if ($hasUnknownPrefixType) {
            Write-Host 'Exact StringProperty usage is unknown because of missing or unrecognized property types.' -ForegroundColor Yellow
        }
        Write-Host "Summary, checking $SelectFirst result(s) and threshold of $([int]($Threshold * 100))%." -ForegroundColor Blue
        if ($hasKnownExceededQuota -or $hasCandidateThresholdExceeded) {
            Write-Host 'The mailbox has exceeded the threshold of a named properties quota. See above for which quota(s) have been exceeded.' -ForegroundColor Red
        }
        elseif ($hasReferenceExceededQuota) {
            Write-Host 'One or more reference thresholds were exceeded; on-premises quota applicability is unknown.' -ForegroundColor Yellow
        }
        elseif ($hasUnknownQuota) {
            Write-Host 'The mailbox quota status is unknown for this Exchange Server version or Store property types.' -ForegroundColor Yellow
        }
        else {
            Write-Host 'The mailbox is under the threshold for the reported named properties quotas.' -ForegroundColor Green
        }

        if ($PassThru) {
            foreach ($result in $results) {
                $result
            }
            foreach ($definition in $definitions) {
                $definition
            }
        }
    }
}