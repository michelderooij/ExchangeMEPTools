function Get-StoreNamedProperty {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity,

        [string]$PropertyName,

        [switch]$ResolvePropertyType
    )

    $progressId = 73
    $activity = 'Retrieving mailbox named properties from Exchange Store'
    $queryContext = @{
        BatchSize        = 10
        ProbeOutcomes    = @{}
        ProgressId       = $progressId
        ProgressActivity = $activity
    }
    try {
        Write-Progress -Id $progressId -Activity $activity -Status 'Resolving mailbox and database...' -PercentComplete 5
        Get-StoreNamedPropertyCore -Identity $Identity -ProgressId $progressId -ProgressActivity $activity -PropertyName $PropertyName -ResolvePropertyType:$ResolvePropertyType -QueryContext $queryContext
    }
    finally {
        Write-Progress -Id $progressId -Activity $activity -Completed
    }
}

function Get-StoreNamedPropertyCore {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity,

        [Parameter(Mandatory = $true)]
        [int]$ProgressId,

        [Parameter(Mandatory = $true)]
        [string]$ProgressActivity,

        [string]$PropertyName,

        [switch]$ResolvePropertyType,

        [hashtable]$QueryContext
    )

    if (-not (Get-Command -Name 'Get-Mailbox' -ErrorAction SilentlyContinue)) {
        throw 'The Store provider requires Get-Mailbox. Run this command in the Exchange Management Shell or import the Exchange mailbox management commands.'
    }

    if (-not (Get-Command -Name 'Get-MailboxDatabase' -ErrorAction SilentlyContinue)) {
        throw 'The Store provider requires Get-MailboxDatabase. Run this command in the Exchange Management Shell or import the Exchange mailbox management commands.'
    }

    try {
        $mailbox = Get-Mailbox -Identity $Identity -ErrorAction Stop
    }
    catch {
        throw "Identity '$Identity' could not be resolved by Get-Mailbox: $($_.Exception.Message)"
    }

    $primarySmtpAddress = [string]$mailbox.PrimarySmtpAddress
    if ([string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
        throw "Get-Mailbox resolved '$Identity' but did not return PrimarySmtpAddress."
    }

    $mailboxGuid = [guid]::Empty
    if (-not [guid]::TryParse([string]$mailbox.ExchangeGuid, [ref]$mailboxGuid) -or $mailboxGuid -eq [guid]::Empty) {
        throw "Get-Mailbox did not return a valid ExchangeGuid for '$Identity'."
    }

    $databaseName = [string]$mailbox.Database
    if ([string]::IsNullOrWhiteSpace($databaseName)) {
        throw "Get-Mailbox did not return a primary mailbox database for '$Identity'."
    }

    try {
        $database = Get-MailboxDatabase -Identity $databaseName -ErrorAction Stop
    }
    catch {
        throw "Mailbox database '$databaseName' could not be resolved: $($_.Exception.Message)"
    }

    $databaseIdentity = [string]$database.Name
    if ([string]::IsNullOrWhiteSpace($databaseIdentity)) {
        $databaseIdentity = $databaseName
    }

    Write-Progress -Id $ProgressId -Activity $ProgressActivity -Status 'Reading mailbox metadata...' -PercentComplete 30

    $mailboxQuery = 'SELECT MailboxGuid,MailboxInstanceGuid,DisplayName,MailboxNumber FROM Mailbox WHERE MailboxGuid="{0}"' -f $mailboxGuid.ToString('D')
    $mailboxRows = @(Invoke-StoreQuery -Database $databaseIdentity -Query $mailboxQuery -QueryContext $queryContext)
    if ($mailboxRows.Count -ne 1) {
        throw "The Store provider expected one mailbox metadata row for '$primarySmtpAddress' in database '$databaseIdentity' and received $($mailboxRows.Count)."
    }

    $mailboxRow = $mailboxRows[0]
    $mailboxPartitionNumber = 0
    if (-not [int]::TryParse([string]$mailboxRow.MailboxNumber, [ref]$mailboxPartitionNumber) -or $mailboxPartitionNumber -lt 0) {
        throw "The Store mailbox metadata for '$primarySmtpAddress' did not contain a valid MailboxNumber."
    }

    $storeMailboxGuid = [guid]::Empty
    if (-not [guid]::TryParse([string]$mailboxRow.MailboxGuid, [ref]$storeMailboxGuid) -or $storeMailboxGuid -eq [guid]::Empty) {
        throw "The Store mailbox metadata for '$primarySmtpAddress' did not contain a valid MailboxGuid."
    }

    $mailboxInstanceGuid = [guid]::Empty
    $hasMailboxInstanceGuid = [guid]::TryParse([string]$mailboxRow.MailboxInstanceGuid, [ref]$mailboxInstanceGuid) -and
    $mailboxInstanceGuid -ne [guid]::Empty

    $namedPropertyQuery = 'SELECT PropNumber,PropGuid,PropName,PropDispId FROM ExtendedPropertyNameMapping WHERE MailboxPartitionNumber={0}' -f $mailboxPartitionNumber
    Write-Progress -Id $ProgressId -Activity $ProgressActivity -Status 'Fetching named-property mapping rows from the Store...' -PercentComplete 60
    $namedPropertyRows = @(Invoke-StoreQuery -Database $databaseIdentity -Query $namedPropertyQuery -QueryContext $queryContext)
    $namedPropCount = $namedPropertyRows.Count
    if (-not [string]::IsNullOrWhiteSpace($PropertyName)) {
        $namedPropertyRows = @($namedPropertyRows | Where-Object {
                $_.PSObject.Properties['PropName'] -and
                -not [string]::IsNullOrWhiteSpace([string]$_.PropName) -and
                [string]$_.PropName -like $PropertyName
            })
    }
    $selectedNamedPropertyCount = $namedPropertyRows.Count
    $normalizationStatus = if ($ResolvePropertyType) { "Normalizing and resolving types for {0} selected Store mapping rows..." -f $selectedNamedPropertyCount } else { "Normalizing {0} selected Store mapping rows..." -f $selectedNamedPropertyCount }
    Write-Progress -Id $ProgressId -Activity $ProgressActivity -Status $normalizationStatus -PercentComplete 75

    if ($ResolvePropertyType -and $selectedNamedPropertyCount -gt 0) {
        $progressStatus = 'Resolving type for selected Store mapping 1 of {0}...' -f $selectedNamedPropertyCount
        Write-Progress -Id $ProgressId -Activity $ProgressActivity -Status $progressStatus -PercentComplete 75
    }
    $resolutionWatch = if ($ResolvePropertyType) { [Diagnostics.Stopwatch]::StartNew() } else { $null }
    if ($ResolvePropertyType) {
        $QueryContext.ProgressWatch = $resolutionWatch
        $QueryContext.ProgressTotalCount = $selectedNamedPropertyCount
        $QueryContext.ProgressCompletedCount = 0
        $QueryContext.ProgressLastReportedCount = 0
    }
    $batchTypes = @{}
    if ($ResolvePropertyType -and $QueryContext.BatchSize -gt 1) {
        $batchNumbers = @(
            foreach ($candidateRow in $namedPropertyRows) {
                $candidateNumber = 0
                $candidateGuid = [guid]::Empty
                if ([int]::TryParse([string]$candidateRow.PropNumber, [ref]$candidateNumber) -and
                    $candidateNumber -ge 32768 -and $candidateNumber -le 65534 -and
                    [guid]::TryParse([string]$candidateRow.PropGuid, [ref]$candidateGuid) -and $candidateGuid -ne [guid]::Empty) {
                    $candidateNumber
                }
            }
        )
        if ($batchNumbers.Count -gt 1) {
            $batchTypes = Get-StoreNamedPropertyBatchedTypes -Database $databaseIdentity -MailboxPartitionNumber $mailboxPartitionNumber -PropNumbers $batchNumbers -BatchSize $QueryContext.BatchSize -QueryContext $QueryContext
        }
        $fallbackCount = $batchNumbers.Count - $batchTypes.Count
        if ($fallbackCount -gt 0) {
            Write-Verbose ("Using individual type-resolution fallback for {0} mapping(s) not resolved by batch probing; successful negative probes will be reused." -f $fallbackCount)
        }
    }

    $rowIndex = 0
    $progressInterval = if ($ResolvePropertyType) { 10 } else { 25 }
    foreach ($namedProperty in $namedPropertyRows) {
        $rowIndex++
        if (-not $ResolvePropertyType -and (($rowIndex % $progressInterval -eq 0) -or $rowIndex -eq $selectedNamedPropertyCount)) {
            $percentComplete = 75 + [int](25 * $rowIndex / [Math]::Max($selectedNamedPropertyCount, 1))
            $progressStatus = "Normalizing selected Store mapping {0} of {1}..." -f $rowIndex, $selectedNamedPropertyCount
            Write-Progress -Id $ProgressId -Activity $ProgressActivity -Status $progressStatus -PercentComplete $percentComplete
        }

        $propertySetId = [guid]::Empty
        if (-not [guid]::TryParse([string]$namedProperty.PropGuid, [ref]$propertySetId) -or $propertySetId -eq [guid]::Empty) {
            Write-Error -Message "A Store mapping row for mailbox '$primarySmtpAddress' contained an invalid property-set GUID." -Category InvalidData -ErrorId 'InvalidStorePropertySetId' -TargetObject $namedProperty
            continue
        }

        $propertyNumber = 0
        if (-not [int]::TryParse([string]$namedProperty.PropNumber, [ref]$propertyNumber) -or $propertyNumber -lt 32768 -or $propertyNumber -gt 65534) {
            Write-Error -Message "A Store mapping row for mailbox '$primarySmtpAddress' contained an invalid PropNumber." -Category InvalidData -ErrorId 'InvalidStorePropertyNumber' -TargetObject $namedProperty
            continue
        }

        $propertyName = [string]$namedProperty.PropName
        if ($propertyName -eq 'NULL' -and $namedProperty.PSObject.Properties['PropDispId'] -and
            $null -ne $namedProperty.PropDispId -and [string]$namedProperty.PropDispId -ne 'NULL') {
            $propertyName = $null
        }
        $propertyId = $null
        if ([string]::IsNullOrWhiteSpace($propertyName)) {
            $dispid = 0
            if ([int]::TryParse([string]$namedProperty.PropDispId, [ref]$dispid) -and $dispid -ge 0) {
                $propertyId = $dispid
            }
            else {
                Write-Error -Message "A Store mapping row for mailbox '$primarySmtpAddress' had neither a property name nor a valid PropDispId." -Category InvalidData -ErrorId 'InvalidStorePropertyName' -TargetObject $namedProperty
                continue
            }
        }

        $nameOrId = if ($propertyName) { "Name:$propertyName" } else { "Id:$propertyId" }
        if ($ResolvePropertyType) {
            if ($batchTypes.ContainsKey($propertyNumber)) {
                $propertyTypeInfo = $batchTypes[$propertyNumber]
            }
            else {
                $propertyTypeInfo = Get-StoreNamedPropertyType -Database $databaseIdentity -MailboxPartitionNumber $mailboxPartitionNumber -PropNumber $propertyNumber -QueryContext $queryContext
            }
        }
        $mailboxLocation = if ($hasMailboxInstanceGuid) {
            '{0}\{1}' -f $storeMailboxGuid.ToString('D'), $mailboxInstanceGuid.ToString('D')
        }
        else {
            $storeMailboxGuid.ToString('D')
        }

        $definition = [pscustomobject]@{
            PSTypeName        = 'ExchangeNamedProperties.NamedPropertyDefinition'
            MailboxIdentity   = $Identity
            DisplayName       = if ($mailboxRow.PSObject.Properties['DisplayName']) { [string]$mailboxRow.DisplayName } else { $null }
            NamedPropCount    = $namedPropCount
            PropertyNamespace = $propertySetId.ToString('D').ToLowerInvariant()
            PropertyName      = if ($propertyName) { $propertyName } else { $null }
            PropertySetId     = $propertySetId.ToString('D').ToLowerInvariant()
            Identity          = '{0}:{1}\{2}' -f $mailboxLocation, $propertySetId.ToString('D').ToLowerInvariant(), $nameOrId
            PropNumber        = $propertyNumber
            PropDispId        = if ($null -ne $propertyId) { $propertyId } else { $null }
            PropertyType      = $null
            PropertyTypeCode  = $null
        }
        if ($ResolvePropertyType) {
            $definition.PropertyType = $propertyTypeInfo.PropertyType
            $definition.PropertyTypeCode = $propertyTypeInfo.PropertyTypeCode
            if (-not $batchTypes.ContainsKey($propertyNumber)) { $QueryContext.ProgressCompletedCount++ }
            Write-StoreTypeResolutionProgress -QueryContext $QueryContext
        }
        $definition
    }
}

function Write-StoreTypeResolutionProgress {
    [CmdletBinding()]
    param([hashtable]$QueryContext)

    if ($null -eq $QueryContext -or -not $QueryContext.ContainsKey('ProgressTotalCount')) { return }
    $completedCount = $QueryContext.ProgressCompletedCount
    $lastReportedCount = $QueryContext.ProgressLastReportedCount
    if ($completedCount -le $lastReportedCount) { return }
    if ($lastReportedCount -eq 0 -or
        [Math]::Floor($completedCount / 10) -gt [Math]::Floor($lastReportedCount / 10) -or
        $completedCount -eq $QueryContext.ProgressTotalCount) {
        $status = Get-StoreTypeResolutionProgressStatus -CompletedCount $completedCount -TotalCount $QueryContext.ProgressTotalCount -ElapsedSeconds $QueryContext.ProgressWatch.Elapsed.TotalSeconds
        $percentComplete = 75 + [int](25 * $completedCount / [Math]::Max($QueryContext.ProgressTotalCount, 1))
        Write-Progress -Id $QueryContext.ProgressId -Activity $QueryContext.ProgressActivity -Status $status -PercentComplete $percentComplete
        $QueryContext.ProgressLastReportedCount = $completedCount
    }
}

function Get-StoreTypeResolutionProgressStatus {
    [CmdletBinding()]
    param(
        [ValidateRange(0, 2147483647)]
        [int]$CompletedCount,

        [ValidateRange(0, 2147483647)]
        [int]$TotalCount,

        [ValidateRange(0.0, [double]::MaxValue)]
        [double]$ElapsedSeconds
    )

    $speed = if ($ElapsedSeconds -gt 0) { $CompletedCount / $ElapsedSeconds } else { 0.0 }
    $remainingCount = [Math]::Max(0, $TotalCount - $CompletedCount)
    $remainingTime = '--:--'
    if ($remainingCount -eq 0) {
        $remainingTime = '00:00'
    }
    elseif ($speed -gt 0) {
        $minimumSamples = [Math]::Max(10, [Math]::Ceiling($TotalCount * 0.1))
        if ($CompletedCount -lt $minimumSamples -or $ElapsedSeconds -lt 15) {
            $remainingTime = 'Calculating...'
        }
        else {
            $remainingSeconds = [Math]::Ceiling($remainingCount / $speed)
            $remainingTime = '{0:00}:{1:00}' -f [Math]::Floor($remainingSeconds / 60), ($remainingSeconds % 60)
        }
    }

    'Resolving type for selected Store mapping {0} of {1}... | {2:N2} type/sec | Est. time remaining: {3}' -f $CompletedCount, $TotalCount, $speed, $remainingTime
}

function Get-StoreNamedPropertyBatchedTypes {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Database,

        [Parameter(Mandatory = $true)]
        [int]$MailboxPartitionNumber,

        [Parameter(Mandatory = $true)]
        [int[]]$PropNumbers,

        [ValidateRange(2, 10)]
        [int]$BatchSize = 5,

        [hashtable]$QueryContext
    )

    if ($null -eq $QueryContext) { $QueryContext = @{} }
    $resolved = @{}
    $pending = @($PropNumbers | Select-Object -Unique)
    foreach ($candidate in Get-StoreNamedPropertyTypeCandidates) {
        if ($pending.Count -lt 2 -or $QueryContext['BatchDisabled']) { break }
        $observed = Get-StoreNamedPropertyBatchType -Database $Database -MailboxPartitionNumber $MailboxPartitionNumber -PropNumbers $pending -BatchSize $BatchSize -TypeCode $candidate.Key -PropertyType $candidate.Value -QueryContext $QueryContext
        foreach ($number in $observed.Keys) { $resolved[$number] = $observed[$number] }
        $pending = @($pending | Where-Object {
                $probeKey = '{0}|{1}|p{2:x}{3}' -f $Database, $MailboxPartitionNumber, $_, $candidate.Key
                -not $resolved.ContainsKey($_) -and $QueryContext.ProbeOutcomes[$probeKey] -eq 'NotObserved'
            })
    }
    $resolved
}

function Get-StoreNamedPropertyBatchType {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Database,

        [Parameter(Mandatory = $true)]
        [int]$MailboxPartitionNumber,

        [Parameter(Mandatory = $true)]
        [ValidateRange(32768, 65534)]
        [int[]]$PropNumbers,

        [ValidateRange(2, 10)]
        [int]$BatchSize = 5,

        [ValidatePattern('^[0-9a-f]{4}$')]
        [string]$TypeCode = '001f',

        [string]$PropertyType = 'String',

        [hashtable]$QueryContext
    )

    if ($null -eq $QueryContext) { $QueryContext = @{} }
    if (-not $QueryContext.ContainsKey('ProbeOutcomes')) { $QueryContext.ProbeOutcomes = @{} }
    if (-not $QueryContext.ContainsKey('BatchFailuresRemaining')) { $QueryContext.BatchFailuresRemaining = 3 }
    $resolved = @{}
    $uniqueNumbers = @($PropNumbers | Select-Object -Unique)
    $groups = New-Object 'System.Collections.Generic.Queue[object]'
    for ($offset = 0; $offset -lt $uniqueNumbers.Count; $offset += $BatchSize) {
        $lastIndex = [Math]::Min($offset + $BatchSize, $uniqueNumbers.Count) - 1
        $groups.Enqueue(@($uniqueNumbers[$offset..$lastIndex]))
    }
    while ($groups.Count -gt 0 -and -not $QueryContext['BatchDisabled']) {
        $pending = @($groups.Dequeue())
        while ($pending.Count -gt 1 -and -not $QueryContext['BatchDisabled']) {
            $tags = @($pending | ForEach-Object { 'p{0:x}{1}' -f $_, $TypeCode })
            $restrictions = @($tags | ForEach-Object { '{0} != null' -f $_ })
            $query = 'SELECT TOP 1 {0} FROM Message WHERE MailboxPartitionNumber={1} AND ({2})' -f ($tags -join ','), $MailboxPartitionNumber, ($restrictions -join ' OR ')
            try {
                $messages = @(Invoke-StoreQuery -Database $Database -Query $query -QueryContext $QueryContext)
            }
            catch {
                foreach ($tag in $tags) {
                    $QueryContext.ProbeOutcomes['{0}|{1}|{2}' -f $Database, $MailboxPartitionNumber, $tag] = 'Failed'
                }
                $QueryContext.BatchFailuresRemaining--
                if ($QueryContext.BatchFailuresRemaining -le 0) {
                    $QueryContext.BatchDisabled = $true
                    Write-Verbose 'Type-probe batching disabled after three failed batch queries. Falling back to individual probes for remaining mappings.'
                }
                elseif ($pending.Count -gt 2) {
                    $middle = [int][Math]::Floor($pending.Count / 2)
                    $groups.Enqueue(@($pending[0..($middle - 1)]))
                    $groups.Enqueue(@($pending[$middle..($pending.Count - 1)]))
                    Write-Verbose ("Batch type probe failed for {0} mapping(s), type {1}; retrying smaller groups before individual fallback. Batch failures remaining: {2}." -f $pending.Count, $TypeCode, $QueryContext.BatchFailuresRemaining)
                }
                else {
                    Write-Verbose ("Batch type probe failed for {0} mapping(s), type {1}; falling back to individual probes for these mappings. Batch failures remaining: {2}." -f $pending.Count, $TypeCode, $QueryContext.BatchFailuresRemaining)
                }
                break
            }

            if ($messages.Count -eq 0) {
                foreach ($tag in $tags) {
                    $QueryContext.ProbeOutcomes['{0}|{1}|{2}' -f $Database, $MailboxPartitionNumber, $tag] = 'NotObserved'
                }
            }
            $observedCount = 0
            foreach ($number in $pending) {
                $tag = 'p{0:x}{1}' -f $number, $TypeCode
                $probeKey = '{0}|{1}|{2}' -f $Database, $MailboxPartitionNumber, $tag
                if ($messages.Count -gt 0) { $QueryContext.ProbeOutcomes[$probeKey] = 'Inconclusive' }
                foreach ($message in $messages) {
                    if ($null -eq $message) { continue }
                    $property = $message.PSObject.Properties[$tag]
                    if ($property -and $null -ne $property.Value -and $property.Value -isnot [DBNull] -and
                        -not ($property.Value -is [string] -and $property.Value -ceq 'NULL')) {
                        $resolved[$number] = [pscustomobject]@{
                            PropertyType       = $PropertyType
                            PropertyTypeCode   = $TypeCode
                            PropertyTypeStatus = 'Observed'
                            PropertyTypeError  = $null
                        }
                        $QueryContext.ProbeOutcomes[$probeKey] = 'Observed'
                        $observedCount++
                        break
                    }
                }
            }
            if ($QueryContext.ContainsKey('ProgressTotalCount')) {
                $QueryContext.ProgressCompletedCount += $observedCount
                Write-StoreTypeResolutionProgress -QueryContext $QueryContext
            }
            if ($observedCount -eq 0) { break }
            $pending = @($pending | Where-Object { -not $resolved.ContainsKey($_) })
        }
    }
    $resolved
}

function Get-StoreNamedPropertyTypeCandidates {
    ([ordered]@{
        '001f' = 'String'
        '001e' = 'String'
        '0000' = 'Unspecified'
        '0001' = 'Null'
        '0002' = 'Short'
        '0003' = 'Integer'
        '0004' = 'Float'
        '0005' = 'Double'
        '0006' = 'Currency'
        '0007' = 'ApplicationTime'
        '000a' = 'Error'
        '000b' = 'Boolean'
        '000d' = 'Object'
        '0014' = 'Long'
        '0040' = 'SystemTime'
        '0048' = 'CLSID'
        '00fb' = 'ServerEntryId'
        '0105' = 'PT_0105'
        '00fd' = 'Restriction'
        '00fe' = 'Actions'
        '0102' = 'Binary'
    }).GetEnumerator()
}

function Get-StoreNamedPropertyType {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Database,

        [Parameter(Mandatory = $true)]
        [int]$MailboxPartitionNumber,

        [Parameter(Mandatory = $true)]
        [int]$PropNumber,

        [hashtable]$QueryContext
    )

    $typeCandidates = @(Get-StoreNamedPropertyTypeCandidates)
    $typeCodes = @($typeCandidates | ForEach-Object { $_.Key })
    $propertyNumberHex = '{0:x}' -f $PropNumber
    $queryErrors = New-Object 'System.Collections.Generic.List[object]'

    foreach ($candidate in $typeCandidates) {
        $typeCode = $candidate.Key
        $propertyTag = 'p{0}{1}' -f $propertyNumberHex, $typeCode
        $probeKey = '{0}|{1}|{2}' -f $Database, $MailboxPartitionNumber, $propertyTag
        if ($QueryContext -and $QueryContext['ProbeOutcomes'] -and
            $QueryContext.ProbeOutcomes[$probeKey] -eq 'NotObserved') {
            continue
        }
        $query = 'SELECT TOP 1 {1} FROM Message WHERE MailboxPartitionNumber={0} AND {1} != null' -f $MailboxPartitionNumber, $propertyTag
        try {
            $messages = @(Invoke-StoreQuery -Database $Database -Query $query -QueryContext $QueryContext)
        }
        catch {
            $queryErrors.Add([pscustomobject]@{ PropertyTypeCode = $typeCode; Message = $_.Exception.Message })
            continue
        }

        $matchingProperty = @($messages | Where-Object {
                $property = $_.PSObject.Properties[$propertyTag]
                $property -and $null -ne $property.Value -and
                $property.Value -isnot [DBNull] -and
                -not ($property.Value -is [string] -and $property.Value -ceq 'NULL')
            } | Select-Object -First 1)
        if ($matchingProperty.Count -gt 0) {
            return [pscustomobject]@{
                PropertyType       = $candidate.Value
                PropertyTypeCode   = $typeCode
                PropertyTypeStatus = 'Observed'
                PropertyTypeError  = $null
            }
        }
    }

    if ($queryErrors.Count -eq $typeCodes.Count) {
        $failedTypeCodes = @($queryErrors | Select-Object -ExpandProperty PropertyTypeCode) -join ', '
        return [pscustomobject]@{
            PropertyType       = $null
            PropertyTypeCode   = $null
            PropertyTypeStatus = 'QueryFailed'
            PropertyTypeError  = "No type was observed; all Store queries failed for candidate MAPI type codes $failedTypeCodes. First error: $($queryErrors[0].Message)"
        }
    }

    if ($queryErrors.Count -gt 0) {
        $failedTypeCodes = @($queryErrors | Select-Object -ExpandProperty PropertyTypeCode) -join ', '
        $successfulProbeCount = $typeCodes.Count - $queryErrors.Count
        return [pscustomobject]@{
            PropertyType       = $null
            PropertyTypeCode   = $null
            PropertyTypeStatus = 'PartialQueryFailure'
            PropertyTypeError  = "No type was observed. $successfulProbeCount candidate Store query/queries completed without finding a non-null property value; queries for candidate MAPI type codes $failedTypeCodes failed. First error: $($queryErrors[0].Message)"
        }
    }

    [pscustomobject]@{
        PropertyType       = $null
        PropertyTypeCode   = $null
        PropertyTypeStatus = 'NotObserved'
        PropertyTypeError  = $null
    }
}