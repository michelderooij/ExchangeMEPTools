function Remove-MailboxExtendedProperty {
    <#
    .SYNOPSIS
    Removes one explicitly selected MAPI extended property from matching mailbox items.

    .DESCRIPTION
    Removes only the requested extended-property definition from each supplied item. It never deletes the mailbox item. Every input is validated and guarded by ShouldProcess, then approved removals are sent to EWS in batches: one batched bind and one batched update per batch. Progress is updated after each batch. Use PassThru to return structured results. Properties already absent from an item are reported as successful. Conflicts, access failures, and validation failures are reported separately.

    Configure the EWS connection with Connect-EWSSession before removal. The command uses the configured session and its cached service for the target mailbox.

    .PARAMETER Identity
    Optional target mailbox SMTP address, alias, GUID, or another identifier resolvable by Get-Mailbox. If omitted, the mailbox is inferred from pipeline search results.

    .PARAMETER MailboxExtendedProperty
    Search-MailboxExtendedProperty results with ItemId and a typed property definition to remove.

    .PARAMETER BatchSize
    Maximum number of removal requests sent in one batched EWS bind/update request. All inputs are processed; larger input sets use multiple batches. Default 100.

    .PARAMETER RetryCount
    Maximum retries for transient EWS server-busy responses. Permanent validation, access, and change-key errors are not retried.

    .PARAMETER Verify
    Rebind each batch after update and confirm the properties are absent.

    .PARAMETER PassThru
    Return structured removal results including the originating search result, and display property identity and subject for actual removals.

    .EXAMPLE
    $matches | Remove-MailboxExtendedProperty -Identity user@contoso.com -WhatIf

    .EXAMPLE
    $matches | Remove-MailboxExtendedProperty -Identity user@contoso.com -Confirm
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [Alias('MailboxIdentity', 'PrimarySmtpAddress')]
        [string]$Identity,

        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [Alias('SearchResult')]
        [psobject]$MailboxExtendedProperty,

        [ValidateRange(1, 1000)]
        [int]$BatchSize = 100,

        [ValidateRange(0, 10)]
        [int]$RetryCount,

        [switch]$Verify,

        [switch]$PassThru
    )

    begin {
        $totalCount = 0
        $mailboxCache = @{}
        $pendingEntries = New-Object 'System.Collections.Generic.List[object]'
        $removalCounts = @{ Removed = 0; AlreadyAbsent = 0; Skipped = 0; Failed = 0 }
        $progressActivity = 'Removing mailbox extended properties'
        if (-not $PSBoundParameters.ContainsKey('RetryCount')) {
            $RetryCount = if ($script:ConnectionOptions -and $script:ConnectionOptions.PSObject.Properties['RetryCount']) { [int]$script:ConnectionOptions.RetryCount } else { 3 }
        }
        function Get-RemovalErrorCode {
            param([string]$Message, [string]$ServiceErrorCode, [Exception]$Exception)

            $text = "$ServiceErrorCode $Message"
            if ($text -match '(?i)ErrorInvalidChangeKey|invalid change key|stale change key') { return 'StaleChangeKey' }
            if ($text -match '(?i)ErrorIrresolvableConflict|conflict') { return 'Conflict' }
            if ($text -match '(?i)ErrorAccessDenied|access denied|unauthorized') { return 'AccessDenied' }
            if ($Exception -is [System.ArgumentException]) { return 'ValidationError' }
            if ($Exception -is [System.InvalidOperationException]) { return 'MailboxMismatch' }
            if ($ServiceErrorCode) { return $ServiceErrorCode }
            if ($Exception) { return $Exception.GetType().Name }
            return 'Unknown'
        }

        function Set-RemovalFailure {
            param($Entry, [string]$ErrorCode, [string]$ErrorMessage)

            $Entry.Outcome = 'Failed'
            $Entry.Verified = $false
            $Entry.ErrorCode = $ErrorCode
            $Entry.ErrorMessage = $ErrorMessage
        }

        function New-RemovalResult {
            param($Entry)

            $removalCounts[$Entry.Outcome]++
            if (-not $PassThru) {
                if ($Entry.Outcome -eq 'Failed') {
                    Write-Warning ("Removal failed for item '{0}': {1}: {2}" -f $Entry.ItemId, $Entry.ErrorCode, $Entry.ErrorMessage)
                }
                return
            }
            if ($PassThru -and $Entry.Outcome -eq 'Removed') {
                Write-Host ('ExtendedPropertyId: {0} | Subject: {1}' -f $Entry.Definition.Identity, $Entry.Subject)
            }
            [pscustomobject]@{
                PSTypeName      = 'ExchangeNamedProperties.RemovalResult'
                MailboxIdentity = $Entry.MailboxIdentity
                ItemId          = $Entry.ItemId
                Subject         = $Entry.Subject
                FolderPath      = $Entry.FolderPath
                PropertySetId   = $Entry.PropertySetId
                PropertyId      = $Entry.PropertyId
                PropertyName    = if ($Entry.PropertyName) { $Entry.PropertyName } else { $null }
                PropertyType    = $Entry.PropertyType
                Action          = 'RemoveExtendedProperty'
                Succeeded       = $Entry.Outcome -in @('Removed', 'AlreadyAbsent')
                Verified        = $Entry.Verified
                ErrorCode       = $Entry.ErrorCode
                ErrorMessage    = $Entry.ErrorMessage
                Timestamp       = [DateTime]::UtcNow.ToString('o', [Globalization.CultureInfo]::InvariantCulture)
                Input           = if ($PassThru) { $Entry.Input } else { $null }
            }
        }

        function Write-RemovalProgress {
            param([int]$BatchNumber, [int]$BatchTotal)

            $processed = $removalCounts.Removed + $removalCounts.AlreadyAbsent + $removalCounts.Skipped + $removalCounts.Failed
            $percent = if ($totalCount -gt 0) { [int][Math]::Min(100, 100 * $processed / $totalCount) } else { 100 }
            $status = 'Batch {0} of {1} | Processed {2} of {3} | Removed {4} | Already absent {5} | Skipped {6} | Failed {7}' -f
            $BatchNumber, $BatchTotal, $processed, $totalCount,
            $removalCounts.Removed, $removalCounts.AlreadyAbsent, $removalCounts.Skipped, $removalCounts.Failed
            Write-Progress -Activity $progressActivity -Status $status -PercentComplete $percent
        }

        function Invoke-RemovalBatch {
            param($Service, [object[]]$Entries)

            # EWS item IDs are case-sensitive; group entries per item so each item is bound and updated once.
            $states = New-Object 'System.Collections.Generic.List[object]'
            $statesById = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([StringComparer]::Ordinal)
            $ewsDefinitions = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([StringComparer]::Ordinal)
            foreach ($entry in $Entries) {
                try {
                    $definitionKey = $entry.Definition.Identity
                    if (-not $ewsDefinitions.ContainsKey($definitionKey)) {
                        $ewsDefinitions[$definitionKey] = ConvertTo-EwsExtendedPropertyDefinition -Definition $entry.Definition
                    }
                    $entry.EwsDefinition = $ewsDefinitions[$definitionKey]
                    if (-not $statesById.ContainsKey($entry.ItemId)) {
                        $state = [pscustomobject]@{
                            ItemId        = $entry.ItemId
                            BoundItem     = $null
                            Entries       = New-Object 'System.Collections.Generic.List[object]'
                            Outcomes      = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([StringComparer]::Ordinal)
                            UpdateFailure = $null
                        }
                        $statesById[$entry.ItemId] = $state
                        $states.Add($state)
                    }
                    $statesById[$entry.ItemId].Entries.Add($entry)
                }
                catch {
                    Set-RemovalFailure -Entry $entry -ErrorCode (Get-RemovalErrorCode -Message $_.Exception.Message -Exception $_.Exception) -ErrorMessage $_.Exception.Message
                }
            }

            if ($states.Count -gt 0) {
                $itemIds = @($states | ForEach-Object { $_.ItemId })
                $definitions = @($ewsDefinitions.Values)
                $boundStates = New-Object 'System.Collections.Generic.List[object]'
                try {
                    $bindResponses = @(Get-EwsItemBatch -Service $Service -ItemId $itemIds -PropertyDefinition $definitions -RetryCount $RetryCount)
                    for ($index = 0; $index -lt $states.Count; $index++) {
                        $state = $states[$index]
                        $response = if ($index -lt $bindResponses.Count) { $bindResponses[$index] } else { $null }
                        $itemProperty = if ($response) { $response.PSObject.Properties['Item'] } else { $null }
                        $boundItem = if ($itemProperty -and $itemProperty -isnot [System.Management.Automation.PSParameterizedProperty]) { $itemProperty.Value } else { $null }
                        if (-not $response -or [string]$response.Result -eq 'Error' -or -not $boundItem) {
                            $serviceErrorCode = if ($response) { [string]$response.ErrorCode } else { $null }
                            $message = if ($response -and $response.ErrorMessage) { [string]$response.ErrorMessage } elseif ($itemProperty -is [System.Management.Automation.PSParameterizedProperty]) { 'EWS returned an Item member that PowerShell could not resolve to an item object.' } else { 'EWS did not return the item.' }
                            foreach ($entry in $state.Entries) {
                                Set-RemovalFailure -Entry $entry -ErrorCode (Get-RemovalErrorCode -Message $message -ServiceErrorCode $serviceErrorCode) -ErrorMessage $message
                            }
                            continue
                        }
                        $state.BoundItem = $boundItem
                        $boundStates.Add($state)
                    }
                }
                catch {
                    $message = $_.Exception.Message
                    $code = Get-RemovalErrorCode -Message $message -Exception $_.Exception
                    foreach ($state in $states) {
                        foreach ($entry in $state.Entries) {
                            Set-RemovalFailure -Entry $entry -ErrorCode $code -ErrorMessage $message
                        }
                    }
                }

                $dirtyStates = New-Object 'System.Collections.Generic.List[object]'
                foreach ($state in $boundStates) {
                    foreach ($entry in $state.Entries) {
                        $definitionKey = $entry.Definition.Identity
                        if ($state.Outcomes.ContainsKey($definitionKey)) {
                            continue
                        }
                        $existingValue = $null
                        if ($state.BoundItem.TryGetProperty($entry.EwsDefinition, [ref]$existingValue)) {
                            [void]$state.BoundItem.RemoveExtendedProperty($entry.EwsDefinition)
                            $state.Outcomes[$definitionKey] = 'Removed'
                        }
                        else {
                            $state.Outcomes[$definitionKey] = 'AlreadyAbsent'
                        }
                    }
                    if ($state.Outcomes.Values -contains 'Removed') {
                        $dirtyStates.Add($state)
                    }
                }

                if ($dirtyStates.Count -gt 0) {
                    try {
                        $updateResponses = @(Update-EwsItemBatch -Service $Service -Item @($dirtyStates | ForEach-Object { $_.BoundItem }) -RetryCount $RetryCount)
                        for ($index = 0; $index -lt $dirtyStates.Count; $index++) {
                            $response = if ($index -lt $updateResponses.Count) { $updateResponses[$index] } else { $null }
                            if (-not $response -or [string]$response.Result -eq 'Error') {
                                $serviceErrorCode = if ($response) { [string]$response.ErrorCode } else { $null }
                                $message = if ($response -and $response.ErrorMessage) { [string]$response.ErrorMessage } else { 'EWS did not return an update response for the item.' }
                                $dirtyStates[$index].UpdateFailure = [pscustomobject]@{
                                    ErrorCode    = Get-RemovalErrorCode -Message $message -ServiceErrorCode $serviceErrorCode
                                    ErrorMessage = $message
                                }
                            }
                        }
                    }
                    catch {
                        $failure = [pscustomobject]@{
                            ErrorCode    = Get-RemovalErrorCode -Message $_.Exception.Message -Exception $_.Exception
                            ErrorMessage = $_.Exception.Message
                        }
                        foreach ($state in $dirtyStates) {
                            $state.UpdateFailure = $failure
                        }
                    }
                }

                foreach ($state in $boundStates) {
                    foreach ($entry in $state.Entries) {
                        $outcome = $state.Outcomes[$entry.Definition.Identity]
                        if ($outcome -eq 'Removed' -and $state.UpdateFailure) {
                            Set-RemovalFailure -Entry $entry -ErrorCode $state.UpdateFailure.ErrorCode -ErrorMessage $state.UpdateFailure.ErrorMessage
                        }
                        else {
                            $entry.Outcome = $outcome
                            $entry.Verified = $null
                        }
                    }
                }

                if ($Verify) {
                    $verifyStates = @($boundStates | Where-Object { @($_.Entries | Where-Object { $_.Outcome -ne 'Failed' }).Count -gt 0 })
                    if ($verifyStates.Count -gt 0) {
                        $verifyResponses = @()
                        $verifyFailure = $null
                        try {
                            $verifyResponses = @(Get-EwsItemBatch -Service $Service -ItemId @($verifyStates | ForEach-Object { $_.ItemId }) -PropertyDefinition $definitions -RetryCount $RetryCount)
                        }
                        catch {
                            $verifyFailure = "Verification bind failed: $($_.Exception.Message)"
                        }
                        for ($index = 0; $index -lt $verifyStates.Count; $index++) {
                            $state = $verifyStates[$index]
                            $response = if ($index -lt $verifyResponses.Count) { $verifyResponses[$index] } else { $null }
                            $itemProperty = if ($response) { $response.PSObject.Properties['Item'] } else { $null }
                            $verifiedItem = if ($itemProperty -and $itemProperty -isnot [System.Management.Automation.PSParameterizedProperty]) { $itemProperty.Value } else { $null }
                            foreach ($entry in @($state.Entries | Where-Object { $_.Outcome -ne 'Failed' })) {
                                if ($verifyFailure) {
                                    Set-RemovalFailure -Entry $entry -ErrorCode 'VerificationFailed' -ErrorMessage $verifyFailure
                                }
                                elseif (-not $response -or [string]$response.Result -eq 'Error' -or -not $verifiedItem) {
                                    $message = if ($response -and $response.ErrorMessage) { [string]$response.ErrorMessage } elseif ($itemProperty -is [System.Management.Automation.PSParameterizedProperty]) { 'EWS returned an Item member that PowerShell could not resolve to an item object.' } else { 'EWS did not return the item.' }
                                    Set-RemovalFailure -Entry $entry -ErrorCode 'VerificationFailed' -ErrorMessage "Verification bind failed: $message"
                                }
                                else {
                                    $remainingValue = $null
                                    if ($verifiedItem.TryGetProperty($entry.EwsDefinition, [ref]$remainingValue)) {
                                        Set-RemovalFailure -Entry $entry -ErrorCode 'VerificationFailed' -ErrorMessage 'The extended property is still present after the update.'
                                    }
                                    else {
                                        $entry.Verified = $true
                                    }
                                }
                            }
                        }
                    }
                }
            }

            foreach ($entry in $Entries) {
                New-RemovalResult -Entry $entry
            }
        }
    }

    process {
        $totalCount++
        $inputResult = $MailboxExtendedProperty
        $sourceAddress = if ($inputResult.PSObject.Properties['PrimarySmtpAddress']) { [string]$inputResult.PrimarySmtpAddress } else { $null }
        $sourceIdentity = if ($sourceAddress) { $sourceAddress } elseif ($inputResult.PSObject.Properties['MailboxIdentity']) { [string]$inputResult.MailboxIdentity } else { $null }
        $effectiveIdentity = if ($Identity) { $Identity } else { $sourceIdentity }
        $propertyDefinition = $inputResult
        $propertySetId = if ($propertyDefinition.PSObject.Properties['PropertySetId']) { [string]$propertyDefinition.PropertySetId } else { $null }
        if (-not $propertySetId -and $propertyDefinition.PSObject.Properties['PropertyNamespace']) {
            $propertySetId = [string]$propertyDefinition.PropertyNamespace
        }
        $propertyId = if ($propertyDefinition.PSObject.Properties['PropertyId']) { $propertyDefinition.PropertyId } elseif ($propertyDefinition.PSObject.Properties['PropDispId']) { $propertyDefinition.PropDispId } else { $null }

        $entry = [pscustomobject]@{
            Input           = $inputResult
            MailboxIdentity = $effectiveIdentity
            ServiceIdentity = $effectiveIdentity
            MailboxKey      = $null
            ItemId          = if ($inputResult.PSObject.Properties['ItemId']) { [string]$inputResult.ItemId } else { $null }
            Subject         = if ($inputResult.PSObject.Properties['Subject']) { [string]$inputResult.Subject } else { $null }
            FolderPath      = if ($inputResult.PSObject.Properties['FolderPath']) { [string]$inputResult.FolderPath } else { $null }
            PropertySetId   = $propertySetId
            PropertyId      = $propertyId
            PropertyName    = if ($propertyDefinition.PSObject.Properties['PropertyName']) { [string]$propertyDefinition.PropertyName } else { $null }
            PropertyType    = if ($propertyDefinition.PSObject.Properties['PropertyType']) { [string]$propertyDefinition.PropertyType } else { $null }
            Definition      = $null
            EwsDefinition   = $null
            Outcome         = $null
            Verified        = $false
            ErrorCode       = $null
            ErrorMessage    = $null
        }

        try {
            if ([string]::IsNullOrWhiteSpace($effectiveIdentity)) {
                throw [System.ArgumentException]::new('Identity is missing. Supply Identity or pipeline Search-MailboxExtendedProperty results that include MailboxIdentity or PrimarySmtpAddress.')
            }
            if (-not $mailboxCache.ContainsKey($effectiveIdentity)) {
                try {
                    $mailboxCache[$effectiveIdentity] = Resolve-MailboxIdentity -Identity $effectiveIdentity
                }
                catch {
                    $mailboxCache[$effectiveIdentity] = $_.Exception
                }
            }
            $targetMailbox = $mailboxCache[$effectiveIdentity]
            if ($targetMailbox -is [Exception]) {
                throw [System.ArgumentException]::new($targetMailbox.Message)
            }
            $entry.MailboxIdentity = $targetMailbox.Identity
            $entry.MailboxKey = $targetMailbox.PrimarySmtpAddress.ToLowerInvariant()

            if ([string]::IsNullOrWhiteSpace($entry.ItemId)) {
                throw [System.ArgumentException]::new('Input does not contain ItemId. Pipe Search-MailboxExtendedProperty results; Get-MailboxExtendedProperty inventory rows do not identify an item.')
            }
            if ($sourceAddress -and
                -not [string]::Equals($sourceAddress, $targetMailbox.PrimarySmtpAddress, [StringComparison]::OrdinalIgnoreCase)) {
                throw [System.InvalidOperationException]::new("Input item mailbox '$sourceAddress' does not match target mailbox '$($targetMailbox.PrimarySmtpAddress)'.")
            }
            if ([string]::IsNullOrWhiteSpace($entry.PropertyType)) {
                throw [System.ArgumentException]::new('The selected property definition has no PropertyType; supply a validated type before removal.')
            }
            if (([string]::IsNullOrWhiteSpace($entry.PropertyName) -and $null -eq $propertyId) -or
                (-not [string]::IsNullOrWhiteSpace($entry.PropertyName) -and $null -ne $propertyId)) {
                throw [System.ArgumentException]::new('The selected definition must contain exactly one of PropertyName or PropertyId.')
            }

            $entry.Definition = ConvertTo-NormalizedNamedProperty -Definition ([pscustomobject]@{
                    PropertySetId = $propertySetId
                    PropertyName  = $entry.PropertyName
                    PropertyId    = $propertyId
                    PropertyType  = $entry.PropertyType
                })
            if ($PSCmdlet.ShouldProcess("$($targetMailbox.PrimarySmtpAddress) item $($entry.ItemId)", "Remove extended property $($entry.Definition.Identity)")) {
                $pendingEntries.Add($entry)
            }
            else {
                $entry.Outcome = 'Skipped'
                $entry.ErrorCode = 'ShouldProcessDeclined'
                $entry.ErrorMessage = 'Removal was not approved.'
                New-RemovalResult -Entry $entry
            }
        }
        catch {
            Set-RemovalFailure -Entry $entry -ErrorCode (Get-RemovalErrorCode -Message $_.Exception.Message -Exception $_.Exception) -ErrorMessage $_.Exception.Message
            New-RemovalResult -Entry $entry
        }
    }

    end {
        if ($pendingEntries.Count -eq 0) {
            return
        }

        $mailboxKeys = New-Object 'System.Collections.Generic.List[string]'
        $entriesByMailbox = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([StringComparer]::OrdinalIgnoreCase)
        foreach ($entry in $pendingEntries) {
            if (-not $entriesByMailbox.ContainsKey($entry.MailboxKey)) {
                $entriesByMailbox[$entry.MailboxKey] = New-Object 'System.Collections.Generic.List[object]'
                $mailboxKeys.Add($entry.MailboxKey)
            }
            $entriesByMailbox[$entry.MailboxKey].Add($entry)
        }
        $batchTotal = 0
        foreach ($mailboxKey in $mailboxKeys) {
            $batchTotal += [int][Math]::Ceiling($entriesByMailbox[$mailboxKey].Count / $BatchSize)
        }

        $batchNumber = 0
        Write-RemovalProgress -BatchNumber $batchNumber -BatchTotal $batchTotal
        try {
            foreach ($mailboxKey in $mailboxKeys) {
                $mailboxEntries = $entriesByMailbox[$mailboxKey]
                $service = $null
                $serviceFailure = $null
                try {
                    $service = Get-EwsService -Identity $mailboxEntries[0].ServiceIdentity
                }
                catch {
                    $serviceFailure = $_.Exception
                }

                for ($offset = 0; $offset -lt $mailboxEntries.Count; $offset += $BatchSize) {
                    $batchEntries = @($mailboxEntries.GetRange($offset, [Math]::Min($BatchSize, $mailboxEntries.Count - $offset)))
                    $batchNumber++
                    if ($serviceFailure) {
                        $code = Get-RemovalErrorCode -Message $serviceFailure.Message -Exception $serviceFailure
                        foreach ($entry in $batchEntries) {
                            Set-RemovalFailure -Entry $entry -ErrorCode $code -ErrorMessage $serviceFailure.Message
                            New-RemovalResult -Entry $entry
                        }
                    }
                    else {
                        Invoke-RemovalBatch -Service $service -Entries $batchEntries
                    }
                    Write-RemovalProgress -BatchNumber $batchNumber -BatchTotal $batchTotal
                }
            }
        }
        finally {
            Write-Progress -Activity $progressActivity -Completed
        }
    }
}
