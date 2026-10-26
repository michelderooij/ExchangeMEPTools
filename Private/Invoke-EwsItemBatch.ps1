function Get-EwsItemBatch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Service,

        [Parameter(Mandatory = $true)]
        [string[]]$ItemId,

        [Parameter(Mandatory = $true)]
        [object[]]$PropertyDefinition,

        [ValidateRange(0, 10)]
        [int]$RetryCount = 3
    )

    $ewsItemIds = New-Object -TypeName 'System.Collections.Generic.List[Microsoft.Exchange.WebServices.Data.ItemId]'
    foreach ($uniqueId in $ItemId) {
        $ewsItemIds.Add((New-Object -TypeName Microsoft.Exchange.WebServices.Data.ItemId -ArgumentList $uniqueId))
    }
    $propertySet = New-Object -TypeName Microsoft.Exchange.WebServices.Data.PropertySet -ArgumentList ([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
    foreach ($definition in $PropertyDefinition) {
        $propertySet.Add($definition)
    }

    # BindToItems returns one GetItemResponse per requested ID, in request order.
    $responses = Invoke-EwsRequestWithRetry -RetryCount $RetryCount -ScriptBlock {
        $Service.BindToItems($ewsItemIds, $propertySet)
    }
    return , @($responses)
}

function Update-EwsItemBatch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Service,

        [Parameter(Mandatory = $true)]
        [object[]]$Item,

        [ValidateSet('AutoResolve', 'AlwaysOverwrite', 'NeverOverwrite')]
        [string]$ConflictResolution = 'AutoResolve',

        [ValidateRange(0, 10)]
        [int]$RetryCount = 3
    )

    $ewsItems = New-Object -TypeName 'System.Collections.Generic.List[Microsoft.Exchange.WebServices.Data.Item]'
    foreach ($ewsItem in $Item) {
        $ewsItems.Add($ewsItem)
    }
    $conflictMode = [Enum]::Parse([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode], $ConflictResolution, $true)

    # UpdateItems returns one UpdateItemResponse per item, in request order. SaveOnly and SendToNone
    # prevent message or meeting update traffic when only a property is removed.
    $responses = Invoke-EwsRequestWithRetry -RetryCount $RetryCount -ScriptBlock {
        $Service.UpdateItems(
            $ewsItems,
            $null,
            $conflictMode,
            [Microsoft.Exchange.WebServices.Data.MessageDisposition]::SaveOnly,
            [Microsoft.Exchange.WebServices.Data.SendInvitationsOrCancellationsMode]::SendToNone
        )
    }
    return , @($responses)
}
