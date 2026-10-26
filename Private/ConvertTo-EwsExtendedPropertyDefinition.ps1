function ConvertTo-EwsExtendedPropertyDefinition {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Definition
    )

    Import-EwsManagedApi | Out-Null

    $normalized = if ($Definition.PSObject.TypeNames -contains 'ExchangeNamedProperties.NamedPropertyDefinition') {
        $Definition
    }
    else {
        ConvertTo-NormalizedNamedProperty -Definition $Definition
    }

    try {
        $ewsTypeProperty = $normalized.PSObject.Properties['EwsMapiPropertyType']
        $mapiTypeName = if ($ewsTypeProperty) { [string]$ewsTypeProperty.Value } else { $null }
        if ([string]::IsNullOrWhiteSpace($mapiTypeName)) {
            $propertyTypeProperty = $normalized.PSObject.Properties['PropertyType']
            $mapiTypeName = if ($propertyTypeProperty) { [string]$propertyTypeProperty.Value } else { $null }
        }
        $mapiType = [Enum]::Parse([Microsoft.Exchange.WebServices.Data.MapiPropertyType], $mapiTypeName, $true)
        if (-not [Enum]::IsDefined([Microsoft.Exchange.WebServices.Data.MapiPropertyType], $mapiType)) {
            throw "MapiPropertyType '$mapiTypeName' is not defined by the loaded EWS Managed API."
        }

        $propertyTagProperty = $normalized.PSObject.Properties['PropertyTag']
        if ($propertyTagProperty -and $null -ne $propertyTagProperty.Value) {
            return New-Object -TypeName Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition -ArgumentList @([int]$propertyTagProperty.Value, $mapiType)
        }
        if ($normalized.PropertyName) {
            return New-Object -TypeName Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition -ArgumentList @([guid]$normalized.PropertySetId, [string]$normalized.PropertyName, $mapiType)
        }
        $propertyId = if ($normalized.PSObject.Properties['PropertyId']) { $normalized.PropertyId } elseif ($normalized.PSObject.Properties['PropDispId']) { $normalized.PropDispId } else { $null }
        if ($null -ne $propertyId) {
            return New-Object -TypeName Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition -ArgumentList @([guid]$normalized.PropertySetId, [int]$propertyId, $mapiType)
        }

        throw 'The property definition does not contain a supported property tag, name, or numeric ID.'
    }
    catch {
        throw "Could not convert named-property definition '$($normalized.Identity)' to an EWS ExtendedPropertyDefinition. Verify the property type and loaded EWS Managed API version: $($_.Exception.Message)"
    }
}