function ConvertTo-NormalizedNamedProperty {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $true, ValueFromPipeline = $true)]
		[psobject]$Definition,

		[string]$DiscoverySource = 'Pipeline',

		[string]$MailboxIdentity,

		[string]$PrimarySmtpAddress
	)

	process {
		$propertySetId = $null
		if ($Definition.PSObject.Properties['PropertySetId']) {
			$propertySetId = [string]$Definition.PropertySetId
		}
		elseif ($Definition.PSObject.Properties['PropertyNamespace']) {
			$propertySetId = [string]$Definition.PropertyNamespace
		}

		if ($propertySetId) {
			$parsedGuid = [guid]::Empty
			if (-not [guid]::TryParse($propertySetId, [ref]$parsedGuid) -or $parsedGuid -eq [guid]::Empty) {
				throw "PropertySetId '$propertySetId' is not a valid non-empty GUID."
			}
			$propertySetId = $parsedGuid.ToString('D').ToLowerInvariant()
		}

		$propertyName = $null
		if ($Definition.PSObject.Properties['PropertyName'] -and
			-not [string]::IsNullOrWhiteSpace([string]$Definition.PropertyName)) {
			$propertyName = [string]$Definition.PropertyName
		}

		$propertyId = $null
		$propertyIdText = $null
		$isStorePropDispId = $false
		$propertyIdProperty = $Definition.PSObject.Properties['PropertyId']
		if ($propertyIdProperty -and $null -ne $propertyIdProperty.Value -and
			-not [string]::IsNullOrWhiteSpace([string]$propertyIdProperty.Value)) {
			$propertyIdText = [string]$propertyIdProperty.Value
		}
		else {
			$propDispIdProperty = $Definition.PSObject.Properties['PropDispId']
			if ($propDispIdProperty -and $null -ne $propDispIdProperty.Value -and
				-not [string]::IsNullOrWhiteSpace([string]$propDispIdProperty.Value)) {
				$propertyIdText = [string]$propDispIdProperty.Value
				$isStorePropDispId = $true
			}
		}
		if ($null -ne $propertyIdText) {
			$parsedId = 0
			if (-not [int]::TryParse($propertyIdText, [ref]$parsedId) -or
				($isStorePropDispId -and ($parsedId -lt 0 -or $parsedId -gt 65535)) -or
				(-not $isStorePropDispId -and ($parsedId -lt 32768 -or $parsedId -gt 65534))) {
				if ($isStorePropDispId) {
					throw "PropDispId '$propertyIdText' must be a non-negative 16-bit value."
				}
				throw "PropertyId '$propertyIdText' must be a numeric named-property ID from 32768 through 65534."
			}
			$propertyId = $parsedId
		}

		$propertyTag = $null
		if ($Definition.PSObject.Properties['PropertyTag'] -and $null -ne $Definition.PropertyTag -and
			-not [string]::IsNullOrWhiteSpace([string]$Definition.PropertyTag)) {
			$tagText = ([string]$Definition.PropertyTag).Trim()
			$parsedTag = [uint32]0
			$tagStyle = [Globalization.NumberStyles]::Integer
			if ($tagText.StartsWith('0x', [StringComparison]::OrdinalIgnoreCase)) {
				$tagText = $tagText.Substring(2)
				$tagStyle = [Globalization.NumberStyles]::HexNumber
			}
			if (-not [uint32]::TryParse($tagText, $tagStyle, [Globalization.CultureInfo]::InvariantCulture, [ref]$parsedTag)) {
				throw "PropertyTag '$($Definition.PropertyTag)' must be a 32-bit decimal or hexadecimal value."
			}
			$propertyTag = $parsedTag
		}

		if ($null -ne $propertyTag) {
			if ($propertySetId -or $propertyName -or $null -ne $propertyId) {
				throw 'PropertyTag cannot be combined with a named-property PropertySetId, PropertyName, or PropertyId.'
			}
		}
		elseif (-not $propertySetId -or (($null -eq $propertyName) -eq ($null -eq $propertyId))) {
			throw 'A definition must contain PropertySetId and exactly one of PropertyName or PropertyId, or a standalone PropertyTag.'
		}

		if (-not $Definition.PSObject.Properties['PropertyType'] -or
			[string]::IsNullOrWhiteSpace([string]$Definition.PropertyType)) {
			throw 'PropertyType is required.'
		}

		$typeInput = ([string]$Definition.PropertyType).Trim()
		$typeAliases = @{
			'DateTime' = 'SystemTime'
			'Guid'     = 'CLSID'
			'Int16'    = 'Short'
			'Int32'    = 'Integer'
			'Int64'    = 'Long'
		}
		if ($typeAliases.ContainsKey($typeInput)) {
			$typeInput = $typeAliases[$typeInput]
		}

		$supportedTypes = @('ApplicationTime', 'Binary', 'Boolean', 'CLSID', 'Currency', 'Double', 'Error', 'Float', 'Integer', 'Long', 'Null', 'Object', 'Short', 'SystemTime', 'String')
		$normalizedType = $supportedTypes |
		Where-Object { [string]::Equals($_, $typeInput, [StringComparison]::OrdinalIgnoreCase) } |
		Select-Object -First 1
		if (-not $normalizedType) {
			throw "PropertyType '$($Definition.PropertyType)' is unsupported. Supported types are: $($supportedTypes -join ', ')."
		}

		$nameOrId = if ($propertyName) { "Name:$propertyName" } else { "Id:$propertyId" }
		$definitionIdentity = if ($null -ne $propertyTag) {
			'Tag:{0:X8}' -f $propertyTag
		}
		else {
			'{0}\{1}' -f $propertySetId, $nameOrId
		}

		[pscustomobject]@{
			PSTypeName            = 'ExchangeNamedProperties.NamedPropertyDefinition'
			MailboxIdentity       = $MailboxIdentity
			PrimarySmtpAddress    = $PrimarySmtpAddress
			MailboxLocation       = $null
			PropertyTag           = $propertyTag
			PropertyNamespace     = $propertySetId
			PropertyId            = $propertyId
			PropertyName          = $propertyName
			PropertyType          = $normalizedType
			PropertySetId         = $propertySetId
			PropertyTypeCode      = $null
			EwsMapiPropertyType   = $normalizedType
			Identity              = $definitionIdentity
			IsValid               = $true
			ObjectState           = 'Unchanged'
			DiscoverySource       = $DiscoverySource
			DiscoveryCompleteness = 'SuppliedDefinitions'
			ItemCount             = $null
			FolderCount           = $null
			FirstSeenFolderPath   = $null
			FirstSeenItemId       = $null
			Errors                = @()
		}
	}
}
