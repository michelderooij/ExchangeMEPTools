function Test-NamedPropertyDefinition {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $true)]
		[psobject]$Definition
	)

	try {
		$normalized = ConvertTo-NormalizedNamedProperty -Definition $Definition
		[pscustomobject]@{
			IsValid = $true
			Definition = $normalized
			ErrorMessage = $null
		}
	}
	catch {
		[pscustomobject]@{
			IsValid = $false
			Definition = $null
			ErrorMessage = $_.Exception.Message
		}
	}
}
