function ConvertFrom-EwsExtendedProperty {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return $null
    }
    if ($Value -is [byte[]]) {
        return [Convert]::ToBase64String($Value)
    }
    if ($Value -is [DateTime]) {
        return $Value.ToUniversalTime().ToString('o', [Globalization.CultureInfo]::InvariantCulture)
    }
    if ($Value -is [guid]) {
        return $Value.ToString('D').ToLowerInvariant()
    }
    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        $convertedValues = New-Object 'System.Collections.Generic.List[object]'
        foreach ($entry in $Value) {
            $convertedValues.Add((ConvertFrom-EwsExtendedProperty -Value $entry))
        }
        return ,$convertedValues.ToArray()
    }

    return $Value
}