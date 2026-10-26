function Get-NamedPropertyLimits {
    [CmdletBinding()]
    param(
        [string]$Path = (Join-Path $PSScriptRoot '..\Data\NamedPropertyLimits.json')
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return [pscustomobject]@{
            Source = 'Unknown'
            Limits = @()
        }
    }

    try {
        $configuration = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        throw "Named-property limits file '$Path' could not be read: $($_.Exception.Message)"
    }

    [pscustomobject]@{
        ExchangeServerSource                 = if ($configuration.ExchangeServer -and $configuration.ExchangeServer.Source) { [string]$configuration.ExchangeServer.Source } else { 'Unknown' }
        ExchangeServerLimits                 = if ($configuration.ExchangeServer -and $configuration.ExchangeServer.Limits) { @($configuration.ExchangeServer.Limits) } else { @() }
        ExchangeOnlineReferenceSource        = if ($configuration.ExchangeOnlineReference -and $configuration.ExchangeOnlineReference.Source) { [string]$configuration.ExchangeOnlineReference.Source } else { 'Unknown' }
        ExchangeOnlineReferenceUrl           = if ($configuration.ExchangeOnlineReference -and $configuration.ExchangeOnlineReference.Url) { [string]$configuration.ExchangeOnlineReference.Url } else { $null }
        ExchangeOnlineReferenceApplicability = if ($configuration.ExchangeOnlineReference -and $configuration.ExchangeOnlineReference.Applicability) { [string]$configuration.ExchangeOnlineReference.Applicability } else { 'Unknown' }
        ExchangeOnlineReferenceLimits        = if ($configuration.ExchangeOnlineReference -and $configuration.ExchangeOnlineReference.Limits) { @($configuration.ExchangeOnlineReference.Limits) } else { @() }
    }
}