function Get-EwsService {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity,

        [string]$EwsUrl,

        [string]$EwsManagedApiPath,

        [System.Management.Automation.PSCredential]$Credential,

        [switch]$UseDefaultCredentials,

        [string]$ExchangeVersion,

        [ValidateRange(1, 600)]
        [int]$RequestTimeoutSeconds,

        [switch]$TraceEnabled,

        [string]$TraceFilePath,

        [switch]$Impersonate,

        [switch]$TrustAll
    )

    if (-not $script:ConnectionOptions -and [string]::IsNullOrWhiteSpace($EwsUrl)) {
        Write-Warning 'No EWS session is configured. Run Connect-EWSSession before searching or removing mailbox properties, or supply -EwsUrl to the command. Autodiscover will still be attempted.'
    }

    $mailbox = Resolve-MailboxIdentity -Identity $Identity
    $configuredOptions = $script:ConnectionOptions
    if (-not $configuredOptions) {
        $configuredOptions = [pscustomobject]@{
            EwsUrl                = $null
            EwsManagedApiPath     = $null
            Credential            = $null
            ExchangeVersion       = 'Exchange2010'
            RequestTimeoutSeconds = 100
            TraceEnabled          = $false
            TraceFilePath         = $null
            Impersonate           = $false
            TrustAll              = $false
        }
    }

    $options = [pscustomobject]@{
        EwsUrl                = if ($PSBoundParameters.ContainsKey('EwsUrl')) { $EwsUrl } else { $configuredOptions.EwsUrl }
        EwsManagedApiPath     = if ($PSBoundParameters.ContainsKey('EwsManagedApiPath')) { $EwsManagedApiPath } else { $configuredOptions.EwsManagedApiPath }
        Credential            = if ($UseDefaultCredentials) { $null } elseif ($PSBoundParameters.ContainsKey('Credential')) { $Credential } else { $configuredOptions.Credential }
        UseDefaultCredentials = $true
        ExchangeVersion       = if ($PSBoundParameters.ContainsKey('ExchangeVersion')) { $ExchangeVersion } else { $configuredOptions.ExchangeVersion }
        RequestTimeoutSeconds = if ($PSBoundParameters.ContainsKey('RequestTimeoutSeconds')) { $RequestTimeoutSeconds } else { $configuredOptions.RequestTimeoutSeconds }
        TraceEnabled          = if ($PSBoundParameters.ContainsKey('TraceEnabled')) { [bool]$TraceEnabled } else { $configuredOptions.TraceEnabled }
        TraceFilePath         = if ($PSBoundParameters.ContainsKey('TraceFilePath')) { $TraceFilePath } else { $configuredOptions.TraceFilePath }
        Impersonate           = if ($PSBoundParameters.ContainsKey('Impersonate')) { [bool]$Impersonate } else { $configuredOptions.Impersonate }
        TrustAll              = if ($PSBoundParameters.ContainsKey('TrustAll')) { [bool]$TrustAll } else { [bool]$configuredOptions.TrustAll }
    }
    $options.UseDefaultCredentials = -not [bool]$options.Credential

    if ($options.TraceFilePath -and -not $options.TraceEnabled) {
        throw 'TraceFilePath requires TraceEnabled.'
    }

    $isOverride = $PSBoundParameters.Count -gt 1
    $cacheKey = $mailbox.PrimarySmtpAddress.ToLowerInvariant()
    if (-not $isOverride -and $script:EwsServiceCache.ContainsKey($cacheKey)) {
        return $script:EwsServiceCache[$cacheKey]
    }

    $service = New-EwsService -Mailbox $mailbox -Options $options
    if (-not $isOverride) {
        $script:EwsServiceCache[$cacheKey] = $service
    }

    return $service
}