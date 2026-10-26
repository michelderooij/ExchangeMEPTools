function Get-EWSSession {
    <#
    .SYNOPSIS
    Reports the current module-scoped EWS connection configuration.

    .DESCRIPTION
    Returns connection defaults and live cached EWS service references. Session contains the service when exactly one is cached; otherwise use CachedServices to select a mailbox. Live service objects include authentication configuration and must not be serialized or logged as session diagnostics.

    .PARAMETER Identity
    Select a cached service by mailbox SMTP address, alias, or GUID. Does not create a service if none is cached for the mailbox.

    .EXAMPLE
    Get-EWSSession

    .EXAMPLE
    Get-EWSSession -Identity user@contoso.com
    #>
    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$Identity
    )

    $options = $script:ConnectionOptions
    $selectedCacheKey = $null
    if ($PSBoundParameters.ContainsKey('Identity')) {
        $selectedCacheKey = $Identity
        if (-not $script:EwsServiceCache.ContainsKey($selectedCacheKey)) {
            $mailbox = Resolve-MailboxIdentity -Identity $Identity
            $selectedCacheKey = $mailbox.PrimarySmtpAddress
        }
        if (-not $script:EwsServiceCache.ContainsKey($selectedCacheKey)) {
            throw "No cached EWS session exists for mailbox '$Identity'. Connect-EWSSession with -Identity first."
        }
    }
    $activeServices = @(
        foreach ($entry in $script:EwsServiceCache.GetEnumerator()) {
            if ($selectedCacheKey -and $entry.Key -ne $selectedCacheKey) {
                continue
            }
            $serviceUrl = $null
            if ($entry.Value -and $entry.Value.Url) {
                $serviceUrl = $entry.Value.Url.AbsoluteUri
            }
            [pscustomobject]@{
                MailboxIdentity = $entry.Key
                EwsUrl          = $serviceUrl
                Session         = $entry.Value
            }
        }
    )

    $authenticationIdentity = if ($options) { $options.AuthenticationIdentity } else { $null }
    $impersonationEnabled = if ($options) { [bool]$options.Impersonate } else { $false }
    $impersonationIdentity = if ($impersonationEnabled) {
        if ($selectedCacheKey) { $selectedCacheKey } elseif ($options.Identity) { $options.Identity } else { @($activeServices | Select-Object -ExpandProperty MailboxIdentity) }
    }
    else {
        $null
    }

    [pscustomobject]@{
        PSTypeName             = 'ExchangeNamedProperties.EWSSession'
        Connected              = [bool]$options
        EwsUrl                 = if ($selectedCacheKey) { $activeServices[0].EwsUrl } elseif ($options) { $options.EwsUrl } else { $null }
        AuthenticationMethod   = if ($options) { $options.AuthenticationMethod } else { $null }
        AuthenticationIdentity = $authenticationIdentity
        ConfiguredIdentity     = if ($options) { $options.Identity } else { $null }
        ImpersonationEnabled   = $impersonationEnabled
        ImpersonationIdentity  = $impersonationIdentity
        RetryCount             = if ($options) { $options.RetryCount } else { 3 }
        RetryCountScope        = 'Default for Search and Remove; a per-command -RetryCount overrides it.'
        TraceEnabled           = if ($options) { [bool]$options.TraceEnabled } else { $false }
        TraceFilePath          = if ($options) { $options.TraceFilePath } else { $null }
        TrustAll               = if ($options) { [bool]$options.TrustAll } else { $false }
        ExchangeVersion        = if ($options) { $options.ExchangeVersion } else { $null }
        RequestTimeoutSeconds  = if ($options) { $options.RequestTimeoutSeconds } else { $null }
        CachedServiceCount     = $activeServices.Count
        CachedServices         = $activeServices
        Session                = if ($activeServices.Count -eq 1) { $activeServices[0].Session } else { $null }
    }
}