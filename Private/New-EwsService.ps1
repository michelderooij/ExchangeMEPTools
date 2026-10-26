function New-EwsService {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Mailbox,

        [Parameter(Mandatory = $true)]
        [psobject]$Options
    )

    Import-EwsManagedApi -Path $Options.EwsManagedApiPath | Out-Null

    if ($Options.TrustAll -and -not $Options.EwsUrl) {
        throw 'TrustAll requires an explicit HTTPS EwsUrl so certificate bypass can be scoped to one authority.'
    }

    if ($Options.TrustAll) {
        $trustUri = $null
        if (-not [uri]::TryCreate($Options.EwsUrl, [UriKind]::Absolute, [ref]$trustUri) -or
            $trustUri.Scheme -ne [Uri]::UriSchemeHttps -or [string]::IsNullOrEmpty($trustUri.Host)) {
            throw 'TrustAll requires an absolute HTTPS EwsUrl.'
        }
        Enable-EwsTrustAll -EwsUri $trustUri
    }

    try {
        $exchangeVersionType = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]
        $exchangeVersion = [Enum]::Parse($exchangeVersionType, $Options.ExchangeVersion, $true)
        if (-not [Enum]::IsDefined($exchangeVersionType, $exchangeVersion)) {
            throw "Unsupported ExchangeVersion '$($Options.ExchangeVersion)'."
        }
    }
    catch {
        throw "ExchangeVersion '$($Options.ExchangeVersion)' is not supported by the loaded EWS Managed API. Use a version exposed by Microsoft.Exchange.WebServices.Data.ExchangeVersion."
    }

    try {
        $service = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ExchangeService -ArgumentList $exchangeVersion
    }
    catch {
        throw "The EWS Managed API could not create an ExchangeService for version '$($Options.ExchangeVersion)': $($_.Exception.Message)"
    }

    $service.Timeout = [int]$Options.RequestTimeoutSeconds * 1000
    if ($Options.Credential) {
        $networkCredential = $Options.Credential.GetNetworkCredential()
        $ewsCredential = New-Object -TypeName Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList $networkCredential
        $service.Credentials = $ewsCredential
        $service.UseDefaultCredentials = $false
    }
    else {
        $service.UseDefaultCredentials = $true
    }

    if ($Options.Impersonate) {
        $service.ImpersonatedUserId = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ImpersonatedUserId -ArgumentList @(
            [Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,
            $Mailbox.PrimarySmtpAddress
        )
    }

    if ($Options.EwsUrl) {
        $ewsUri = $null
        if (-not [uri]::TryCreate($Options.EwsUrl, [UriKind]::Absolute, [ref]$ewsUri) -or
            $ewsUri.Scheme -ne [Uri]::UriSchemeHttps -or [string]::IsNullOrEmpty($ewsUri.Host) -or
            -not [string]::IsNullOrEmpty($ewsUri.UserInfo)) {
            throw 'EwsUrl must be an absolute HTTPS URL without embedded user information.'
        }

        $service.Url = $ewsUri
    }
    else {
        $redirectCallback = [Microsoft.Exchange.WebServices.Autodiscover.AutodiscoverRedirectionUrlValidationCallback] {
            param([string]$RedirectionUrl)
            Test-EwsHttpsRedirect -RedirectionUrl $RedirectionUrl
        }

        try {
            $service.AutodiscoverUrl($Mailbox.PrimarySmtpAddress, $redirectCallback)
        }
        catch {
            $message = "No explicit EWS URL was supplied or configured, and Autodiscover could not locate EWS for mailbox '$($Mailbox.PrimarySmtpAddress)'. If Autodiscover is not configured in your environment, run Connect-EWSSession -EwsUrl 'https://your-exchange-server/EWS/Exchange.asmx' -UseDefaultCredentials, then retry the command. Alternatively, supply -EwsUrl directly to Search-MailboxExtendedProperty or Remove-MailboxExtendedProperty. Import-Module -Force clears saved connection settings; reconnect after reloading the module. If you expect Autodiscover to work, check DNS and EWS availability. Original error: $($_.Exception.Message)"
            throw [System.InvalidOperationException]::new($message, $_.Exception)
        }

        if (-not $service.Url -or $service.Url.Scheme -ne [Uri]::UriSchemeHttps) {
            throw 'Autodiscover did not produce an HTTPS EWS endpoint. The connection was rejected.'
        }
    }

    if ($Options.TraceEnabled -and $Options.TraceFilePath) {
        $traceDirectory = Split-Path -Parent $Options.TraceFilePath
        if ($traceDirectory -and -not (Test-Path -LiteralPath $traceDirectory -PathType Container)) {
            throw "TraceFilePath directory '$traceDirectory' does not exist."
        }

        Add-Content -LiteralPath $Options.TraceFilePath -Value ("{0} EWS service configured for {1}; SOAP payload tracing is disabled." -f [DateTime]::UtcNow.ToString('o'), $Mailbox.PrimarySmtpAddress) -Encoding UTF8
    }

    return $service
}