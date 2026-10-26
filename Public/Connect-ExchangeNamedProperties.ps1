function ConvertTo-EwsEndpointUri {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    $url = $Value.Trim()
    if ($url -notmatch '^(?i)https?://') {
        $url = $url.Replace('\', '/')
        if ($url -notmatch '/') {
            $url = '{0}/EWS/Exchange.asmx' -f $url
        }
        elseif ($url.EndsWith('/')) {
            $url = '{0}EWS/Exchange.asmx' -f $url
        }
        $url = 'https://{0}' -f $url
    }

    $uri = $null
    if (-not [uri]::TryCreate($url, [UriKind]::Absolute, [ref]$uri) -or
        $uri.Scheme -ne [Uri]::UriSchemeHttps -or [string]::IsNullOrEmpty($uri.Host) -or
        -not [string]::IsNullOrEmpty($uri.UserInfo)) {
        throw 'EwsUrl must be an HTTPS URL or a host/endpoint shorthand without embedded user information.'
    }

    $uri
}

function Connect-EWSSession {
    <#
    .SYNOPSIS
    Configures EWS connection settings for this module session.

    .DESCRIPTION
    Verifies that required Exchange Management Shell commands are available and loads ManagedStoreDiagnosticFunctions.ps1 from $exinstallpath when needed. Stores EWS connection defaults for the current PowerShell module session. By default, EWS uses the current Windows security context. A PSCredential can be supplied instead. If Identity is supplied, the function resolves the mailbox and validates EWS service setup.

    .PARAMETER Identity
    Optional mailbox SMTP address, alias, GUID, or other identity resolvable by Get-Mailbox.

    .PARAMETER EwsUrl
    Optional HTTPS EWS endpoint. A bare host defaults to /EWS/Exchange.asmx, and host/path input is prefixed with https://. If omitted, EWS Autodiscover is used.

    .PARAMETER EwsManagedApiPath
    Optional path to Microsoft.Exchange.WebServices.dll or its containing directory.

    .PARAMETER Credential
    Optional credential used with EWS WebCredentials.

    .PARAMETER UseDefaultCredentials
    Explicitly select the current Windows security context. This is the default when Credential is omitted.

    .PARAMETER ExchangeVersion
    EWS request schema version supported by the installed Managed API.

    .PARAMETER RequestTimeoutSeconds
    EWS request timeout in seconds.

    .PARAMETER RetryCount
    Default number of retries for transient EWS requests. Search and removal commands can override this per invocation.

    .PARAMETER TraceEnabled
    Enable safe connection trace logging. SOAP payloads and authorization headers are not written.

    .PARAMETER TraceFilePath
    Optional path for safe connection trace messages.

    .PARAMETER Impersonate
    Opt in to EWS mailbox impersonation for the resolved target mailbox.

    .PARAMETER TrustAll
    Opt in to bypass certificate validation for the explicit EWS URL authority only. This is insecure and intended only for controlled testing. It requires EwsUrl and remains active until Disconnect-EWSSession.

    .EXAMPLE
    Connect-EWSSession -EwsUrl $env:EXCHANGE_EWS_URL

    .EXAMPLE
    Connect-EWSSession -Identity user@contoso.com -UseDefaultCredentials
    #>

    [CmdletBinding(DefaultParameterSetName = 'CurrentUser')]
    param(
        [Parameter(ParameterSetName = 'CurrentUser')]
        [Parameter(ParameterSetName = 'Credential')]
        [string]$Identity,

        [Parameter(ParameterSetName = 'CurrentUser')]
        [Parameter(ParameterSetName = 'Credential')]
        [ValidateNotNullOrEmpty()]
        [string]$EwsUrl,

        [Parameter(ParameterSetName = 'CurrentUser')]
        [Parameter(ParameterSetName = 'Credential')]
        [ValidateNotNullOrEmpty()]
        [string]$EwsManagedApiPath,

        [Parameter(Mandatory = $true, ParameterSetName = 'Credential')]
        [System.Management.Automation.PSCredential]$Credential,

        [Parameter(ParameterSetName = 'CurrentUser')]
        [switch]$UseDefaultCredentials,

        [Parameter(ParameterSetName = 'CurrentUser')]
        [Parameter(ParameterSetName = 'Credential')]
        [ValidateSet('Exchange2007_SP1', 'Exchange2010', 'Exchange2010_SP1', 'Exchange2010_SP2', 'Exchange2013', 'Exchange2013_SP1')]
        [string]$ExchangeVersion = 'Exchange2010',

        [Parameter(ParameterSetName = 'CurrentUser')]
        [Parameter(ParameterSetName = 'Credential')]
        [ValidateRange(1, 600)]
        [int]$RequestTimeoutSeconds = 100,

        [Parameter(ParameterSetName = 'CurrentUser')]
        [Parameter(ParameterSetName = 'Credential')]
        [ValidateRange(0, 10)]
        [int]$RetryCount = 3,

        [Parameter(ParameterSetName = 'CurrentUser')]
        [Parameter(ParameterSetName = 'Credential')]
        [switch]$TraceEnabled,

        [Parameter(ParameterSetName = 'CurrentUser')]
        [Parameter(ParameterSetName = 'Credential')]
        [ValidateNotNullOrEmpty()]
        [string]$TraceFilePath,

        [Parameter(ParameterSetName = 'CurrentUser')]
        [Parameter(ParameterSetName = 'Credential')]
        [switch]$Impersonate

        ,
        [Parameter(ParameterSetName = 'CurrentUser')]
        [Parameter(ParameterSetName = 'Credential')]
        [switch]$TrustAll
    )

    if ($EwsUrl) {
        $configuredUri = ConvertTo-EwsEndpointUri -Value $EwsUrl
        $EwsUrl = $configuredUri.AbsoluteUri
    }

    if ($TrustAll -and -not $EwsUrl) {
        throw 'TrustAll requires an explicit HTTPS EwsUrl; it is not supported with Autodiscover.'
    }

    $exchangeShellStatus = Test-ExchangeManagementShell

    if ($TraceFilePath -and -not $TraceEnabled) {
        throw 'TraceFilePath requires TraceEnabled.'
    }

    if ($Credential -and $UseDefaultCredentials) {
        throw 'Specify either Credential or UseDefaultCredentials, not both.'
    }

    $authenticationMethod = if ($Credential) { 'PSCredential' } else { 'DefaultWindowsCredentials' }
    $authenticationIdentity = if ($Credential) {
        $Credential.UserName
    }
    else {
        [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    }

    $options = [pscustomobject]@{
        EwsUrl                        = $EwsUrl
        EwsManagedApiPath             = $EwsManagedApiPath
        Credential                    = $Credential
        UseDefaultCredentials         = -not [bool]$Credential
        AuthenticationMethod         = $authenticationMethod
        AuthenticationIdentity       = $authenticationIdentity
        Identity                      = $Identity
        ExchangeVersion               = $ExchangeVersion
        RequestTimeoutSeconds         = $RequestTimeoutSeconds
        RetryCount                    = $RetryCount
        TraceEnabled                  = [bool]$TraceEnabled
        TraceFilePath                 = $TraceFilePath
        Impersonate                   = [bool]$Impersonate
        TrustAll                       = [bool]$TrustAll
        ExchangeManagementShellLoaded = $exchangeShellStatus.ExchangeManagementShellLoaded
    }

    Disable-EwsTrustAll
    Import-EwsManagedApi -Path $EwsManagedApiPath | Out-Null

    if ($TrustAll) {
        Write-Warning "TrustAll is enabled for the explicit EWS authority '$($configuredUri.Authority)'. Certificate mismatches will be accepted for that authority until Disconnect-EWSSession. Use only for controlled testing."
        Enable-EwsTrustAll -EwsUri $configuredUri
    }

    $script:ConnectionOptions = $options
    $script:EwsServiceCache = @{}

    if ($Identity) {
        $mailbox = Resolve-MailboxIdentity -Identity $Identity
        $service = New-EwsService -Mailbox $mailbox -Options $options
        $script:EwsServiceCache[$mailbox.PrimarySmtpAddress.ToLowerInvariant()] = $service
    }

    [pscustomobject]@{
        PSTypeName           = 'ExchangeNamedProperties.EWSSession'
        Connected            = $true
        Identity             = if ($Identity) { $Identity } else { $null }
        PrimarySmtpAddress   = if ($Identity) { $mailbox.PrimarySmtpAddress } else { $null }
        EwsUrl               = if ($Identity -and $service) { $service.Url.AbsoluteUri } else { $EwsUrl }
        Authentication       = $authenticationMethod
        AuthenticationIdentity = $authenticationIdentity
        ImpersonationEnabled = [bool]$Impersonate
        TrustAll             = [bool]$TrustAll
    }
}