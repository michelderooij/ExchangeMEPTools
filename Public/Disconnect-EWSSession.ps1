function Disconnect-EWSSession {
    <#
    .SYNOPSIS
    Clears EWS connection settings and cached services for this module session.

    .DESCRIPTION
    Removes module-session connection configuration and references to cached EWS services. No credentials or settings are persisted by this module.

    .EXAMPLE
    Disconnect-EWSSession
    #>
    [CmdletBinding(SupportsShouldProcess = $false)]
    param()

    Disable-EwsTrustAll
    $script:ConnectionOptions = $null
    $script:EwsServiceCache = @{}
}