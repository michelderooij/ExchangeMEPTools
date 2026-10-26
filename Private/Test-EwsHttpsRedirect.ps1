function Test-EwsHttpsRedirect {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RedirectionUrl
    )

    $uri = $null
    if (-not [uri]::TryCreate($RedirectionUrl, [UriKind]::Absolute, [ref]$uri)) {
        return $false
    }

    return $uri.Scheme -eq [Uri]::UriSchemeHttps -and [string]::IsNullOrEmpty($uri.UserInfo)
}