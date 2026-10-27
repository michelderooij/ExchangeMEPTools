function Enable-EwsTrustAll {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [uri]$EwsUri
    )

    if ($EwsUri.Scheme -ne [Uri]::UriSchemeHttps -or [string]::IsNullOrEmpty($EwsUri.Host)) {
        throw 'TrustAll can be used only with an explicit HTTPS EwsUrl.'
    }

    $trustState = $script:EwsTrustPolicyState
    if (-not $trustState) {
        throw 'TrustAll certificate policy state is unavailable. Reload the ExchangeMEPTools module.'
    }

    [System.Threading.Monitor]::Enter($trustState.SyncRoot)
    try {
        if (-not $trustState.IsInstalled) {
            $trustState.PreviousCallback = [System.Net.ServicePointManager]::ServerCertificateValidationCallback
            $callbackScript = {
                param($sender, $certificate, $chain, $errors)

                $request = $sender -as [System.Net.HttpWebRequest]
                [System.Threading.Monitor]::Enter($trustState.SyncRoot)
                try {
                    if ($request -and $trustState.Authorities.Contains($request.RequestUri.Authority)) {
                        return $true
                    }
                    $previousCallback = $trustState.PreviousCallback
                }
                finally {
                    [System.Threading.Monitor]::Exit($trustState.SyncRoot)
                }

                if ($previousCallback) {
                    return $previousCallback.Invoke($sender, $certificate, $chain, $errors)
                }
                return $errors -eq [System.Net.Security.SslPolicyErrors]::None
            }.GetNewClosure()
            $trustState.TrustCallback = [System.Net.Security.RemoteCertificateValidationCallback]$callbackScript
            [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $trustState.TrustCallback
            $trustState.IsInstalled = $true
        }

        $trustState.Authorities.Add($EwsUri.Authority) | Out-Null
    }
    finally {
        [System.Threading.Monitor]::Exit($trustState.SyncRoot)
    }
}

function Disable-EwsTrustAll {
    [CmdletBinding()]
    param()

    $trustState = $script:EwsTrustPolicyState
    if (-not $trustState) {
        return
    }

    [System.Threading.Monitor]::Enter($trustState.SyncRoot)
    try {
        if ($trustState.IsInstalled -and [object]::ReferenceEquals(
                [System.Net.ServicePointManager]::ServerCertificateValidationCallback,
                $trustState.TrustCallback
            )) {
            [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $trustState.PreviousCallback
        }

        $trustState.Authorities.Clear()
        $trustState.PreviousCallback = $null
        $trustState.TrustCallback = $null
        $trustState.IsInstalled = $false
    }
    finally {
        [System.Threading.Monitor]::Exit($trustState.SyncRoot)
    }
}