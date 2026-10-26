function Invoke-EwsRequestWithRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,

        [ValidateRange(0, 10)]
        [int]$RetryCount = 3,

        [ValidateRange(0, 300)]
        [int]$InitialBackoffMilliseconds = 500,

        [scriptblock]$DelayScriptBlock = {
            param([int]$DurationMilliseconds)
            Start-Sleep -Milliseconds $DurationMilliseconds
        }
    )

    for ($attempt = 0; ; $attempt++) {
        try {
            return ,(& $ScriptBlock)
        }
        catch {
            $exception = $_.Exception
            $message = $exception.Message
            $isTransient = $message -match '(?i)ErrorServerBusy|ServerBusy|throttl|back.?off|temporar(il)?y unavailable'
            if (-not $isTransient -or $attempt -ge $RetryCount) {
                if ($isTransient -and $attempt -ge $RetryCount) {
                    throw "EWS request remained busy after $($attempt + 1) attempt(s): $message"
                }
                throw
            }

            $backoffMilliseconds = $InitialBackoffMilliseconds * [math]::Pow(2, $attempt)
            $serverBackoff = $null
            if ($exception.PSObject.Properties['BackOffMilliseconds']) {
                $serverBackoff = [int]$exception.BackOffMilliseconds
            }
            if ($exception.Data -and $exception.Data.Contains('BackOffMilliseconds')) {
                $serverBackoff = [int]$exception.Data['BackOffMilliseconds']
            }
            if ($null -ne $serverBackoff -and $serverBackoff -gt 0) {
                $backoffMilliseconds = $serverBackoff
            }

            & $DelayScriptBlock ([int][math]::Min($backoffMilliseconds, 300000))
        }
    }
}