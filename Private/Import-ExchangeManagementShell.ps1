function Import-ExchangeManagementShell {
    [CmdletBinding()]
    param()

    $missingCommands = @($script:RequiredExchangeManagementShellCommands | Where-Object {
            -not (Get-Command -Name $_ -ErrorAction SilentlyContinue)
        })
    if ($missingCommands.Count -eq 0) {
        return [pscustomobject]@{ Loaded = $true; Detected = $true; Connected = $false; MissingCommands = @() }
    }

    try {
        $installPath = Get-ExchangeInstallPath
    }
    catch {
        Write-Warning "ExchangeNamedProperties depends on the Exchange Management Shell. Exchange install path lookup failed: $($_.Exception.Message)"
        return [pscustomobject]@{ Loaded = $false; Detected = $false; Connected = $false; MissingCommands = $missingCommands }
    }
    if ([string]::IsNullOrWhiteSpace($installPath)) {
        Write-Warning 'ExchangeNamedProperties depends on the Exchange Management Shell. Exchange install path could not be determined.'
        return [pscustomobject]@{ Loaded = $false; Detected = $false; Connected = $false; MissingCommands = $missingCommands }
    }

    $remoteExchangeScript = Join-Path $installPath 'Bin\RemoteExchange.ps1'
    if (-not (Test-Path -LiteralPath $remoteExchangeScript -PathType Leaf)) {
        Write-Warning "ExchangeNamedProperties depends on the Exchange Management Shell. RemoteExchange.ps1 was not found at '$remoteExchangeScript'."
        return [pscustomobject]@{ Loaded = $false; Detected = $false; Connected = $false; MissingCommands = $missingCommands }
    }

    try {
        Write-Host 'Attempting to load the Exchange Management Shell ..'
        Set-StrictMode -Off
        . $remoteExchangeScript 6>$null
        $connectCommand = Get-Command -Name 'Connect-ExchangeServer' -ErrorAction SilentlyContinue
        if (-not $connectCommand) {
            throw 'RemoteExchange.ps1 did not define Connect-ExchangeServer.'
        }
        Connect-ExchangeServer -Auto | Out-Null
    }
    catch {
        Write-Warning "ExchangeNamedProperties depends on the Exchange Management Shell. EMS loading failed: $($_.Exception.Message)"
        return [pscustomobject]@{ Loaded = $false; Detected = $false; Connected = $false; MissingCommands = $missingCommands }
    }
    finally {
        Set-StrictMode -Version Latest
    }

    $missingCommands = @($script:RequiredExchangeManagementShellCommands | Where-Object {
            -not (Get-Command -Name $_ -ErrorAction SilentlyContinue)
        })
    if ($missingCommands.Count -gt 0) {
        Write-Warning "ExchangeNamedProperties depends on the Exchange Management Shell. Required cmdlet(s) are still unavailable: $($missingCommands -join ', ')."
        return [pscustomobject]@{ Loaded = $false; Detected = $false; Connected = $false; MissingCommands = $missingCommands }
    }

    [pscustomobject]@{ Loaded = $true; Detected = $false; Connected = $true; MissingCommands = @() }
}

function Get-ExchangeNamedPropertiesBanner {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Version
    )

    @(
        '*' * 78
        "ExchangeNamedProperties version $Version"
        '*' * 78
    ) -join [Environment]::NewLine
}