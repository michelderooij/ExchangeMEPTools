function Test-ExchangeManagementShell {
    [CmdletBinding()]
    param()

    $requiredCommands = $script:RequiredExchangeManagementShellCommands
    $missingCommands = @($requiredCommands | Where-Object { -not (Get-Command -Name $_ -ErrorAction SilentlyContinue) })
    if ($missingCommands.Count -gt 0) {
        throw "Exchange Management Shell is not loaded or is incomplete. Missing required command(s): $($missingCommands -join ', '). Run this command from the Exchange Management Shell."
    }

    if (-not (Get-Command -Name 'Get-StoreQuery' -ErrorAction SilentlyContinue)) {
        $exchangeInstallPath = Get-ExchangeInstallPath
        if ([string]::IsNullOrWhiteSpace($exchangeInstallPath)) {
            throw 'The Exchange Management Shell is loaded, but $exinstallpath is not set and ExchangeInstallPath is empty. Set $exinstallpath before connecting so the Store diagnostics script can be loaded.'
        }

        $diagnosticScript = Join-Path $exchangeInstallPath 'Scripts\ManagedStoreDiagnosticFunctions.ps1'
        if (-not (Test-Path -LiteralPath $diagnosticScript -PathType Leaf)) {
            throw "ManagedStoreDiagnosticFunctions.ps1 was not found at '$diagnosticScript'. Verify `$exinstallpath points to the Exchange installation directory."
        }

        try {
            . $diagnosticScript
        }
        catch {
            throw "ManagedStoreDiagnosticFunctions.ps1 could not be loaded. Confirm the Exchange Management Shell can run Get-ExchangeDiagnosticInfo: $($_.Exception.Message)"
        }
    }

    if (-not (Get-Command -Name 'Get-StoreQuery' -ErrorAction SilentlyContinue)) {
        throw 'ManagedStoreDiagnosticFunctions.ps1 loaded but Get-StoreQuery is still unavailable. Confirm the Exchange Management Shell can run Get-ExchangeDiagnosticInfo.'
    }

    [pscustomobject]@{
        ExchangeManagementShellLoaded = $true
        StoreQueryAvailable = $true
        ExchangeInstallPath = Get-ExchangeInstallPath
    }
}