function Invoke-StoreQuery {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Database,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Query,

        [hashtable]$QueryContext
    )

    $storeQueryCommand = Get-Command -Name 'Get-StoreQuery' -ErrorAction SilentlyContinue
    if (-not $storeQueryCommand) {
        $exchangeInstallPath = Get-ExchangeInstallPath
        if ([string]::IsNullOrWhiteSpace($exchangeInstallPath)) {
            throw 'Get-StoreQuery is unavailable. Run this provider in the Exchange Management Shell and ensure $exinstallpath is set.'
        }

        $diagnosticScript = Join-Path $exchangeInstallPath 'Scripts\ManagedStoreDiagnosticFunctions.ps1'
        if (-not (Test-Path -LiteralPath $diagnosticScript -PathType Leaf)) {
            throw "Get-StoreQuery is unavailable and ManagedStoreDiagnosticFunctions.ps1 was not found at '$diagnosticScript'. Load the supported Exchange diagnostics script or run this command in the Exchange Management Shell."
        }

        try {
            . $diagnosticScript
        }
        catch {
            throw "ManagedStoreDiagnosticFunctions.ps1 could not be loaded: $($_.Exception.Message)"
        }

        $storeQueryCommand = Get-Command -Name 'Get-StoreQuery' -ErrorAction SilentlyContinue
        if (-not $storeQueryCommand) {
            throw 'ManagedStoreDiagnosticFunctions.ps1 loaded but did not define Get-StoreQuery.'
        }
    }

    $useWorkerContext = $null -ne $QueryContext -and
        $storeQueryCommand.Parameters.ContainsKey('Server') -and
        $storeQueryCommand.Parameters.ContainsKey('ProcessId')
    for ($attempt = 0; $attempt -lt 2; $attempt++) {
        try {
            $arguments = @{ Query = $Query; Unlimited = $true; ErrorAction = 'Stop' }
            if ($useWorkerContext) {
                if ($QueryContext['Database'] -ne $Database -or -not $QueryContext['Server'] -or -not $QueryContext['ProcessId']) {
                    $status = Get-MailboxDatabase -Identity $Database -Status -ErrorAction Stop
                    $workerProcessId = 0
                    if (-not $status.Mounted -or [string]::IsNullOrWhiteSpace([string]$status.MountedOnServer) -or
                        -not [int]::TryParse([string]$status.WorkerProcessId, [ref]$workerProcessId) -or $workerProcessId -le 0) {
                        throw "Database '$Database' has no accessible mounted Store worker."
                    }
                    $QueryContext['Database'] = $Database
                    $QueryContext['Server'] = [string]$status.MountedOnServer
                    $QueryContext['ProcessId'] = $workerProcessId
                }
                $arguments.Server = $QueryContext['Server']
                $arguments.ProcessId = $QueryContext['ProcessId']
            }
            else {
                $arguments.Database = $Database
            }
            & $storeQueryCommand @arguments
            return
        }
        catch {
            if ($useWorkerContext -and $attempt -eq 0 -and
                $_.Exception.Message -match '(?i)unable to bind to the RPC endpoint') {
                [void]$QueryContext.Remove('Database')
                [void]$QueryContext.Remove('Server')
                [void]$QueryContext.Remove('ProcessId')
                continue
            }
            throw "Get-StoreQuery failed for database '$Database': $($_.Exception.Message)"
        }
    }
}