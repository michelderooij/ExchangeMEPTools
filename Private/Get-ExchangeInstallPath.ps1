function Get-ExchangeInstallPath {
    [CmdletBinding()]
    param()

    $installPathVariable = Get-Variable -Name 'exinstallpath' -Scope Global -ErrorAction SilentlyContinue
    if ($installPathVariable -and -not [string]::IsNullOrWhiteSpace([string]$installPathVariable.Value)) {
        return [string]$installPathVariable.Value
    }

    if (-not [string]::IsNullOrWhiteSpace($env:ExchangeInstallPath)) {
        return [string]$env:ExchangeInstallPath
    }

    $setupKey = 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup'
    try {
        $setup = Get-ItemProperty -LiteralPath $setupKey -Name 'MsiInstallPath' -ErrorAction Stop
        if (-not [string]::IsNullOrWhiteSpace([string]$setup.MsiInstallPath)) {
            return [string]$setup.MsiInstallPath
        }
    }
    catch {
    }

    return $null
}