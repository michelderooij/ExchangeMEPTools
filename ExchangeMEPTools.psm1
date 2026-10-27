param(
    [hashtable]$ImportOptions = @{}
)

Set-StrictMode -Version Latest

$script:RequiredExchangeManagementShellCommands = @(
    'Get-Mailbox'
    'Get-MailboxDatabase'
    'Get-ExchangeDiagnosticInfo'
)
$script:ModuleVersion = (Import-PowerShellDataFile -Path (Join-Path $PSScriptRoot 'ExchangeMEPTools.psd1')).ModuleVersion

$script:ConnectionOptions = $null
$script:EwsServiceCache = @{}
$script:EwsTrustPolicyState = [pscustomobject]@{
    SyncRoot         = New-Object object
    Authorities      = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    PreviousCallback = $null
    TrustCallback    = $null
    IsInstalled      = $false
}

Get-ChildItem -Path (Join-Path $PSScriptRoot 'Private') -Filter '*.ps1' -File |
Sort-Object -Property Name |
ForEach-Object { . $_.FullName }

Get-ChildItem -Path (Join-Path $PSScriptRoot 'Public') -Filter '*.ps1' -File |
Sort-Object -Property Name |
ForEach-Object { . $_.FullName }

Export-ModuleMember -Function @(
    'Connect-EWSSession'
    'Disconnect-EWSSession'
    'Get-EWSSession'
    'Get-MailboxExtendedProperty'
    'Search-MailboxExtendedProperty'
    'Test-MailboxExtendedProperty'
    'Remove-MailboxExtendedProperty'
) -Alias @()

$skipExchangeShell = $ImportOptions -and $ImportOptions.ContainsKey('SkipExchangeShell') -and [bool]$ImportOptions['SkipExchangeShell']

Write-Host ('***********************************************************************')
Write-Host ('Exchange Mailbox Extended Property Tools v{0}' -f $script:ModuleVersion)
Write-Host 'https://github.com/michelderooij/ExchangeMEPTools'
Write-Host '***********************************************************************'

if (-not $skipExchangeShell) {
    $exchangeShellStatus = Import-ExchangeManagementShell
    if ($exchangeShellStatus.Detected) {
        Write-Host 'Exchange Management Shell cmdlets detected.'
    }
}