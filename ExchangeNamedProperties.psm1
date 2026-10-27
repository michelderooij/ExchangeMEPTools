param(
    [hashtable]$ImportOptions = @{}
)

Set-StrictMode -Version Latest

$script:RequiredExchangeManagementShellCommands = @(
    'Get-Mailbox'
    'Get-MailboxDatabase'
    'Get-ExchangeDiagnosticInfo'
)
$script:ModuleVersion = (Import-PowerShellDataFile -Path (Join-Path $PSScriptRoot 'ExchangeNamedProperties.psd1')).ModuleVersion

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

Set-Alias -Name Connect-ExchangeNamedProperties -Value Connect-EWSSession -Scope Script
Set-Alias -Name Disconnect-ExchangeNamedProperties -Value Disconnect-EWSSession -Scope Script

Export-ModuleMember -Function @(
    'Connect-EWSSession'
    'Disconnect-EWSSession'
    'Get-EWSSession'
    'Get-MailboxExtendedProperty'
    'Search-MailboxExtendedProperty'
    'Test-MailboxExtendedProperty'
    'Remove-MailboxExtendedProperty'
) -Alias @(
    'Connect-ExchangeNamedProperties'
    'Disconnect-ExchangeNamedProperties'
)

$skipExchangeShell = $ImportOptions -and $ImportOptions.ContainsKey('SkipExchangeShell') -and [bool]$ImportOptions['SkipExchangeShell']
Write-Host (Get-ExchangeNamedPropertiesBanner -Version $script:ModuleVersion)
if (-not $skipExchangeShell) {
    $exchangeShellStatus = Import-ExchangeManagementShell
    if ($exchangeShellStatus.Detected) {
        Write-Host 'Exchange Management Shell cmdlets detected.'
    }
}