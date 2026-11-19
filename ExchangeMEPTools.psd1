@{
    RootModule        = 'ExchangeMEPTools.psm1'
    ModuleVersion     = '0.9.0'
    GUID              = '2f72888d-bb2e-4b6e-9b3d-44f3fd1d7c93'
    Author            = 'Michel de Rooij'
    CompanyName       = 'EighTwOne'
    Copyright         = '(c) Michel de Rooij. All rights reserved.'
    Description       = 'Analyze, search, assess, and remove catalogued MAPI named properties from Exchange Server on-premises mailboxes through EWS.'
    PowerShellVersion = '5.1'
    FunctionsToExport = @(
        'Connect-EWSSession'
        'Disconnect-EWSSession'
        'Get-EWSSession'
        'Get-MailboxExtendedProperty'
        'Search-MailboxExtendedProperty'
        'Test-MailboxExtendedProperty'
        'Remove-MailboxExtendedProperty'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
    PrivateData       = @{
        PSData = @{
            Tags         = @('Exchange', 'EWS', 'MAPI', 'NamedProperties', 'ExtendedProperties')
            LicenseUri   = 'https://opensource.org/license/mit'
            ProjectUri   = 'https://github.com/'
            ReleaseNotes = 'Version 0.9.0; see CHANGELOG.md.'
        }
    }
}