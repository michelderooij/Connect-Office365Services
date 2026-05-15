@{
    # Module identity
    RootModule        = 'Connect-Office365Services.psm1'
    ModuleVersion     = '4.0.5'
    GUID              = '6b44e9b9-f24f-4b27-bce0-bfef01a75d31'
    Author            = 'Michel de Rooij'
    CompanyName       = 'EighTwOne'
    Copyright         = '(c) Michel de Rooij. All rights reserved.'
    Description       = 'Helper functions to connect to Microsoft 365 / Office 365 services and manage their PowerShell modules.'

    # Compatibility
    PowerShellVersion = '5.1'

    # Functions exported from this module
    FunctionsToExport = @(
        # Connect / disconnect functions
        'Connect-EXO',
        'Connect-Exchange',
        'Connect-SCC',
        'Connect-MSTeams',
        'Connect-AIP',
        'Connect-SPO',
        'Connect-PowerApps',
        'Connect-MG',
        'Connect-PowerBI',
        'Connect-PnP',
        'Connect-Office365',
        'Disconnect-Office365',

        # Credential / Identity functions
        'Get-Office365Credential',
        'Get-OnPremisesCredential',
        'Get-Office365Tenant',
        'Get-ExchangeOnPremisesFQDN',
        'Get-TenantID',

        # Session / connectivity
        'Get-Office365Session',
        'Test-Office365Connectivity',

        # Environment / preferences
        'Set-Office365Environment',
        'Get-Office365Services',
        'Get-Office365ServicesPreferences',
        'Set-Office365ServicesPreferences',

        # Module management
        'Select-Office365Modules',
        'Update-Office365Modules',
        'Optimize-Office365Modules',
        'Show-Office365Modules',
        'Save-Office365ModuleState',
        'Restore-Office365ModuleState',
        'Export-Office365ModuleConfig',
        'Import-Office365ModuleConfig'
    )

    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()

    # Module metadata for PSGallery
    PrivateData       = @{
        PSData = @{
            Tags         = @('Office365', 'Microsoft365', 'ExchangeOnline', 'AzureAD', 'SharePoint', 'Teams', 'MicrosoftTeams', 'PowerApps', 'Connect', 'M365')
            LicenseUri   = 'https://github.com/michelderooij/Connect-Office365Services/blob/main/LICENSE.md'
            ProjectUri   = 'https://github.com/michelderooij/Connect-Office365Services'
            ReleaseNotes = 'https://github.com/michelderooij/Connect-Office365Services/blob/main/CHANGELOG.md'
        }
    }
}
