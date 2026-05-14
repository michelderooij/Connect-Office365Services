function Initialize-ServicesState {
    <#
    .SYNOPSIS
    Initializes the module-scoped $script:myOffice365Services state hashtable and default settings.
    Called automatically when the module is imported. Reads persisted preferences from
    %APPDATA%\Office365Services\config.json when present; falls back to built-in defaults.
    #>

    # Initialize module state hashtable when not already present
    if ($null -eq $script:myOffice365Services) {
        $script:myOffice365Services = @{}
    }

    # Load preferences (config.json when present, otherwise built-in defaults)
    $local:prefs = Read-Office365ServicesPreferences

    # Prerelease preference
    $script:myOffice365Services['AllowPrerelease'] = [bool]$local:prefs['AllowPrerelease']

    # Module install scope
    $script:myOffice365Services['Scope'] = [string]$local:prefs['Scope']

    # Proxy / session options
    $script:myOffice365Services['ProxyAccessType'] = [string]$local:prefs['ProxyAccessType']
    $script:myOffice365Services['SessionOptions']  = New-PSSessionOption -ProxyAccessType $local:prefs['ProxyAccessType']

    # Banner / quote suppression
    $script:myOffice365Services['NoBanner'] = [bool]$local:prefs['NoBanner']
    $script:myOffice365Services['NoQuote']  = [bool]$local:prefs['NoQuote']

    # Modern auth state (populated by Get-Office365Credential / Get-Office365AccessToken)
    $script:myOffice365Services['Office365UPN']   = ''
    $script:myOffice365Services['MsalAccount']    = $null
    $script:myOffice365Services['MsalApp']             = $null   # PublicClientApplication instance — Graph Command Line Tools (MSAL.NET)
    $script:myOffice365Services['MsalTokenCacheBytes'] = $null   # serialized MSAL token cache — persists Graph tokens across PCA rebuilds
    $script:myOffice365Services['MsalNetWarned']       = $false  # suppress repeated "MSAL.NET not found" warnings
    # Well-known Microsoft Graph Command Line Tools public client — override via Set-Office365ServicesPreferences if needed
    $script:myOffice365Services['MsalClientId']   = '14d82eec-204b-4c2f-b7e8-296a70dab67e'

    # Initialize environment & endpoints
    Set-Office365Environment -Environment $local:prefs['AzureEnvironment']

    # In-session cache for online version lookups (populated by Show/Update; 60-min TTL)
    # Key: module name (string), Value: PSCustomObject { Version; Fetched }
    if ($null -eq $script:myOffice365Services['OnlineVersionCache']) {
        $script:myOffice365Services['OnlineVersionCache'] = @{}
    }
}
