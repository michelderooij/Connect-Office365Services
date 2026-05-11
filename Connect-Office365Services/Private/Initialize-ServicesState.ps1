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

    # Initialize environment & endpoints
    Set-Office365Environment -Environment $local:prefs['AzureEnvironment']
}
