function Set-Office365ServicesPreferences {
    <#
    .SYNOPSIS
    Set persistent user preferences for Connect-Office365Services.

    .DESCRIPTION
    Updates the supplied preference values in the active module state and
    writes them to:
        %APPDATA%\Office365Services\config.json

    The config file is read on every module import. If the file does not yet
    exist, built-in defaults are used. The file is created the first time any
    preference is changed. Use Get-Office365ServicesPreferences to view the
    current preference values.

    .PARAMETER AllowPrerelease
    Whether to include prerelease versions when finding, installing, or updating
    modules. Defaults to $false. Can still be overridden per-call by explicitly
    passing -AllowPrerelease to any module-management cmdlet.

    .PARAMETER AzureEnvironment
    The Microsoft 365 cloud environment to target. Valid values:
    Default (AzureCloud), Germany, China, AzurePPE, GCC, GCCHigh, DoD.
    Defaults to 'Default'.

    .PARAMETER Scope
    The PowerShell module installation scope used by Install/Update operations.
    Valid values: AllUsers, CurrentUser. Defaults to 'AllUsers'.

    .PARAMETER ProxyAccessType
    The proxy access type used when creating PSSessionOption for remoting.
    Valid values: None, WinHttpConfig, AutoDetect, IEConfig, NoProxyServer.
    Defaults to 'None'.

    .PARAMETER NoBanner
    When set to $true, suppresses the ASCII art banner shown during module import.
    Defaults to $false.

    .PARAMETER NoQuote
    When set to $true, suppresses the random quote shown during module import.
    Defaults to $false.

    .PARAMETER NoReport
    When set to $true, suppresses the list of found modules shown during module import.
    Defaults to $false.

    .PARAMETER NoAutoConnect
    When set to $true, prevents the connect functions from prompting for credentials
    automatically. Run Get-Office365Credential explicitly before connecting.
    Defaults to $false.

    .EXAMPLE
    Set-Office365ServicesPreferences -AllowPrerelease $true -Scope CurrentUser
    Enables prerelease modules and switches the install scope to CurrentUser.

    .EXAMPLE
    Set-Office365ServicesPreferences -AzureEnvironment GCCHigh
    Switches all service endpoints to the US Government GCC High cloud.
    #>
    [CmdletBinding()]
    param(
        [System.Nullable[bool]]$AllowPrerelease,

        [ValidateSet('Germany', 'China', 'AzurePPE', 'GCC', 'GCCHigh', 'DoD', 'Default')]
        [string]$AzureEnvironment,

        [ValidateSet('AllUsers', 'CurrentUser')]
        [string]$Scope,

        [ValidateSet('None', 'WinHttpConfig', 'AutoDetect', 'IEConfig', 'NoProxyServer')]
        [string]$ProxyAccessType,

        [System.Nullable[bool]]$NoBanner,

        [System.Nullable[bool]]$NoQuote,

        [System.Nullable[bool]]$NoReport,

        [System.Nullable[bool]]$NoAutoConnect
    )

    # ── Apply each supplied preference to module state ────────────────────────
    $local:changed = $false

    if ($PSBoundParameters.ContainsKey('AllowPrerelease')) {
        $script:myOffice365Services['AllowPrerelease'] = [bool]$AllowPrerelease
        $local:changed = $true
    }

    if ($PSBoundParameters.ContainsKey('AzureEnvironment')) {
        Set-Office365Environment -Environment $AzureEnvironment
        $local:changed = $true
    }

    if ($PSBoundParameters.ContainsKey('Scope')) {
        $script:myOffice365Services['Scope'] = $Scope
        $local:changed = $true
    }

    if ($PSBoundParameters.ContainsKey('ProxyAccessType')) {
        $script:myOffice365Services['ProxyAccessType'] = $ProxyAccessType
        $script:myOffice365Services['SessionOptions']  = New-PSSessionOption -ProxyAccessType $ProxyAccessType
        $local:changed = $true
    }

    if ($PSBoundParameters.ContainsKey('NoBanner')) {
        $script:myOffice365Services['NoBanner'] = [bool]$NoBanner
        $local:changed = $true
    }

    if ($PSBoundParameters.ContainsKey('NoQuote')) {
        $script:myOffice365Services['NoQuote'] = [bool]$NoQuote
        $local:changed = $true
    }

    if ($PSBoundParameters.ContainsKey('NoReport')) {
        $script:myOffice365Services['NoReport'] = [bool]$NoReport
        $local:changed = $true
    }

    if ($PSBoundParameters.ContainsKey('NoAutoConnect')) {
        $script:myOffice365Services['NoAutoConnect'] = [bool]$NoAutoConnect
        $local:changed = $true
    }

    # ── Persist to config.json when anything changed ──────────────────────────
    if ($local:changed) {
        Save-Office365ServicesPreferences
        Get-Office365ServicesPreferences
    }
}
