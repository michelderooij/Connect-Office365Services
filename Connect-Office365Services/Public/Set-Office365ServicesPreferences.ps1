function Set-Office365ServicesPreferences {
    <#
    .SYNOPSIS
    View or set persistent user preferences for Connect-Office365Services.

    .DESCRIPTION
    When called without parameters, displays the current preference values.
    When one or more parameters are supplied, the values are updated in the
    active module state and written to:
        %APPDATA%\Office365Services\config.json

    The config file is read on every module import. If the file does not yet
    exist, built-in defaults are used. The file is created the first time any
    preference is changed.

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

    .EXAMPLE
    Set-Office365ServicesPreferences
    Displays the current preference values.

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
        [string]$ProxyAccessType
    )

    # ── No parameters: display current preferences ────────────────────────────
    if ($PSBoundParameters.Count -eq 0) {
        [PSCustomObject][ordered]@{
            AllowPrerelease  = [bool]$script:myOffice365Services['AllowPrerelease']
            AzureEnvironment = [string]$script:myOffice365Services['AzureEnvironmentName']
            Scope            = [string]$script:myOffice365Services['Scope']
            ProxyAccessType  = [string]$script:myOffice365Services['ProxyAccessType']
        }
        return
    }

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

    # ── Persist to config.json when anything changed ──────────────────────────
    if ($local:changed) {
        Save-Office365ServicesPreferences
    }
}
