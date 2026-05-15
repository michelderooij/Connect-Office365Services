function Get-Office365ServicesPreferences {
    <#
    .SYNOPSIS
    Displays the current user preferences for Connect-Office365Services.

    .DESCRIPTION
    Returns all persistent preference values together with the location of the
    preferences file and whether that file currently exists on disk.

    To change preferences use Set-Office365ServicesPreferences.

    .EXAMPLE
    Get-Office365ServicesPreferences
    Returns the current preference values and the preferences file location.
    #>
    [CmdletBinding()]
    param()

    $local:configPath = Join-Path -Path ([System.Environment]::GetFolderPath(
        [System.Environment+SpecialFolder]::ApplicationData)) -ChildPath 'Office365Services\config.json'

    [PSCustomObject][ordered]@{
        AllowPrerelease  = [bool]$script:myOffice365Services['AllowPrerelease']
        AzureEnvironment = [string]$script:myOffice365Services['AzureEnvironmentName']
        Scope            = [string]$script:myOffice365Services['Scope']
        ProxyAccessType  = [string]$script:myOffice365Services['ProxyAccessType']
        NoBanner         = [bool]$script:myOffice365Services['NoBanner']
        NoQuote          = [bool]$script:myOffice365Services['NoQuote']
        NoReport         = [bool]$script:myOffice365Services['NoReport']
        NoAutoConnect    = [bool]$script:myOffice365Services['NoAutoConnect']
        PreferencesFile  = $local:configPath
        FileExists       = (Test-Path -Path $local:configPath -PathType Leaf)
    }
}
