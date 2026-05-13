function Import-Office365ModuleConfig {
    <#
    .SYNOPSIS
    Imports a preferences file exported by Export-Office365ModuleConfig.

    .DESCRIPTION
    Copies the specified JSON file to %APPDATA%\Office365Services\config.json,
    replacing the current preferences file. The active module session state is
    updated immediately without requiring a module reload.

    .PARAMETER File
    The path to the JSON file to import.

    .EXAMPLE
    Import-Office365ModuleConfig -File C:\Backup\O365Config.json
    Imports preferences and module state from the specified file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$File
    )

    if (-not (Test-Path -Path $File -PathType Leaf)) {
        Write-Error ('File not found: ''{0}''' -f $File)
        return
    }

    # Validate JSON before touching the live config
    try {
        $null = Get-Content -Path $File -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        Write-Error ('File is not valid JSON: {0}' -f $_)
        return
    }

    $local:configDir  = Join-Path -Path ([System.Environment]::GetFolderPath(
        [System.Environment+SpecialFolder]::ApplicationData)) -ChildPath 'Office365Services'
    $local:configPath = Join-Path -Path $local:configDir -ChildPath 'config.json'

    if (-not (Test-Path -Path $local:configDir -PathType Container)) {
        $null = New-Item -Path $local:configDir -ItemType Directory -Force
    }

    try {
        Copy-Item -Path $File -Destination $local:configPath -Force -ErrorAction Stop
        Write-Host ('Config imported from ''{0}''' -f $File)
    }
    catch {
        Write-Error ('Failed to copy config file: {0}' -f $_)
        return
    }

    # Reload preference keys into live session state
    $local:prefs = Read-Office365ServicesPreferences
    $script:myOffice365Services['AllowPrerelease'] = [bool]$local:prefs['AllowPrerelease']
    $script:myOffice365Services['Scope']           = [string]$local:prefs['Scope']
    $script:myOffice365Services['ProxyAccessType'] = [string]$local:prefs['ProxyAccessType']
    $script:myOffice365Services['SessionOptions']  = New-PSSessionOption -ProxyAccessType $local:prefs['ProxyAccessType']
    $script:myOffice365Services['NoBanner']        = [bool]$local:prefs['NoBanner']
    $script:myOffice365Services['NoQuote']         = [bool]$local:prefs['NoQuote']
    Set-Office365Environment -Environment $local:prefs['AzureEnvironment']
    Write-Verbose ('Session state updated from imported config.')
}
