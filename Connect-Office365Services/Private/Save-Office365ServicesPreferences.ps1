function Save-Office365ServicesPreferences {
    <#
    .SYNOPSIS
    Persists current user preferences to %APPDATA%\Office365Services\config.json.
    Called automatically by Set-Office365ServicesPreferences when any value changes.
    #>
    $local:configDir  = Join-Path -Path ([System.Environment]::GetFolderPath(
        [System.Environment+SpecialFolder]::ApplicationData)) -ChildPath 'Office365Services'
    $local:configPath = Join-Path -Path $local:configDir -ChildPath 'config.json'

    if (-not (Test-Path -Path $local:configDir -PathType Container)) {
        $null = New-Item -Path $local:configDir -ItemType Directory -Force
    }

    $local:prefs = [ordered]@{
        AllowPrerelease  = [bool]$script:myOffice365Services['AllowPrerelease']
        AzureEnvironment = [string]$script:myOffice365Services['AzureEnvironmentName']
        Scope            = [string]$script:myOffice365Services['Scope']
        ProxyAccessType  = [string]$script:myOffice365Services['ProxyAccessType']
    }

    $local:prefs | ConvertTo-Json | Set-Content -Path $local:configPath -Encoding UTF8 -Force
    Write-Verbose ('Preferences saved to ''{0}''' -f $local:configPath)
}
