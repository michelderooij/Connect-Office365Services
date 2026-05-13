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

    # Read-modify-write: preserve keys we don't manage (e.g. ModuleState)
    $local:existing = [ordered]@{}
    if (Test-Path -Path $local:configPath -PathType Leaf) {
        try {
            $local:json = Get-Content -Path $local:configPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            foreach ($local:prop in $local:json.PSObject.Properties) {
                $local:existing[$local:prop.Name] = $local:prop.Value
            }
        }
        catch {
            Write-Verbose ('Could not read existing config for merge: {0}' -f $_)
        }
    }

    $local:existing['AllowPrerelease']  = [bool]$script:myOffice365Services['AllowPrerelease']
    $local:existing['AzureEnvironment'] = [string]$script:myOffice365Services['AzureEnvironmentName']
    $local:existing['Scope']            = [string]$script:myOffice365Services['Scope']
    $local:existing['ProxyAccessType']  = [string]$script:myOffice365Services['ProxyAccessType']
    $local:existing['NoBanner']         = [bool]$script:myOffice365Services['NoBanner']
    $local:existing['NoQuote']          = [bool]$script:myOffice365Services['NoQuote']

    $local:existing | ConvertTo-Json -Depth 5 | Set-Content -Path $local:configPath -Encoding UTF8 -Force
    Write-Verbose ('Preferences saved to ''{0}''' -f $local:configPath)
}
