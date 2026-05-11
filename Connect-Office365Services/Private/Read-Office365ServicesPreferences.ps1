function Read-Office365ServicesPreferences {
    <#
    .SYNOPSIS
    Reads user preferences from %APPDATA%\Office365Services\config.json.
    Returns a hashtable with merged defaults + persisted values.
    Missing keys in the file fall back to defaults silently.
    #>
    $local:defaults = [ordered]@{
        AllowPrerelease  = $false
        AzureEnvironment = 'Default'
        Scope            = 'AllUsers'
        ProxyAccessType  = 'None'
    }

    $local:configPath = Join-Path -Path ([System.Environment]::GetFolderPath(
        [System.Environment+SpecialFolder]::ApplicationData)) -ChildPath 'Office365Services\config.json'

    if (Test-Path -Path $local:configPath -PathType Leaf) {
        try {
            $local:json = Get-Content -Path $local:configPath -Raw -ErrorAction Stop |
                ConvertFrom-Json -ErrorAction Stop
            foreach ($local:key in @($local:defaults.Keys)) {
                $local:prop = $local:json.PSObject.Properties[$local:key]
                if ($null -ne $local:prop) {
                    $local:defaults[$local:key] = $local:prop.Value
                }
            }
        }
        catch {
            Write-Warning ('Failed to read preferences from ''{0}'': {1}' -f $local:configPath, $_)
        }
    }

    return $local:defaults
}
