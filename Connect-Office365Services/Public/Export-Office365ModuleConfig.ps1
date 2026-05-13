function Export-Office365ModuleConfig {
    <#
    .SYNOPSIS
    Exports the preferences file to a JSON file.

    .DESCRIPTION
    Saves the current module state (via Save-Office365ModuleState), then copies
    the contents of %APPDATA%\Office365Services\config.json to the specified file
    path, formatted as readable JSON.

    Use Import-Office365ModuleConfig to import the file on another machine.

    .PARAMETER File
    The path to the destination JSON file.

    .EXAMPLE
    Export-Office365ModuleConfig -File C:\Backup\O365Config.json
    Exports the current preferences and module state to a file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$File
    )

    $local:configPath = Join-Path -Path ([System.Environment]::GetFolderPath(
            [System.Environment+SpecialFolder]::ApplicationData)) -ChildPath 'Office365Services\config.json'

    # Always snapshot latest module state before exporting
    Save-Office365ModuleState

    if (-not (Test-Path -Path $local:configPath -PathType Leaf)) {
        Write-Warning ('Preferences file not found at ''{0}''.' -f $local:configPath)
        return
    }

    try {
        $local:content = Get-Content -Path $local:configPath -Raw -ErrorAction Stop |
        ConvertFrom-Json -ErrorAction Stop |
        ConvertTo-Json -Depth 5
        Set-Content -Path $File -Value $local:content -Encoding UTF8 -Force -ErrorAction Stop
        Write-Host ('Config exported to ''{0}''' -f $File)
    }
    catch {
        Write-Error ('Failed to export config: {0}' -f $_)
    }
}
