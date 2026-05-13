function Save-Office365ModuleState {
    <#
    .SYNOPSIS
    Saves the installed version of each supported Office 365 module to the preferences file.

    .DESCRIPTION
    Iterates the supported module list, finds any installed modules, records their name
    and version, and stores the result as a 'ModuleState' array in the preferences file
    (%APPDATA%\Office365Services\config.json).

    The saved state can be used by Restore-Office365ModuleState to reinstall the same
    module versions on another machine or after a clean OS installation.

    .EXAMPLE
    Save-Office365ModuleState
    Saves the currently installed module versions to the preferences file.
    #>
    [CmdletBinding()]
    param()

    $local:configDir  = Join-Path -Path ([System.Environment]::GetFolderPath(
        [System.Environment+SpecialFolder]::ApplicationData)) -ChildPath 'Office365Services'
    $local:configPath = Join-Path -Path $local:configDir -ChildPath 'config.json'

    if (-not (Test-Path -Path $local:configDir -PathType Container)) {
        $null = New-Item -Path $local:configDir -ItemType Directory -Force
    }

    $local:AllInstalled = Get-Module -ListAvailable -ErrorAction SilentlyContinue
    $local:Functions    = Get-Office365ModuleInfo
    $local:State        = [System.Collections.Generic.List[object]]::new()

    foreach ($local:Item in $local:Functions) {
        $local:Module = Get-InstalledRepoModule -Name $local:Item.Module -Repo $local:Item.Repo -AllInstalled $local:AllInstalled
        if ($local:Module) {
            $local:Version = Get-ModuleVersionInfo -Module $local:Module
            $local:State.Add([ordered]@{ Module = $local:Item.Module; Version = [string]$local:Version })
            Write-Host ('Saved {0} v{1}' -f $local:Item.Description, $local:Version)
        }
    }

    # Read-modify-write: preserve all other keys in config.json
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

    $local:existing['ModuleState'] = $local:State.ToArray()
    $local:existing | ConvertTo-Json -Depth 5 | Set-Content -Path $local:configPath -Encoding UTF8 -Force
    Write-Host ('{0}Module state saved ({1} module(s)) to ''{2}''' -f [System.Environment]::NewLine, $local:State.Count, $local:configPath)
}
