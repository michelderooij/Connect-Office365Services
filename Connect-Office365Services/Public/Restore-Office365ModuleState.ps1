function Restore-Office365ModuleState {
    <#
    .SYNOPSIS
    Installs supported Office 365 modules from a saved module state.

    .DESCRIPTION
    Reads the 'ModuleState' array from the preferences file and installs each module.
    By default the exact version stored in the state is installed.
    Use -Recent to install the latest available version instead.

    The module state is created by Save-Office365ModuleState.

    .PARAMETER Recent
    When specified, installs the most recent available version of each module rather
    than the pinned version stored in the state.

    .EXAMPLE
    Restore-Office365ModuleState
    Reinstalls each module at the exact version recorded by Save-Office365ModuleState.

    .EXAMPLE
    Restore-Office365ModuleState -Recent
    Installs the latest available version of each saved module.
    #>
    [CmdletBinding()]
    param(
        [switch]$Recent
    )

    $local:configPath = Join-Path -Path ([System.Environment]::GetFolderPath(
        [System.Environment+SpecialFolder]::ApplicationData)) -ChildPath 'Office365Services\config.json'

    if (-not (Test-Path -Path $local:configPath -PathType Leaf)) {
        Write-Warning ('No preferences file found at ''{0}''. Run Save-Office365ModuleState first.' -f $local:configPath)
        return
    }

    try {
        $local:json = Get-Content -Path $local:configPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        Write-Error ('Failed to read preferences file: {0}' -f $_)
        return
    }

    $local:ModuleState = $local:json.ModuleState
    if (-not $local:ModuleState) {
        Write-Warning ('No module state found in the preferences file. Run Save-Office365ModuleState first.')
        return
    }

    $local:Scope   = $script:myOffice365Services['Scope']
    $local:UsePSRG = $script:myOffice365Services['PSResourceGet']

    foreach ($local:Entry in $local:ModuleState) {
        $local:ModuleName    = $local:Entry.Module
        $local:PinnedVersion = $local:Entry.Version

        if ($Recent) {
            Write-Host ('Installing {0} (latest)...' -f $local:ModuleName) -NoNewline
            try {
                Install-myModule -Name $local:ModuleName
                Write-Host (' Done') -ForegroundColor Green
            }
            catch {
                Write-Host (' Failed') -ForegroundColor Red
                Write-Warning ('Could not install {0}: {1}' -f $local:ModuleName, $_.Exception.Message)
            }
        }
        else {
            Write-Host ('Installing {0} v{1}...' -f $local:ModuleName, $local:PinnedVersion) -NoNewline
            try {
                if ($local:UsePSRG) {
                    Install-PSResource -Name $local:ModuleName -Version $local:PinnedVersion -Scope $local:Scope -TrustRepository -ErrorAction Stop
                }
                else {
                    Install-Module -Name $local:ModuleName -RequiredVersion $local:PinnedVersion -Scope $local:Scope -Force -AllowClobber -ErrorAction Stop
                }
                Write-Host (' Done') -ForegroundColor Green
            }
            catch {
                Write-Host (' Failed') -ForegroundColor Red
                Write-Warning ('Could not install {0} v{1}: {2}' -f $local:ModuleName, $local:PinnedVersion, $_.Exception.Message)
            }
        }
    }
}
