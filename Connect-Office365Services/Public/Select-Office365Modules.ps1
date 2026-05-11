function Select-Office365Modules {
    <#
    .SYNOPSIS
    Interactive module selection menu for Office 365 modules.

    .DESCRIPTION
    Provides a text-based interactive menu to select/deselect Office 365 modules.
    Navigate with Up/Down arrow keys, toggle selection with Space, commit with Enter.
    After selection, installs new modules and uninstalls deselected modules.

    .PARAMETER AllowPrerelease
    Allow installation of prerelease modules.

    .EXAMPLE
    Select-Office365Modules
    Opens interactive module selection menu.

    .EXAMPLE
    Select-Office365Modules -AllowPrerelease
    Opens interactive module selection menu with prerelease support.
    #>
    [CmdletBinding()]
    param(
        [switch]$AllowPrerelease
    )

    # Check if running as administrator
    $local:IsAdmin = Test-IsAdministrator
    if (-not $local:IsAdmin) {
        Write-Warning 'Script not running with elevated privileges; module installation/uninstallation may fail'
        $continue = Read-Host "Continue anyway? (y/N)"
        if ($continue -notmatch '^[Yy]') {
            return
        }
    }

    $script:myOffice365Services['AllowPrerelease'] = if ($PSBoundParameters.ContainsKey('AllowPrerelease')) { $AllowPrerelease.IsPresent } else { [bool]$script:myOffice365Services['AllowPrerelease'] }

    # Get module information
    $local:ModuleInfo = Get-Office365ModuleInfo
    $local:CurrentSelection = @{}
    $local:SelectedIndex = 0
    $local:MaxIndex = $local:ModuleInfo.Count - 1

    # Initialize current selection based on installed modules
    $local:AllInstalled = Get-Module -ListAvailable -ErrorAction SilentlyContinue
    foreach ($module in $local:ModuleInfo) {
        $installedModule = Get-InstalledRepoModule -Name $module.Module -Repo $module.Repo -AllInstalled $local:AllInstalled
        $local:CurrentSelection[$module.Module] = $null -ne $installedModule
    }

    # Display single-column menu
    function Show-ModuleMenu {
        param($ModuleInfo, $CurrentSelection, $SelectedIndex)

        Write-Host 'Module Selection'
        Write-Host ('-' * 50)
        Write-Host ('Active Scope: {0,-12}' -f $script:myOffice365Services['Scope'])
        Write-Host ''

        for ($i = 0; $i -lt $ModuleInfo.Count; $i++) {
            $module = $ModuleInfo[$i]
            $isSelected  = $CurrentSelection[$module.Module]
            $isReadOnly  = $module.ReplacedBy -and -not $isSelected
            $checkbox    = if ($isReadOnly) { ' - ' } elseif ($isSelected) { '[x]' } else { '[ ]' }
            $prefix    = if ($i -eq $SelectedIndex) { '>' } else { ' ' }
            $replacedBySuffix = if ($module.ReplacedBy) { ' (Replaced by: {0})' -f $module.ReplacedBy } else { '' }
            $line      = '{0} {1} {2}{3}' -f $prefix, $checkbox, $module.Description, $replacedBySuffix

            if ($i -eq $SelectedIndex) {
                Write-Host $line -ForegroundColor White
            }
            else {
                Write-Host $line
            }
        }

        $local:selectedCount = ($CurrentSelection.Values | Where-Object { $_ } | Measure-Object).Count
        Write-Host ''
        Write-Host ('Selected: {0}/{1}' -f $local:selectedCount, $ModuleInfo.Count)
        Write-Host ''
        Write-Host 'Up/Down: navigate  Space: toggle  S: scope  Enter: confirm  Esc: cancel'
    }

    # Compute menu line count for in-place redraw (9 fixed lines + 1 per module)
    $local:menuLineCount = 9 + $local:ModuleInfo.Count
    $local:menuTopRow    = -1

    # Main menu loop
    $exitMenu  = $false
    $committed = $false

    while (-not $exitMenu) {
        if ($local:menuTopRow -ge 0) {
            $Host.UI.RawUI.CursorPosition = [System.Management.Automation.Host.Coordinates]::new(0, $local:menuTopRow)
        }

        Show-ModuleMenu -ModuleInfo $local:ModuleInfo -CurrentSelection $local:CurrentSelection -SelectedIndex $local:SelectedIndex

        if ($local:menuTopRow -lt 0) {
            $local:menuTopRow = $Host.UI.RawUI.CursorPosition.Y - $local:menuLineCount
        }

        $key = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')

        switch ($key.VirtualKeyCode) {
            38 { # Up arrow
                $local:next = $local:SelectedIndex - 1
                while ($local:next -ge 0 -and $local:ModuleInfo[$local:next].ReplacedBy -and -not $local:CurrentSelection[$local:ModuleInfo[$local:next].Module]) {
                    $local:next--
                }
                if ($local:next -ge 0) { $local:SelectedIndex = $local:next }
            }
            40 { # Down arrow
                $local:next = $local:SelectedIndex + 1
                while ($local:next -le $local:MaxIndex -and $local:ModuleInfo[$local:next].ReplacedBy -and -not $local:CurrentSelection[$local:ModuleInfo[$local:next].Module]) {
                    $local:next++
                }
                if ($local:next -le $local:MaxIndex) { $local:SelectedIndex = $local:next }
            }
            32 { # Spacebar — toggle selection
                $currentModuleInfo = $local:ModuleInfo[$local:SelectedIndex]
                # Deprecated modules (ReplacedBy) can only be deselected when installed; cannot be installed fresh
                if ($currentModuleInfo.ReplacedBy -and -not $local:CurrentSelection[$currentModuleInfo.Module]) {
                    # Not installed and deprecated — do nothing
                }
                else {
                    $local:CurrentSelection[$currentModuleInfo.Module] = -not $local:CurrentSelection[$currentModuleInfo.Module]
                }
            }
            83 { # S — toggle scope
                if ($script:myOffice365Services['Scope'] -eq 'AllUsers') {
                    $script:myOffice365Services['Scope'] = 'CurrentUser'
                }
                else {
                    $script:myOffice365Services['Scope'] = 'AllUsers'
                }
                Save-Office365ServicesPreferences
            }
            13 { # Enter — commit
                $exitMenu  = $true
                $committed = $true
            }
            27 { # Escape — cancel
                $exitMenu  = $true
                $committed = $false
            }
        }
    }

    if (-not $committed) {
        Write-Host 'Operation cancelled.'
        return
    }

    # Process changes
    Write-Host 'Processing module changes...'
    Write-Host ''

    $modulesToInstall   = @()
    $modulesToUninstall = @()
    $local:AllInstalled = Get-Module -ListAvailable -ErrorAction SilentlyContinue

    foreach ($module in $local:ModuleInfo) {
        $moduleName          = $module.Module
        $shouldBeInstalled   = $local:CurrentSelection[$moduleName]

        $installedModule = Get-InstalledRepoModule -Name $moduleName -Repo $module.Repo -AllInstalled $local:AllInstalled
        $isCurrentlyInstalled = $null -ne $installedModule

        if ($shouldBeInstalled -and -not $isCurrentlyInstalled) {
            $modulesToInstall += $module
        }
        elseif (-not $shouldBeInstalled -and $isCurrentlyInstalled) {
            $modulesToUninstall += @{ Module = $module; InstalledModule = $installedModule }
        }
    }

    # Install new modules
    foreach ($module in $modulesToInstall) {
        Write-Host ('Installing {0}...' -f $module.Description)
        try {
            Install-myModule -Name $module.Module -AllowPrerelease:([bool]$script:myOffice365Services['AllowPrerelease']) -AllowClobber
            Write-Host ('  Installed {0}' -f $module.Module)
        }
        catch {
            Write-Error ('Failed to install {0}: {1}' -f $module.Module, $_.Exception.Message)
        }
    }

    # Uninstall removed modules
    foreach ($moduleInfo in $modulesToUninstall) {
        $module = $moduleInfo.Module
        Write-Host ('Uninstalling {0}...' -f $module.Description)
        try {
            $allVersions = Get-Module -Name $module.Module -ListAvailable |
                Where-Object { $_.RepositorySourceLocation -and ([System.Uri]($_.RepositorySourceLocation)).Authority -ieq ([System.Uri]($module.Repo)).Authority }
            foreach ($version in $allVersions) {
                Uninstall-myModule -Name $version.Name -Version $version.Version -IsPrerelease:$version.IsPrerelease
            }
            Write-Host ('  Uninstalled {0}' -f $module.Module)
        }
        catch {
            Write-Error ('Failed to uninstall {0}: {1}' -f $module.Module, $_.Exception.Message)
        }
    }

    if ($modulesToInstall.Count -eq 0 -and $modulesToUninstall.Count -eq 0) {
        Write-Host 'No changes were made.'
    }
    else {
        Write-Host ''
        Write-Host ('Done. Installed: {0}, Uninstalled: {1}' -f $modulesToInstall.Count, $modulesToUninstall.Count)
    }
}

