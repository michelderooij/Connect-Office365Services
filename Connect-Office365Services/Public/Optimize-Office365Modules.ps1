function Optimize-Office365Modules {
    param (
        [switch]$AllowPrerelease
    )

    $local:Functions = Get-Office365ModuleInfo
    if ($PSBoundParameters.ContainsKey('AllowPrerelease')) {
        $script:myOffice365Services['AllowPrerelease'] = $AllowPrerelease.IsPresent
    }

    $local:IsAdmin = Test-IsAdministrator
    if ( $local:IsAdmin) {
        if ( (Get-Process -Name powershell, pwsh -ErrorAction SilentlyContinue | Measure-Object).Count -gt 1) {
            Write-Warning ('Running multiple PowerShell sessions, successful cleanup might be problematic.')
        }
        $local:AllInstalled = Get-Module -ListAvailable -ErrorAction SilentlyContinue
        foreach ( $local:Item in $local:Functions) {

            $local:Module = $local:AllInstalled | Where-Object { $_.Name -eq $local:Item.Module } | Sort-Object -Property Version -Descending

            if ( $local:Module) {
                Write-Host ('Checking {0} .. ' -f $local:Item.Description) -NoNewline

                $local:ModuleVersions = Get-myModule -Name $local:Item.Module -ListAvailable -All -ErrorAction SilentlyContinue
                $local:LatestModule = $local:ModuleVersions | Sort-Object -Property @{e = { [System.Version]($_.Version -replace '[^\d\.]', '') } } -Descending | Select-Object -First 1
                $local:LatestVersion = ($local:LatestModule).Version

                $local:OldModules = $local:ModuleVersions | Where-Object { $_.Version -ne $local:LatestVersion }
                if ( $local:OldModules) {

                    Write-Host ('previous versions found')

                    foreach ( $OldModule in $local:OldModules) {

                        # Uninstall all old versions of the module
                        $local:OldFullVer = Get-ModuleVersionInfo -Module $OldModule
                        Write-Host ('Uninstalling {0} v{1}' -f $OldModule.Name, $local:OldFullVer) -ForegroundColor White
                        try {
                            Uninstall-myModule -Name $OldModule.Name -Version $local:OldFullVer -IsPrerelease:($local:OldFullVer -match '-')
                        }
                        catch {
                            Write-Error ('Problem uninstalling {0} v{1}: {2}' -f $OldModule.Name, $local:OldFullVer, $Error[0].Exception.Message)
                        }
                    }
                }
                else {
                    Write-Host ('OK') -ForegroundColor Green
                }

                # Cleanup required modules as well
                $local:RequiredModules = $local:Module.RequiredModules | Sort-Object -Unique Name

                foreach ( $RequiredModule in $local:RequiredModules) {

                    Write-Host ('Checking {0} .. ' -f $RequiredModule.Name) -NoNewline

                    $local:ModuleVersions = Get-myModule -Name $RequiredModule.Name -ListAvailable -ErrorAction SilentlyContinue
                    $local:LatestModule = $local:ModuleVersions | Sort-Object -Property @{e = { [System.Version]($_.Version -replace '[^\d\.]', '') } } -Descending | Select-Object -First 1
                    $local:LatestVersion = ($local:LatestModule).Version

                    $local:OldModules = $local:ModuleVersions | Where-Object { $_.Version -ne $local:LatestVersion }
                    if ( $local:OldModules) {

                        Write-Host ('needs cleanup')

                        foreach ( $OldModule in $local:OldModules) {

                            $local:OldFullVer = Get-ModuleVersionInfo -Module $OldModule
                            Write-Host ('Uninstalling {0} v{1}' -f $OldModule.Name, $local:OldFullVer)
                            try {
                                Uninstall-myModule -Name $OldModule.Name -Version $local:OldFullVer -IsPrerelease:($local:OldFullVer -match '-')
                            }
                            catch {
                                Write-Error ('Problem uninstalling {0} v{1}: {2}' -f $OldModule.Name, $local:OldFullVer, $Error[0].Exception.Message)
                            }
                        }
                    }
                    else {
                        Write-Host ('OK') -ForegroundColor Green
                    }
                }
            }
        }

        # Final sweep: hard-delete any remaining old version folders from PSModulePath
        Write-Verbose ('Performing sweep for leftover module folders')
        $local:SweptCount = 0
        foreach ($local:Item in $local:Functions) {
            $local:ModuleBasePaths = Get-ModuleInstallPath -Name $local:Item.Module
            foreach ($local:BasePath in $local:ModuleBasePaths) {
                $local:VersionFolders = Get-ChildItem -Path $local:BasePath -Directory -ErrorAction SilentlyContinue
                if ($local:VersionFolders -and $local:VersionFolders.Count -gt 1) {
                    $local:LatestVersionFolder = $local:VersionFolders | Sort-Object -Property {
                        try { [System.Version]($_.Name -replace '[^\d\.]', '') } catch { [System.Version]'0.0' }
                    } -Descending | Select-Object -First 1
                    $local:OldFolders = $local:VersionFolders | Where-Object { $_.FullName -ne $local:LatestVersionFolder.FullName }
                    foreach ($local:OldFolder in $local:OldFolders) {
                        Write-Host ('Removing leftover folder {0} .. ' -f $local:OldFolder.FullName) -NoNewline
                        try {
                            Remove-Item -Path $local:OldFolder.FullName -Recurse -Force -ErrorAction Stop
                            Write-Host ('OK') -ForegroundColor Green
                            $local:SweptCount++
                        }
                        catch {
                            Write-Warning ('Failed to remove {0}: {1}' -f $local:OldFolder.FullName, $_.Exception.Message)
                        }
                    }
                }
            }
        }
        Write-Host ('Sweep complete, {0} leftover folders removed.' -f $local:SweptCount)
    }
    else {
        Write-Warning ('Script not running with elevated privileges; cannot remove modules')
    }
}
