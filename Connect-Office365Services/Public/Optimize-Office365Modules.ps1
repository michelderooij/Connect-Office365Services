function Optimize-Office365Modules {
    param (
        [switch]$AllowPrerelease
    )

    $local:Functions= Get-Office365ModuleInfo
    if ($PSBoundParameters.ContainsKey('AllowPrerelease')) {
        $script:myOffice365Services['AllowPrerelease'] = $AllowPrerelease.IsPresent
    }

    $local:IsAdmin= Test-IsAdministrator
    If( $local:IsAdmin) {
        If( (Get-Process -Name powershell, pwsh -ErrorAction SilentlyContinue | Measure-Object).Count -gt 1) {
            Write-Warning ('Running multiple PowerShell sessions, successful cleanup might be problematic.')
        }
        $local:AllInstalled = Get-Module -ListAvailable -ErrorAction SilentlyContinue
        ForEach ( $local:Item in $local:Functions) {

            $local:Module= $local:AllInstalled | Where-Object { $_.Name -eq $local:Item.Module } | Sort-Object -Property Version -Descending

            If( $local:Module) {
                Write-Host ('Checking {0} .. ' -f $local:Item.Description) -NoNewline

                $local:ModuleVersions= Get-myModule -Name $local:Item.Module -ListAvailable -All  -ErrorAction SilentlyContinue
                $local:LatestModule = $local:ModuleVersions | Sort-Object -Property @{e={ [System.Version]($_.Version -replace '[^\d\.]','')}} -Descending | Select-Object -First 1
                $local:LatestVersion = ($local:LatestModule).Version

                $local:OldModules= $local:ModuleVersions | Where-Object {$_.Version -ne $local:LatestVersion}
                If( $local:OldModules) {

                    Write-Host ('previous versions found')

                    ForEach( $OldModule in $local:OldModules) {

                        # Uninstall all old versions of the module
                        Write-Host ('Uninstalling {0} v{1}' -f $OldModule.Name, $OldModule.Version) -ForegroundColor White
                        Try {
                            Uninstall-myModule -Name $OldModule.Name -Version $OldModule.Version -IsPrerelease:$OldModule.IsPrerelease
                        }
                        Catch {
                            Write-Error ('Problem uninstalling {0} v{1}: {2}' -f $OldModule.Name, $OldModule.Version, $Error[0].Exception.Message)
                        }
                    }
                }
                Else {
                    Write-Host ('OK') -ForegroundColor Green
                }

                # Cleanup required modules as well
                $local:RequiredModules= $local:Module.RequiredModules | Sort-Object -Unique Name

                ForEach( $RequiredModule in $local:RequiredModules) {

                    Write-Host ('Checking {0} .. ' -f $RequiredModule.Name) -NoNewline

                    $local:ModuleVersions= Get-myModule -Name $RequiredModule.Name -ListAvailable -ErrorAction SilentlyContinue
                    $local:LatestModule = $local:ModuleVersions | Sort-Object -Property @{e={ [System.Version]($_.Version -replace '[^\d\.]','')}} -Descending | Select-Object -First 1
                    $local:LatestVersion = ($local:LatestModule).Version

                    $local:OldModules= $local:ModuleVersions | Where-Object {$_.Version -ne $local:LatestVersion}
                    If( $local:OldModules) {

                        Write-Host ('needs cleanup')

                        ForEach( $OldModule in $local:OldModules) {

                            Write-Host ('Uninstalling {0} v{1}' -f $OldModule.Name, $OldModule.Version)
                            Try {
                                Uninstall-myModule -Name $OldModule.Name -Version $OldModule.Version -IsPrerelease:$OldModule.IsPrerelease
                            }
                            Catch {
                                Write-Error ('Problem uninstalling {0} v{1}: {2}' -f $OldModule.Name, $OldModule.Version, $Error[0].Exception.Message)
                            }
                        }
                    }
                    Else {
                        Write-Host ('OK') -ForegroundColor Green
                    }
                }
            }
        }

        # Final sweep: hard-delete any remaining old version folders from PSModulePath
        Write-Host ('')
        Write-Host ('Performing final sweep for leftover module folders...')
        $local:SweptCount = 0
        ForEach ($local:Item in $local:Functions) {
            $local:ModuleBasePaths = Get-ModuleInstallPath -Name $local:Item.Module
            ForEach ($local:BasePath in $local:ModuleBasePaths) {
                $local:VersionFolders = Get-ChildItem -Path $local:BasePath -Directory -ErrorAction SilentlyContinue
                If ($local:VersionFolders -and $local:VersionFolders.Count -gt 1) {
                    $local:LatestVersionFolder = $local:VersionFolders | Sort-Object -Property {
                        try { [System.Version]($_.Name -replace '[^\d\.]', '') } catch { [System.Version]'0.0' }
                    } -Descending | Select-Object -First 1
                    $local:OldFolders = $local:VersionFolders | Where-Object { $_.FullName -ne $local:LatestVersionFolder.FullName }
                    ForEach ($local:OldFolder in $local:OldFolders) {
                        Write-Host ('Hard-deleting leftover folder: {0}' -f $local:OldFolder.FullName) -ForegroundColor Yellow
                        Try {
                            Remove-Item -Path $local:OldFolder.FullName -Recurse -Force -ErrorAction Stop
                            Write-Host ('  Deleted successfully') -ForegroundColor Green
                            $local:SweptCount++
                        }
                        Catch {
                            Write-Warning ('  Failed to delete {0}: {1}' -f $local:OldFolder.FullName, $_.Exception.Message)
                        }
                    }
                }
            }
        }
        If ($local:SweptCount -eq 0) {
            Write-Host ('No leftover folders found.') -ForegroundColor Green
        }
        Else {
            Write-Host ('Swept {0} leftover folder(s).' -f $local:SweptCount) -ForegroundColor Cyan
        }
    }
    Else {
        Write-Warning ('Script not running with elevated privileges; cannot remove modules')
    }
}
