function Update-Office365Modules {
    param (
        [switch]$AllowPrerelease
    )

    $local:Functions= Get-Office365ModuleInfo
    if ($PSBoundParameters.ContainsKey('AllowPrerelease')) {
        $script:myOffice365Services['AllowPrerelease'] = $AllowPrerelease.IsPresent
    }

    $local:IsAdmin= [System.Security.principal.windowsprincipal]::new([System.Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
    If( $local:IsAdmin) {
        If( (Get-Process -Name powershell, pwsh -ErrorAction SilentlyContinue | Measure-Object).Count -gt 1) {
            Write-Warning ('Running multiple PowerShell sessions, successful updating might be problematic.')
        }
        ForEach ( $local:Item in $local:Functions) {

            $local:Module= Get-myModule -Name ('{0}' -f $local:Item.Module) | Sort-Object -Property Version -Descending | Select-Object -First 1

            If( ($local:Module).RepositorySourceLocation) {

                $local:Version = Get-ModuleVersionInfo -Module $local:Module
                Write-Host ('Checking {0}' -f $local:Item.Description) -NoNewLine

                $local:NewerAvailable= $false
                $OnlineModule = Find-myModule -Name $local:Item.Module -ErrorAction SilentlyContinue
                If( $OnlineModule) {
                    Write-Host (': Local:{0}, Online:{1}' -f $local:Version, $OnlineModule.version)
                    If( (Compare-TextVersionNumber -Version $local:Version -CompareTo $OnlineModule.version) -eq 1) {
                        $local:NewerAvailable= $true
                    }
                }
                Else {
                        # Not installed from online or cannot determine
                        Write-Host ('Local:{0} Online:N/A' -f $local:Version)
                }

                If( $local:NewerAvailable) {
                    $local:UpdateSuccess= $false
                    Try {
                        Update-myModule -Name $local:Item.Module
                        $local:UpdateSuccess= $true
                    }
                    Catch {
                        Write-Error ('Problem updating {0}:{1}' -f $local:Item.Module, $Error[0].Exception.Message)
                    }

                    If( $local:UpdateSuccess) {

                        $local:ModuleVersions= Get-myModule -Name $local:Item.Module -ListAvailable -All

                        $local:Module = $local:ModuleVersions | Sort-Object -Property @{e={ [System.Version]($_.Version -replace '[^\d\.]','')}} -Descending | Select-Object -First 1
                        $local:LatestVersion = ($local:Module).Version
                        Write-Host ('Updated {0} to version {1}' -f $local:Item.Description, $local:LatestVersion) -ForegroundColor Green

                        # Uninstall all old versions of module & dependencies
                        If( $OnlineModule) {
                            ForEach( $DependencyModule in $Module.Dependencies) {

                                $local:DepModuleVersions= Get-myModule -Name $DependencyModule.Name -ListAvailable
                                $local:DepModule = $local:DepModuleVersions | Sort-Object -Property @{e={ [System.Version]($_.Version -replace '[^\d\.]','')}} -Descending | Select-Object -First 1
                                $local:DepLatestVersion = ($local:DepModule).Version
                                $local:OldDepModules= $local:DepModuleVersions | Where-Object {$_.Version -ne $local:DepLatestVersion}
                                $local:OldDepModules | ForEach-Object {
                                    $DepModule= $_
                                    Write-Host ('Uninstalling dependency {0} version {1}' -f $DepModule.Name, $DepModule.Version)
                                    Try {
                                        Uninstall-myModule -Name $DepModule.Name -Version $DepModule.Version -IsPrerelease:$DepModule.IsPrerelease
                                    }
                                    Catch {
                                        Write-Error ('Problem uninstalling {0} version {1}' -f $DepModule.Name, $DepModule.Version)
                                    }
                                }
                            }
                            $local:OldModules= $local:ModuleVersions | Where-Object {$_.Version -ne $local:LatestVersion}
                            If( $local:OldModules) {
                                ForEach( $OldModule in $local:OldModules) {
                                    Write-Host ('Uninstalling {0} version {1}' -f $local:Item.Description, $OldModule.Version) -ForegroundColor White
                                    Try {
                                        Uninstall-myModule -Name $OldModule.Name -Version $OldModule.Version -IsPrerelease:$OldModule.IsPrerelease
                                    }
                                    Catch {
                                        Write-Error ('Problem uninstalling {0} version {1}' -f $OldModule.Name, $OldModule.Version)
                                    }
                                }
                            }
                        }
                    }
                    Else {
                        # Problem during update
                    }
                }
                Else {
                    # No update available
                }
            }
            Else {
                # Not installed
            }
        }
    }
    Else {
        Write-Warning ('Script not running with elevated privileges; cannot update modules')
    }
}
