function Uninstall-myModule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string]$Name,
        [Parameter(Mandatory=$true)]
        $Version,
        [switch]$IsPrerelease
    )
    Process {
        # Unload module from current session before attempting uninstall
        Remove-Module -Name $Name -Force -ErrorAction SilentlyContinue

        $local:MaxRetries   = 10
        $local:Attempt      = 0
        $local:CmdletSuccess= $false
        $local:FatalError   = $false

        While( $local:Attempt -lt $local:MaxRetries -and -not $local:CmdletSuccess -and -not $local:FatalError) {
            $local:Attempt++
            Try {
                If( $script:myOffice365Services['PSResourceGet']) {
                    Uninstall-PSResource -Name $Name -Version ([string]$Version) -Scope $script:myOffice365Services['Scope'] -SkipDependencyCheck -Prerelease:$IsPrerelease -ErrorAction Stop
                }
                Else {
                    Uninstall-Module -Name $Name -RequiredVersion ([string]$Version) -AllowPrerelease:$IsPrerelease -Force -ErrorAction Stop
                }
                $local:CmdletSuccess= $true
            }
            Catch {
                Switch -Regex ($PSItem.FullyQualifiedErrorId) {
                    '^AdminPrivilegesRequiredForUninstall,' {
                        Write-Warning ('Unable to uninstall {0} v{1}: Administrator rights required' -f $Name, $Version)
                        $local:FatalError= $true
                    }
                    '^(UnableToUninstallAsOtherModulesNeedThisModule|UninstallPSResourcePackageIsaDependency),' {
                        Write-Warning ('Unable to uninstall {0} v{1}: other modules depend on it' -f $Name, $Version)
                        $local:FatalError= $true
                    }
                    Default {
                        If( $local:Attempt -ge $local:MaxRetries) {
                            Write-Warning ('Problem uninstalling {0} v{1}: {2}' -f $Name, $Version, $PSItem.Exception.Message)
                        }
                    }
                }
                If( -not $local:FatalError -and -not $local:CmdletSuccess -and $local:Attempt -lt $local:MaxRetries) {
                    Start-Sleep -Seconds 1
                }
            }
        }

        # Hard-delete fallback: remove remaining version folder(s) from PSModulePath
        If( -not $local:CmdletSuccess -and -not $local:FatalError) {
            $local:VersionStr = [string]$Version
            $local:RemainingPaths = Get-ModuleInstallPath -Name $Name -Version $local:VersionStr

            # Also try without prerelease suffix in case folder uses only the base version number
            If( -not $local:RemainingPaths) {
                $local:BaseVersion = $local:VersionStr -replace '\-.*$', ''
                If( $local:BaseVersion -ne $local:VersionStr) {
                    $local:RemainingPaths = Get-ModuleInstallPath -Name $Name -Version $local:BaseVersion
                }
            }

            ForEach( $local:ModPath in $local:RemainingPaths) {
                Try {
                    Remove-Item -Path $local:ModPath -Recurse -Force -ErrorAction Stop
                    Write-Verbose ('Hard-deleted module folder: {0}' -f $local:ModPath)
                }
                Catch {
                    Write-Warning ('Failed to hard-delete module folder {0}: {1}' -f $local:ModPath, $_.Exception.Message)
                }
            }
        }
    }
}
