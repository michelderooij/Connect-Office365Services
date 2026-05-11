function Update-Office365Modules {
    param (
        [switch]$AllowPrerelease,
        [switch]$Refresh
    )

    $local:Functions= Get-Office365ModuleInfo
    $local:UsePre = if ($AllowPrerelease.IsPresent) { $true } else { [bool]$script:myOffice365Services['AllowPrerelease'] }
    if ($Refresh) { $script:myOffice365Services['OnlineVersionCache'].Clear() }

    $local:IsAdmin= Test-IsAdministrator
    If( $local:IsAdmin) {
        If( (Get-Process -Name powershell, pwsh -ErrorAction SilentlyContinue | Measure-Object).Count -gt 1) {
            Write-Host ('Running multiple PowerShell sessions, successful updating might be problematic.') -ForegroundColor $script:myConsoleColors.Warning
        }
    }
    $local:ProgramFilesPath= [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::ProgramFiles)

    # Pre-fetch all online versions in parallel (PS 7+) to avoid N sequential network calls
    # in the loop below. Cache is shared with Show-Office365Modules (60-min TTL).
    if ($PSVersionTable.PSVersion.Major -ge 7) {
        $local:UsePSRG  = $script:myOffice365Services['PSResourceGet']
        $local:ToFetch  = $local:Functions.Module | Where-Object {
            $local:e = $script:myOffice365Services['OnlineVersionCache'][$_]
            $null -eq $local:e -or ([datetime]::Now - $local:e.Fetched).TotalMinutes -ge 60
        }
        if ($local:ToFetch) {
            $local:ToFetch | ForEach-Object -Parallel {
                $local:n = $_
                $local:o = if ($using:UsePSRG) {
                    Find-PSResource -Name $local:n -Prerelease:$using:UsePre -ErrorAction SilentlyContinue
                } else {
                    Find-Module -Name $local:n -AllowPrerelease:$using:UsePre -ErrorAction SilentlyContinue
                }
                [PSCustomObject]@{ Name=$local:n; Version=if($local:o){[string]$local:o.Version}else{$null} }
            } -ThrottleLimit 10 | ForEach-Object {
                $script:myOffice365Services['OnlineVersionCache'][$_.Name] = [PSCustomObject]@{
                    Version = $_.Version; Fetched = [datetime]::Now
                }
            }
        }
    }

    $local:AllInstalled = Get-Module -ListAvailable -ErrorAction SilentlyContinue

    ForEach ( $local:Item in $local:Functions) {

        $local:Module= Get-InstalledRepoModule -Name $local:Item.Module -Repo $local:Item.Repo -AllInstalled $local:AllInstalled

        If( ($local:Module).RepositorySourceLocation) {

            If( -not $local:IsAdmin -and ($local:Module.ModuleBase -like "$local:ProgramFilesPath*")) {
                Write-Host ('{0}: Skipped (installed in AllUsers scope; re-run as administrator to update)' -f $local:Item.Description) -ForegroundColor $script:myConsoleColors.Warning
                Continue
            }

            $local:ModuleScope = if ($local:Module.ModuleBase -like "$local:ProgramFilesPath*") { 'AllUsers' } else { 'CurrentUser' }

            $local:Version = Get-ModuleVersionInfo -Module $local:Module
            Write-Host ('Checking {0}' -f $local:Item.Description) -NoNewLine

            $local:NewerAvailable= $false
            $local:OnlineVer = Get-OnlineModuleVersion -Name $local:Item.Module
            If( $local:OnlineVer) {
                Write-Host (' v{0} (Online v{1})' -f $local:Version, $local:OnlineVer) -NoNewline
                If( (Compare-TextVersionNumber -Version $local:Version -CompareTo $local:OnlineVer) -eq 1) {
                    Write-Host (' Update available') -ForegroundColor $script:myconsoleColors.Error
                    $local:NewerAvailable= $true
                }
                Else{ 
                    Write-Host ''
                }
            }
            Else {
                    # Not installed from online or cannot determine
                    Write-Host (' v{0} (Online N/A)' -f $local:Version) -ForegroundColor $script:myconsoleColors.Warning

            }

            If( $local:NewerAvailable) {
                $local:UpdateSuccess= $false
                Try {
                    Update-myModule -Name $local:Item.Module -Scope $local:ModuleScope
                    $local:UpdateSuccess= $true
                }
                Catch {
                    Write-Host ('Problem updating {0}: {1}' -f $local:Item.Module, $_.Exception.Message) -ForegroundColor $script:myConsoleColors.Error
                }

                If( $local:UpdateSuccess) {

                    Write-Host ('Updated {0} to version {1}' -f $local:Item.Description, $local:OnlineVer) -ForegroundColor $script:myConsoleColors.OK

                    # Uninstall all older versions; use OnlineVer as the new baseline so we
                    # correctly identify old versions regardless of which package manager ran.
                    $local:ModuleVersions= Get-Module -Name $local:Item.Module -ListAvailable -ErrorAction SilentlyContinue |
                        Where-Object { $_.RepositorySourceLocation -and ([System.Uri]($_.RepositorySourceLocation)).Authority -ieq ([System.Uri]($local:Item.Repo)).Authority }
                    $local:LatestVersion = $local:OnlineVer

                    # Uninstall all old versions of module & dependencies
                    If( $local:OnlineVer) {
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
                                    Write-Warning ('Problem uninstalling {0} v{1}: {2}' -f $DepModule.Name, $DepModule.Version, $_.Exception.Message)
                                }
                            }
                        }
                        $local:OldModules= $local:ModuleVersions | Where-Object {$_.Version -ne $local:LatestVersion}
                        If( $local:OldModules) {
                            ForEach( $OldModule in $local:OldModules) {
                                Write-Host ('Uninstalling {0} version {1}' -f $local:Item.Description, $OldModule.Version)
                                Try {
                                    Uninstall-myModule -Name $OldModule.Name -Version $OldModule.Version -IsPrerelease:$OldModule.IsPrerelease
                                }
                                Catch {
                                    Write-Warning ('Problem uninstalling {0} v{1}: {2}' -f $OldModule.Name, $OldModule.Version, $_.Exception.Message)
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
