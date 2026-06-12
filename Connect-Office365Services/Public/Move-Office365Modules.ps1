function Move-Office365Modules {
    <#
    .SYNOPSIS
    Moves installed Office 365 modules from one PowerShell scope to another.

    .DESCRIPTION
    Finds all supported Office 365 modules installed in the scope opposite to TargetScope,
    installs each of those versions in TargetScope, then uninstalls them from their original scope.
    The Connect-Office365Services module itself is never processed.
    Administrator privileges are required because operations touching AllUsers scope need elevation.

    .PARAMETER TargetScope
    The destination scope to move modules into. Must be AllUsers or CurrentUser.
    Only modules currently installed in the opposite scope are processed.

    .EXAMPLE
    Move-Office365Modules -TargetScope AllUsers
    Moves all supported modules installed in CurrentUser scope to AllUsers scope.

    .EXAMPLE
    Move-Office365Modules -TargetScope CurrentUser
    Moves all supported modules installed in AllUsers scope to CurrentUser scope.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet('AllUsers', 'CurrentUser')]
        [string]$TargetScope
    )

    $local:SelfModuleName = $MyInvocation.MyCommand.Module.Name
    $local:SourceScope = if ($TargetScope -eq 'AllUsers') { 'CurrentUser' } else { 'AllUsers' }

    $local:IsAdmin = Test-IsAdministrator
    if (-not $local:IsAdmin) {
        Write-Warning 'Administrator privileges are required to move modules to or from AllUsers scope.'
        return
    }

    if ((Get-Process -Name powershell, pwsh -ErrorAction SilentlyContinue | Measure-Object).Count -gt 1) {
        Write-Warning 'Running multiple PowerShell sessions; successful module operations might be problematic.'
    }

    $local:ModuleInfoList = Get-Office365ModuleInfo | Where-Object { $_.Module -ne $local:SelfModuleName }
    $local:AllInstalled = Get-Module -ListAvailable -ErrorAction SilentlyContinue
    $local:MovedCount = 0
    $local:FailedCount = 0

    # Warn if Connect-Office365Services itself is not installed in the target scope.
    # The running module cannot move itself; the user must reinstall it manually.
    $local:SelfInTarget = $local:AllInstalled |
        Where-Object { $_.Name -eq $local:SelfModuleName } |
        Where-Object { (Get-ModuleScope -Module $_) -eq $TargetScope }
    if (-not $local:SelfInTarget) {
        Write-Warning ('{0} is not installed in {1} scope. This module cannot move itself — after this operation completes, reinstall it manually in the {1} scope:' -f $local:SelfModuleName, $TargetScope)
        if ($TargetScope -eq 'AllUsers') {
            Write-Warning ('  Install-Module -Name {0} -Scope AllUsers' -f $local:SelfModuleName)
        }
        else {
            Write-Warning ('  Install-Module -Name {0} -Scope CurrentUser' -f $local:SelfModuleName)
        }
    }

    foreach ($local:Item in $local:ModuleInfoList) {
        $local:AllVersions = $local:AllInstalled |
            Where-Object { $_.Name -eq $local:Item.Module } |
            Where-Object {
                $_.RepositorySourceLocation -and
                ([System.Uri]($_.RepositorySourceLocation)).Authority -ieq ([System.Uri]($local:Item.Repo)).Authority
            }

        $local:Candidates = $local:AllVersions | Where-Object { (Get-ModuleScope -Module $_) -eq $local:SourceScope }
        $local:AlreadyInTarget = $local:AllVersions | Where-Object { (Get-ModuleScope -Module $_) -eq $TargetScope }

        if (-not $local:Candidates) {
            foreach ($local:TargetModule in $local:AlreadyInTarget) {
                $local:TargetVer = Get-ModuleVersionInfo -Module $local:TargetModule
                Write-Host ('{0} v{1} already in {2}' -f $local:TargetModule.Name, $local:TargetVer, $TargetScope)
            }
            continue
        }

        foreach ($local:Module in $local:Candidates) {
            $local:FullVer = Get-ModuleVersionInfo -Module $local:Module
            $local:IsPrerelease = $local:FullVer -match '-'

            Write-Host ('Moving {0} v{1} from {2} to {3}...' -f $local:Module.Name, $local:FullVer, $local:SourceScope, $TargetScope)

            try {
                if ($script:myOffice365Services['PSResourceGet']) {
                    Install-PSResource -Name $local:Module.Name -Version ('[{0}]' -f $local:FullVer) -Scope $TargetScope -TrustRepository -Prerelease:$local:IsPrerelease -ErrorAction Stop
                }
                else {
                    Install-Module -Name $local:Module.Name -RequiredVersion $local:FullVer -Scope $TargetScope -Force -AllowClobber -AllowPrerelease:$local:IsPrerelease -ErrorAction Stop
                }
            }
            catch {
                $local:ErrMsg = $_.Exception.Message
                if ($local:IsPrerelease -and ($local:ErrMsg -match 'could not be installed from repository|No match was found for the specified search criteria')) {
                    Write-Warning ('Skipping {0} v{1}: pre-release version not available on the repository. Install it manually from its original source.' -f $local:Module.Name, $local:FullVer)
                }
                else {
                    Write-Error ('Failed to install {0} v{1} in {2}: {3}' -f $local:Module.Name, $local:FullVer, $TargetScope, $local:ErrMsg)
                }
                $local:FailedCount++
                continue
            }

            # Verify the module is actually present in the target scope before removing it from source
            $local:Verified = Get-Module -Name $local:Module.Name -ListAvailable -ErrorAction SilentlyContinue |
                Where-Object {
                    $_.RepositorySourceLocation -and
                    ([System.Uri]($_.RepositorySourceLocation)).Authority -ieq ([System.Uri]($local:Item.Repo)).Authority
                } |
                Where-Object { (Get-ModuleScope -Module $_) -eq $TargetScope } |
                Where-Object { (Get-ModuleVersionInfo -Module $_) -eq $local:FullVer }

            if (-not $local:Verified) {
                Write-Error ('Install of {0} v{1} in {2} could not be verified; source copy retained.' -f $local:Module.Name, $local:FullVer, $TargetScope)
                $local:FailedCount++
                continue
            }

            Write-Host ('  Installed {0} v{1} in {2}' -f $local:Module.Name, $local:FullVer, $TargetScope)

            try {
                Uninstall-myModule -Name $local:Module.Name -Version $local:FullVer -IsPrerelease:$local:IsPrerelease -Scope $local:SourceScope
                Write-Host ('  Uninstalled {0} v{1} from {2}' -f $local:Module.Name, $local:FullVer, $local:SourceScope)
                $local:MovedCount++
            }
            catch {
                Write-Error ('Failed to uninstall {0} v{1} from {2}: {3}' -f $local:Module.Name, $local:FullVer, $local:SourceScope, $_.Exception.Message)
                $local:FailedCount++
            }
        }
    }

    Write-Host ''
    Write-Host ('Done. Moved: {0}, Failed: {1}' -f $local:MovedCount, $local:FailedCount)
}
