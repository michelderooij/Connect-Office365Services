function Update-myModule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string[]]$Name,
        [ValidateSet('CurrentUser','AllUsers')]
        [string]$Scope = $script:myOffice365Services['Scope']
    )
    Process {
        If( $script:myOffice365Services['PSResourceGet']) {
            Try {
                Update-PSResource -Name $Name -Scope $Scope -Force -AcceptLicense -Prerelease:$script:myOffice365Services['AllowPrerelease'] -TrustRepository -ErrorAction Stop
            }
            Catch {
                # Update-PSResource failed (e.g. package not tracked in this scope).
                # Re-install via PSResourceGet with -Reinstall to force an upgrade regardless
                # of how the module was originally installed (Install-Module or Install-PSResource).
                Install-PSResource -Name $Name -Scope $Scope -Reinstall -AcceptLicense -Prerelease:$script:myOffice365Services['AllowPrerelease'] -TrustRepository -ErrorAction Stop
            }
        }
        Else {
            # Note: Update-Module does not support -Scope, -AllowClobber, or -AcceptLicense.
            Update-Module -Name $Name -Force -AllowPrerelease:$script:myOffice365Services['AllowPrerelease'] -ErrorAction Stop
        }
    }
}
