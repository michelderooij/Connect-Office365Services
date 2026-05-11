function Update-myModule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string[]]$Name
    )
    Process {
        If( $script:myOffice365Services['PSResourceGet']) {
            Update-PSResource -Name $Name -Scope $script:myOffice365Services['Scope'] -Force -AcceptLicense -Prerelease:$script:myOffice365Services['AllowPrerelease'] -TrustRepository
        }
        Else {
            Update-Module -Name $Name -Scope $script:myOffice365Services['Scope'] -Force -AllowClobber -AcceptLicense -AllowPrerelease:$script:myOffice365Services['AllowPrerelease']
        }
    }
}
