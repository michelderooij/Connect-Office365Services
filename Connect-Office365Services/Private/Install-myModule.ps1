function Install-myModule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string[]]$Name,
        [switch]$AllowPrerelease,
        [switch]$AllowClobber
    )
    Process {
        If( $script:myOffice365Services['PSResourceGet']) {
            Install-PSResource -Name $Name -Prerelease:$AllowPrerelease -Scope $script:myOffice365Services['Scope'] -TrustRepository
        }
        Else {
            Install-Module -Name $Name -Force -AllowClobber:$AllowClobber -AllowPrerelease:$AllowPrerelease -Scope $script:myOffice365Services['Scope']
        }
    }
}
