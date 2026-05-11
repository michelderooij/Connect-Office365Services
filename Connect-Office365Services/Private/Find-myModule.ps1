function Find-myModule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false, ValueFromPipeline=$true)]
        [string[]]$Name
    )
    Process {
        If( $script:myOffice365Services['PSResourceGet']) {
            Find-PSResource -Name $Name -Prerelease:$script:myOffice365Services['AllowPrerelease'] -ErrorAction SilentlyContinue
        }
        Else {
            Find-Module -Name $Name -AllowPrerelease:$script:myOffice365Services['AllowPrerelease'] -ErrorAction SilentlyContinue
        }
    }
}
