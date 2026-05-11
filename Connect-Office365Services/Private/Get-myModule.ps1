function Get-myModule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false, ValueFromPipeline=$true)]
        [string[]]$Name,
        [switch]$ListAvailable,
        [switch]$AllowPrerelease,
        [switch]$All
    )
    Process {
        If( $script:myOffice365Services['PSResourceGet']) {
            If( $All) {
                # -Version '*' retrieves all installed versions
                Get-PSResource -Name $Name -Version '*' -ErrorAction SilentlyContinue
            }
            Else {
                Get-PSResource -Name $Name -Scope $script:myOffice365Services['Scope'] -ErrorAction SilentlyContinue
            }
        }
        Else {
            Get-Module -Name $Name -ListAvailable:$ListAvailable -ErrorAction SilentlyContinue
        }
    }
}
