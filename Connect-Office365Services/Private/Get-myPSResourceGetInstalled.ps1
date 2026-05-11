function Get-myPSResourceGetInstalled {
    If( $script:myOffice365Services['PSResourceGet']) {
        # Already determined
    }
    Else {
        $script:myOffice365Services['PSResourceGet']= $null -ne (Get-Module -Name Microsoft.PowerShell.PSResourceGet -ListAvailable -ErrorAction SilentlyContinue)
    }
}
