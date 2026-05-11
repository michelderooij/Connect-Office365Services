function Get-myPSResourceGetInstalled {
    param(
        # Optional: pass result of (Get-Module -ListAvailable) to avoid an extra filesystem scan.
        $AllInstalled = $null
    )
    If( -not $script:myOffice365Services['PSResourceGet']) {
        $local:candidates = If ($AllInstalled) {
            $AllInstalled | Where-Object Name -eq 'Microsoft.PowerShell.PSResourceGet' | Select-Object -First 1
        } Else {
            Get-Module -Name Microsoft.PowerShell.PSResourceGet -ListAvailable -ErrorAction SilentlyContinue
        }
        $script:myOffice365Services['PSResourceGet'] = $null -ne $local:candidates
    }
}
