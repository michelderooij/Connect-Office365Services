function Connect-IPPSSession {
    If(!( Get-Module -Name ExchangeOnlineManagement -ErrorAction SilentlyContinue)) {
        Import-Module -Name ExchangeOnlineManagement -ErrorAction SilentlyContinue
    }
    If( Get-Command -Name Connect-IPPSSession -ErrorAction SilentlyContinue) {
        # Fixed: added null guard for credentials before accessing .UserName
        If ( !($script:myOffice365Services['Office365Credentials'])) { Get-Office365Credentials }
        Write-Host ('Connecting to Security & Compliance Center ..')
        $script:myOffice365Services['SessionCC'] = ExchangeOnlineManagement\Connect-IPPSSession -ConnectionUri $script:myOffice365Services['SCCConnectionEndpointUri'] -UserPrincipalName ($script:myOffice365Services['Office365Credentials']).UserName -PSSessionOption $script:myOffice365Services['SessionOptions']
        If ( $script:myOffice365Services['SessionCC'] ) {
            Import-PSSession -Session $script:myOffice365Services['SessionCC'] -AllowClobber
        }
    }
    Else {
        Write-Error -Message 'Cannot connect to Security & Compliance Center - module not installed or not loading.'
    }
}
