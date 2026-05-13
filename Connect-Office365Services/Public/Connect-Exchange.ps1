function Connect-Exchange {
    If ( !($script:myOffice365Services['OnPremisesCredentials'])) { Get-OnPremisesCredentials }
    If ( !($script:myOffice365Services['ExchangeOnPremisesFQDN'])) { Get-ExchangeOnPremisesFQDN }
    # Fixed: removed erroneous '!' — only connect when credentials ARE present
    If ( $script:myOffice365Services['OnPremisesCredentials']) {
        Write-Host ('Connecting to Exchange On-Premises {0} using {1} ..' -f $script:myOffice365Services['ExchangeOnPremisesFQDN'], $script:myOffice365Services['OnPremisesCredentials'].UserName)
        $script:myOffice365Services['SessionExchange'] = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$($script:myOffice365Services['ExchangeOnPremisesFQDN'])/PowerShell" -Credential $script:myOffice365Services['OnPremisesCredentials'] -Authentication Kerberos -AllowRedirection -SessionOption $script:myOffice365Services['SessionOptions']
        If ( $script:myOffice365Services['SessionExchange']) {Import-PSSession -Session $script:myOffice365Services['SessionExchange'] -AllowClobber}
    }
}
