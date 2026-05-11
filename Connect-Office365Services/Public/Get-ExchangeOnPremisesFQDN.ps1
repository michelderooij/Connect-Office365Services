function Get-ExchangeOnPremisesFQDN {
    $script:myOffice365Services['ExchangeOnPremisesFQDN'] = Read-Host -Prompt 'Enter Exchange On-Premises endpoint, e.g. exchange1.contoso.com'
}
