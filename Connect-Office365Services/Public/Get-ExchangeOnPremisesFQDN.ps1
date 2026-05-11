function Get-ExchangeOnPremisesFQDN {
    # RFC 1123 hostname: labels of 1-63 chars separated by dots, no scheme or path allowed
    $local:HostnamePattern = '^([a-zA-Z0-9]([a-zA-Z0-9\-]{0,61}[a-zA-Z0-9])?\.)+[a-zA-Z]{2,}$'
    While ($true) {
        $local:input = (Read-Host -Prompt 'Enter Exchange On-Premises endpoint, e.g. exchange1.contoso.com').Trim()
        If( [string]::IsNullOrEmpty( $local:input)) {
            return
        }
        If( $local:input -match $local:HostnamePattern) {
            $script:myOffice365Services['ExchangeOnPremisesFQDN'] = $local:input
            return
        }
        Write-Warning ('Invalid hostname "{0}". Enter a valid FQDN (e.g. exchange1.contoso.com) or leave empty to cancel.' -f $local:input)
    }
}
