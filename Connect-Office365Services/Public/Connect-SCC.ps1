function Connect-SCC {
    if (-not (Get-Module -Name ExchangeOnlineManagement -ErrorAction SilentlyContinue)) {
        Import-Module -Name ExchangeOnlineManagement -ErrorAction SilentlyContinue
    }
    if (Get-Command -Name ExchangeOnlineManagement\Connect-IPPSSession -ErrorAction SilentlyContinue) {
        # Ensure we have an account cached (MSAL) or credentials (legacy)
        if ( -not $script:myOffice365Services['Office365UPN'] -and -not $script:myOffice365Services['Office365Credential']) {
            Get-Office365Credential
        }

        Write-Host 'Connecting to Security & Compliance Center ..'

        $local:connectParams = @{
            ConnectionUri   = $script:myOffice365Services['SCCConnectionEndpointUri']
            PSSessionOption = $script:myOffice365Services['SessionOptions']
        }

        # Modern auth: try AccessToken first, fall back to UPN (WAM SSO), then legacy credential
        $local:exoToken = Get-Office365AccessToken -Scope 'https://outlook.office365.com/.default'
        if ($local:exoToken) {
            $local:connectParams['AccessToken'] = ConvertTo-SecureString $local:exoToken -AsPlainText -Force
            $local:connectParams['UserPrincipalName'] = $script:myOffice365Services['Office365UPN']
        }
        elseif ($script:myOffice365Services['Office365UPN']) {
            $local:connectParams['UserPrincipalName'] = $script:myOffice365Services['Office365UPN']
        }
        elseif ($script:myOffice365Services['Office365Credential']) {
            $local:connectParams['UserPrincipalName'] = $script:myOffice365Services['Office365Credential'].UserName
        }

        $script:myOffice365Services['SessionCC'] = Connect-IPPSSession @local:connectParams
        if ( $script:myOffice365Services['SessionCC'] ) {
            Import-PSSession -Session $script:myOffice365Services['SessionCC'] -AllowClobber
        }
    }
    else {
        Write-Error -Message 'Cannot connect to Security & Compliance Center - module not installed or not loading.'
    }
}
