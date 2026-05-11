function Get-Office365Tenant {
    If( $script:myOffice365Services['Office365Credentials']) {
        $local:OpenIdInfo= Invoke-RestMethod ('https://login.windows.net/{0}/.well-known/openid-configuration' -f ($script:myOffice365Services['Office365Credentials'].UserName.Split('@')[1])) -Method GET
        $script:myOffice365Services['Office365Tenant']= $local:OpenIdInfo.userinfo_endpoint.Split('/')[3]
    }
    Else {
        $script:myOffice365Services['Office365Tenant'] = Read-Host -Prompt 'Enter tenant ID, e.g. contoso for contoso.onmicrosoft.com'
    }
}
