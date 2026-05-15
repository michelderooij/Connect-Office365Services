function Get-Office365Tenant {
    $local:upn = if ($script:myOffice365Services['Office365UPN']) {
        $script:myOffice365Services['Office365UPN']
    }
    elseif ($script:myOffice365Services['Office365Credential']) {
        $script:myOffice365Services['Office365Credential'].UserName
    }
    If ($local:upn) {
        $local:domain = ($local:upn -split '@')[1]
        $local:OpenIdInfo = Invoke-RestMethod ('https://login.windows.net/{0}/.well-known/openid-configuration' -f $local:domain) -Method GET
        $script:myOffice365Services['Office365Tenant'] = $local:OpenIdInfo.userinfo_endpoint.Split('/')[3]
    }
    Else {
        $script:myOffice365Services['Office365Tenant'] = Read-Host -Prompt 'Enter tenant ID, e.g. contoso for contoso.onmicrosoft.com'
    }
}
