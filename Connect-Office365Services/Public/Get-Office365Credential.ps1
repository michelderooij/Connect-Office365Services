function Get-Office365Credential {
    # Attempt modern auth via MSAL.NET (no PSCredential needed when available).
    $local:graphEnvironment = Get-Office365GraphEnvironment
    $local:graphScope = '{0}/.default' -f $local:graphEnvironment.GraphEndpoint
    $local:msalToken = Get-Office365AccessToken -Scope $local:graphScope
    if (-not $local:msalToken) {
        # MSAL.NET not available or token acquisition failed — fall back to PSCredential.
        $local:prevUser = if ($script:myOffice365Services['Office365Credential']) {
            $script:myOffice365Services['Office365Credential'].UserName
        } else { '' }
        $script:myOffice365Services['Office365Credential'] = $host.ui.PromptForCredential(
            'Office 365 Credentials',
            'Please enter your Office 365 credentials',
            $local:prevUser, '')
    }
    Get-TenantID
    Set-WindowTitle
}
