function Connect-AIP {
    if (-not (Get-Module -Name AIPService)) {
        Import-Module -Name AIPService -ErrorAction SilentlyContinue
    }
    if (Get-Command -Name Connect-AipService -ErrorAction SilentlyContinue) {
        # Ensure we have an account cached (MSAL) or credentials (legacy)
        if ( -not $script:myOffice365Services['Office365UPN'] -and -not $script:myOffice365Services['Office365Credential']) {
            if ($script:myOffice365Services['NoAutoConnect']) {
                Write-Error 'No credentials cached. Run Get-Office365Credential first or supply credentials explicitly.'
                return
            }
            Get-Office365Credential
        }

        # Modern auth: acquire AADRM-scoped token and pass via -AccessToken (plain String)
        $local:aipToken = Get-Office365AccessToken -Scope 'https://api.aadrm.com/.default'
        if ($local:aipToken) {
            $local:displayName = $script:myOffice365Services['Office365UPN']
            Write-Host ('Connecting to Azure Information Protection using {0}' -f $local:displayName)
            Connect-AipService -AccessToken $local:aipToken -TenantId $script:myOffice365Services['TenantID']
        }
        elseif ($script:myOffice365Services['Office365Credential']) {
            # Legacy PSCredential fallback
            Write-Host ('Connecting to Azure Information Protection using {0}' -f $script:myOffice365Services['Office365Credential'].UserName)
            Connect-AipService -Credential $script:myOffice365Services['Office365Credential']
        }
        else {
            Write-Host 'Connecting to Azure Information Protection ..'
            Connect-AipService
        }
        $script:myOffice365Services['ConnectedAIP'] = $true
    }
    else {
        Write-Error -Message 'Cannot connect to Azure Information Protection - module not installed or not loading.'
    }
}
