function Connect-AIP {
    If (-not (Get-Module -Name AIPService -ListAvailable)) {
        Import-Module -Name AIPService -ErrorAction SilentlyContinue
    }
    If (Get-Command -Name Connect-AipService -ErrorAction SilentlyContinue) {
        # Ensure we have an account cached (MSAL) or credentials (legacy)
        If ( -not $script:myOffice365Services['Office365UPN'] -and -not $script:myOffice365Services['Office365Credential']) {
            Get-Office365Credential
        }

        # Modern auth: acquire AADRM-scoped token and pass via -AccessToken (plain String)
        $local:aipToken = Get-Office365AccessToken -Scope 'https://api.aadrm.com/.default'
        If ($local:aipToken) {
            $local:displayName = $script:myOffice365Services['Office365UPN']
            Write-Host ('Connecting to Azure Information Protection using {0}' -f $local:displayName)
            Connect-AipService -AccessToken $local:aipToken -TenantId $script:myOffice365Services['TenantID']
        }
        ElseIf ($script:myOffice365Services['Office365Credential']) {
            # Legacy PSCredential fallback
            Write-Host ('Connecting to Azure Information Protection using {0}' -f $script:myOffice365Services['Office365Credential'].UserName)
            Connect-AipService -Credential $script:myOffice365Services['Office365Credential']
        }
        Else {
            Write-Host 'Connecting to Azure Information Protection ..'
            Connect-AipService
        }
    }
    Else {
        Write-Error -Message 'Cannot connect to Azure Information Protection - module not installed or not loading.'
    }
}
