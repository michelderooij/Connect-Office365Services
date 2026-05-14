function Connect-MSTeams {
    if (-not (Get-Module -Name MicrosoftTeams -ListAvailable)) {
        Import-Module -Name MicrosoftTeams -ErrorAction SilentlyContinue
    }
    if (Get-Command -Name Connect-MicrosoftTeams -ErrorAction SilentlyContinue) {
        if ( -not $script:myOffice365Services['Office365UPN'] -and -not $script:myOffice365Services['Office365Credential']) {
            Get-Office365Credential
        }
        $local:upn = if ($script:myOffice365Services['Office365UPN']) {
            $script:myOffice365Services['Office365UPN']
        }
        else {
            $script:myOffice365Services['Office365Credential'].UserName
        }
        Write-Host ('Connecting to Microsoft Teams using {0} ..' -f $local:upn)

        # Attempt silent token injection so the Teams module does not open a second browser prompt.
        # Connect-MicrosoftTeams accepts:
        #   -AccessToken   — Graph access token (SecureString)
        #   -MsAccessToken — Teams service token (SecureString)
        # Both are acquired silently from the MSAL cache seeded by Get-Office365Credential.
        $local:graphToken = Get-Office365AccessToken -Scope 'https://graph.microsoft.com/.default'
        $local:teamsToken = Get-Office365AccessToken -Scope 'https://api.spaces.skype.com/.default'
        if ($local:graphToken -and $local:teamsToken) {
            Connect-MicrosoftTeams `
                -AccessToken (ConvertTo-SecureString $local:graphToken -AsPlainText -Force) `
                -MsAccessToken (ConvertTo-SecureString $local:teamsToken -AsPlainText -Force) `
                -TenantId $script:myOffice365Services['TenantID']
        }
        else {
            # Fall back: let the Teams module run its own MSAL/WAM flow using the known UPN.
            Connect-MicrosoftTeams -AccountId $local:upn -TenantId $script:myOffice365Services['TenantID']
        }
    }
    else {
        Write-Error -Message 'Cannot connect to Microsoft Teams - module not installed or not loading.'
    }
}
