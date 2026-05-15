function Connect-SPO {
    if (-not (Get-Module -Name Microsoft.Online.SharePoint.PowerShell)) {
        Import-Module -Name Microsoft.Online.SharePoint.PowerShell -ErrorAction SilentlyContinue
    }
    if (Get-Command -Name Connect-SPOService -ErrorAction SilentlyContinue) {
        if ( -not $script:myOffice365Services['Office365UPN'] -and -not $script:myOffice365Services['Office365Credential']) {
            if ($script:myOffice365Services['NoAutoConnect']) {
                Write-Error 'No credentials cached. Run Get-Office365Credential first or supply credentials explicitly.'
                return
            }
            Get-Office365Credential
        }
        $local:upn = if ($script:myOffice365Services['Office365UPN']) {
            $script:myOffice365Services['Office365UPN']
        }
        else {
            $script:myOffice365Services['Office365Credential'].UserName
        }
        Write-Host 'Connecting to SharePoint Online ..'
        # Resolve tenant name for the admin URL when not already cached.
        # Priority: onmicrosoft.com UPN → Graph domains API → error
        if (-not $script:myOffice365Services['Office365Tenant']) {
            if ($local:upn -like '*@*.onmicrosoft.com') {
                $script:myOffice365Services['Office365Tenant'] = ($local:upn -split '@')[1] -replace '\.onmicrosoft\.com$', ''
            }
            else {
                # Resolve the initial domain (*.onmicrosoft.com) from Graph using the cached MSAL token.
                # Fetch all domains and filter client-side — avoids OData advanced query permission requirements.
                $local:graphToken = Get-Office365AccessToken -Scope 'https://graph.microsoft.com/.default'
                if ($local:graphToken) {
                    try {
                        $local:domainResp = Invoke-RestMethod `
                            -Uri 'https://graph.microsoft.com/v1.0/domains' `
                            -Headers @{ Authorization = "Bearer $local:graphToken" } `
                            -ErrorAction Stop
                        $local:initialDomain = $local:domainResp.value |
                        Where-Object { $_.isInitial -eq $true } |
                        Select-Object -First 1 -ExpandProperty id
                        if ($local:initialDomain -match '^(.+)\.onmicrosoft\.com$') {
                            $script:myOffice365Services['Office365Tenant'] = $Matches[1]
                        }
                    }
                    catch {
                        Write-Verbose ('Graph domain lookup failed: {0}' -f $_.Exception.Message)
                    }
                }
            }
        }
        if (-not $script:myOffice365Services['Office365Tenant']) {
            Get-Office365Tenant
        }
        if (-not $script:myOffice365Services['Office365Tenant']) {
            Write-Error 'Cannot determine SharePoint Online tenant name. Run Get-Office365Tenant.'
            return
        }
        $local:adminUrl = 'https://{0}-admin.sharepoint.com' -f $script:myOffice365Services['Office365Tenant']
        $local:Parms = @{ Url = $local:adminUrl }
        if ($script:myOffice365Services['SharePointRegion']) {
            $local:Parms['Region'] = $script:myOffice365Services['SharePointRegion']
        }
        # Prefer MSAL token injection; fall back to PSCredential so SPO triggers its own
        # modern auth flow (MFA prompt only) rather than failing with no auth context.
        # SPO's OAuth resource is the tenant root URL, not the admin URL.
        $local:spoToken = Get-Office365AccessToken -Scope ('https://{0}.sharepoint.com/.default' -f $script:myOffice365Services['Office365Tenant'])
        $local:spoCmd = Get-Command -Name Connect-SPOService -ErrorAction SilentlyContinue
        $local:connected = $false

        # Auth priority 1: inject MSAL access token when the parameter is available.
        if ($local:spoToken -and $local:spoCmd.Parameters.ContainsKey('AccessToken')) {
            try {
                Connect-SPOService @local:Parms -AccessToken $local:spoToken -ErrorAction Stop
                $local:connected = $true
            }
            catch {
                Write-Verbose ('SPO access-token auth failed, trying credential fallback: {0}' -f $_.Exception.Message)
            }
        }

        # Auth priority 2: PSCredential, or let SPO trigger its own modern-auth browser prompt.
        if (-not $local:connected) {
            if ($script:myOffice365Services['Office365Credential'] -and $local:spoCmd.Parameters.ContainsKey('Credential')) {
                $local:Parms['Credential'] = $script:myOffice365Services['Office365Credential']
            }
            try {
                Connect-SPOService @local:Parms -ErrorAction Stop
                $local:connected = $true
            }
            catch {
                Write-Error ('Cannot connect to SharePoint Online: {0}' -f $_.Exception.Message)
            }
        }

        if ($local:connected) {
            $script:myOffice365Services['ConnectedSPO'] = $true
        }
    }
    else {
        Write-Error -Message 'Cannot connect to SharePoint Online - module not installed or not loading.'
    }
}
