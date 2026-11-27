function Connect-SPO {
    <#
    .SYNOPSIS
    Connects to SharePoint Online.

    .DESCRIPTION
    Uses the supplied admin URL and credentials when provided. Otherwise, resolves
    the URL and authentication from the cached Office 365 tenant and identity.
    Other parameters supported by Connect-SPOService are passed directly to it.

    .PARAMETER Url
    The HTTPS URL of the SharePoint Online administration site.

    .PARAMETER Credential
    The credentials to pass to Connect-SPOService instead of using a cached token.
    If no tenant is cached and the username uses a custom domain, also supply -Url.

    .EXAMPLE
    Connect-SPO
    Connects using the cached tenant and identity.

    .EXAMPLE
    Connect-SPO -Url https://contoso-admin.sharepoint.com -Credential $credential
    Connects using the specified administration URL and credentials.
    #>
    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            $local:uri = $null
            if (-not [uri]::TryCreate($_, [System.UriKind]::Absolute, [ref]$local:uri) -or
                $local:uri.Scheme -ne 'https' -or -not $local:uri.Host -or $local:uri.UserInfo) {
                throw 'Url must be an absolute HTTPS URL without embedded credentials.'
            }
            $true
        })]
        [string]$Url,

        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]$Credential
    )

    dynamicparam {
        Get-Office365ConnectParameters -Module 'Microsoft.Online.SharePoint.PowerShell' -Command 'Connect-SPOService' -Exclude @('Url', 'Credential')
    }
    process {
        if (-not (Get-Module -Name Microsoft.Online.SharePoint.PowerShell)) {
            Import-Module -Name Microsoft.Online.SharePoint.PowerShell -ErrorAction SilentlyContinue
        }
        if (Get-Command -Name Connect-SPOService -ErrorAction SilentlyContinue) {
            $local:explicitCredential = $PSBoundParameters.ContainsKey('Credential')
            $local:explicitParams = Get-Office365ExplicitConnectParameters -BoundParameters $PSBoundParameters -Exclude @('Url')
            $local:explicitAuth = $local:explicitParams.Count -gt 0
            $local:spoCmd = Get-Command -Name Connect-SPOService -ErrorAction SilentlyContinue
            if ($local:explicitCredential -and -not $local:spoCmd.Parameters.ContainsKey('Credential')) {
                return (New-Office365ConnectionStatus -Service 'SharePoint Online' -StateKey 'ConnectedSPO' -Connected $false -ErrorMessage 'The installed Connect-SPOService command does not support the Credential parameter.')
            }
            if (-not $local:explicitAuth -and -not $script:myOffice365Services['Office365UPN'] -and -not $script:myOffice365Services['Office365Credential']) {
                if ($script:myOffice365Services['NoAutoConnect']) {
                    return (New-Office365ConnectionStatus -Service 'SharePoint Online' -StateKey 'ConnectedSPO' -Connected $false -ErrorMessage 'No credentials cached. Run Get-Office365Credential first or supply credentials explicitly.')
                }
                try {
                    $null = Get-Office365Credential
                }
                catch {
                    return (New-Office365ConnectionStatus -Service 'SharePoint Online' -StateKey 'ConnectedSPO' -Connected $false -ErrorMessage ('Could not acquire credentials for SharePoint Online: {0}' -f $_.Exception.Message))
                }
            }
            $local:upn = if ($local:explicitCredential) {
                $Credential.UserName
            }
            elseif ($script:myOffice365Services['Office365UPN']) {
                $script:myOffice365Services['Office365UPN']
            }
            else {
                $script:myOffice365Services['Office365Credential'].UserName
            }
            Write-Host 'Connecting to SharePoint Online ..'
            # Resolve tenant name for the admin URL when not already cached.
            # Priority: onmicrosoft.com UPN → Graph domains API → error
            if (-not $Url -and -not $script:myOffice365Services['Office365Tenant']) {
                if ($local:upn -like '*@*.onmicrosoft.com') {
                    $script:myOffice365Services['Office365Tenant'] = ($local:upn -split '@')[1] -replace '\.onmicrosoft\.com$', ''
                }
                elseif (-not $local:explicitAuth) {
                    # Resolve the initial domain (*.onmicrosoft.com) from Graph using the cached MSAL token.
                    # Fetch all domains and filter client-side — avoids OData advanced query permission requirements.
                    try {
                        $local:graphEnvironment = Get-Office365GraphEnvironment
                        $local:graphToken = Get-Office365AccessToken -Scope ('{0}/.default' -f $local:graphEnvironment.GraphEndpoint)
                    }
                    catch {
                        return (New-Office365ConnectionStatus -Service 'SharePoint Online' -StateKey 'ConnectedSPO' -Connected $false -ErrorMessage ('Could not acquire a tenant lookup token: {0}' -f $_.Exception.Message))
                    }
                    if ($local:graphToken) {
                        try {
                            $local:domainResp = Invoke-RestMethod `
                                -Uri ('{0}/v1.0/domains' -f $local:graphEnvironment.GraphEndpoint) `
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
            if (-not $Url) {
                if (-not $script:myOffice365Services['Office365Tenant']) {
                    if ($local:explicitAuth) {
                        return (New-Office365ConnectionStatus -Service 'SharePoint Online' -StateKey 'ConnectedSPO' -Connected $false -ErrorMessage 'Cannot determine the SharePoint Online tenant for the supplied options. Supply -Url.')
                    }
                    try {
                        $null = Get-Office365Tenant
                    }
                    catch {
                        return (New-Office365ConnectionStatus -Service 'SharePoint Online' -StateKey 'ConnectedSPO' -Connected $false -ErrorMessage ('Could not resolve the SharePoint Online tenant: {0}' -f $_.Exception.Message))
                    }
                }
                if (-not $script:myOffice365Services['Office365Tenant']) {
                    return (New-Office365ConnectionStatus -Service 'SharePoint Online' -StateKey 'ConnectedSPO' -Connected $false -ErrorMessage 'Cannot determine SharePoint Online tenant name. Run Get-Office365Tenant.')
                }
                $Url = 'https://{0}-admin.sharepoint.com' -f $script:myOffice365Services['Office365Tenant']
            }
            if ($local:explicitAuth) {
                $local:explicitParams['Url'] = $Url
                if ($local:explicitCredential -and -not $local:explicitParams.ContainsKey('Region') -and $script:myOffice365Services['SharePointRegion']) {
                    $local:explicitParams['Region'] = $script:myOffice365Services['SharePointRegion']
                }
                try {
                    $null = Connect-SPOService @local:explicitParams -ErrorAction Stop
                }
                catch {
                    return (New-Office365ConnectionStatus -Service 'SharePoint Online' -StateKey 'ConnectedSPO' -Connected $false -ErrorMessage ('Cannot connect to SharePoint Online: {0}' -f $_.Exception.Message))
                }
                return (New-Office365ConnectionStatus -Service 'SharePoint Online' -StateKey 'ConnectedSPO' -Connected $true -Detail $Url)
            }
            $local:Parms = @{ Url = $Url }
            if ($script:myOffice365Services['SharePointRegion']) {
                $local:Parms['Region'] = $script:myOffice365Services['SharePointRegion']
            }
            # Prefer MSAL token injection; fall back to PSCredential so SPO triggers its own
            # modern auth flow (MFA prompt only) rather than failing with no auth context.
            # SPO's OAuth resource is the tenant root URL, not the admin URL.
            if ($PSBoundParameters.ContainsKey('Url')) {
                $local:uri = [uri]$Url
                $local:resourceAuthority = $local:uri.Authority -replace '^([^.]+)-admin(\.)', '$1$2'
                $local:scope = '{0}://{1}/.default' -f $local:uri.Scheme, $local:resourceAuthority
            }
            else {
                $local:scope = 'https://{0}.sharepoint.com/.default' -f $script:myOffice365Services['Office365Tenant']
            }
            try {
                $local:spoToken = Get-Office365AccessToken -Scope $local:scope
            }
            catch {
                return (New-Office365ConnectionStatus -Service 'SharePoint Online' -StateKey 'ConnectedSPO' -Connected $false -ErrorMessage ('Could not acquire a SharePoint Online token: {0}' -f $_.Exception.Message))
            }
            $local:connected = $false

            # Auth priority 1: inject MSAL access token when the parameter is available.
            if ($local:spoToken -and $local:spoCmd.Parameters.ContainsKey('AccessToken')) {
                try {
                    $null = Connect-SPOService @local:Parms -AccessToken $local:spoToken -ErrorAction Stop
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
                    $null = Connect-SPOService @local:Parms -ErrorAction Stop
                    $local:connected = $true
                }
                catch {
                    $local:connectionError = 'Cannot connect to SharePoint Online: {0}' -f $_.Exception.Message
                }
            }

            if ($local:connected) {
                New-Office365ConnectionStatus -Service 'SharePoint Online' -StateKey 'ConnectedSPO' -Connected $true -Detail $Url
            }
            else {
                New-Office365ConnectionStatus -Service 'SharePoint Online' -StateKey 'ConnectedSPO' -Connected $false -ErrorMessage $local:connectionError
            }
        }
        else {
            New-Office365ConnectionStatus -Service 'SharePoint Online' -StateKey 'ConnectedSPO' -Connected $false -ErrorMessage 'Cannot connect to SharePoint Online - module not installed or not loading.'
        }
    }
}
