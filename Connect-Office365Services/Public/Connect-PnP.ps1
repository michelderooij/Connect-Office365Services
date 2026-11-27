function Connect-PnP {
    <#
    .SYNOPSIS
    Connects to a SharePoint Online site collection using PnP PowerShell.

    .DESCRIPTION
    Imports PnP.PowerShell and calls Connect-PnPOnline for the specified site URL.
    When -SiteUrl is omitted the tenant root site (https://<tenant>.sharepoint.com)
    is used, provided the tenant name is already configured in the module state.

    An MSAL token scoped to the site's host is acquired when possible so no
    additional browser prompt is needed. Falls back to Connect-PnPOnline's own
    interactive (-Interactive) flow otherwise.
    Explicit vendor parameters use Connect-PnPOnline's own authentication instead.

    .PARAMETER SiteUrl
    The full URL of the SharePoint site collection to connect to.
    When omitted, defaults to https://<tenant>.sharepoint.com.

    .EXAMPLE
    Connect-PnP
    Connects to the tenant root site using the cached tenant name.

    .EXAMPLE
    Connect-PnP -SiteUrl https://contoso.sharepoint.com/sites/HR
    Connects to the HR site collection.
    #>
    [CmdletBinding()]
    param(
        [Alias('Url')]
        [string]$SiteUrl
    )

    dynamicparam {
        Get-Office365ConnectParameters -Module 'PnP.PowerShell' -Command 'Connect-PnPOnline' -Exclude @('Url', 'ReturnConnection')
    }
    process {
        # Resolve site URL — parameter wins, then fall back to tenant root
        if (-not $SiteUrl) {
            $local:tenant = $script:myOffice365Services['Office365Tenant']
            if (-not $local:tenant) {
                return (New-Office365ConnectionStatus -Service 'PnP (SharePoint)' -StateKey 'ConnectedPnP' -Connected $false -ErrorMessage 'No -SiteUrl supplied and no tenant name is configured. Run Get-Office365Tenant or supply -SiteUrl explicitly.')
            }
            $SiteUrl = 'https://{0}.sharepoint.com' -f $local:tenant
        }

        # Module guard
        if (-not (Get-Module -Name PnP.PowerShell -ErrorAction SilentlyContinue)) {
            Import-Module -Name PnP.PowerShell -ErrorAction SilentlyContinue
        }
        if (-not (Get-Command -Name Connect-PnPOnline -ErrorAction SilentlyContinue)) {
            return (New-Office365ConnectionStatus -Service 'PnP (SharePoint)' -StateKey 'ConnectedPnP' -Connected $false -ErrorMessage 'Cannot connect via PnP - module not installed or not loading. Install PnP.PowerShell.')
        }

        $local:connectParams = Get-Office365ExplicitConnectParameters -BoundParameters $PSBoundParameters -Exclude @('SiteUrl')
        if ($local:connectParams.Count -gt 0) {
            $local:connectParams['Url'] = $SiteUrl
            try {
                Write-Host ('Connecting to {0} using the specified options ..' -f $SiteUrl)
                $null = Connect-PnPOnline @local:connectParams -ErrorAction Stop
                if ((Get-Command -Name Get-PnPConnection -ErrorAction SilentlyContinue) -and -not (Get-PnPConnection -ErrorAction Stop)) {
                    return (New-Office365ConnectionStatus -Service 'PnP (SharePoint)' -StateKey 'ConnectedPnP' -Connected $false -ErrorMessage 'Connect-PnPOnline completed without an active PnP connection.')
                }
            }
            catch {
                $script:myOffice365Services['PnPSiteUrl'] = ''
                return (New-Office365ConnectionStatus -Service 'PnP (SharePoint)' -StateKey 'ConnectedPnP' -Connected $false -ErrorMessage ('Cannot connect to PnP: {0}' -f $_.Exception.Message))
            }
            $script:myOffice365Services['PnPSiteUrl'] = $SiteUrl
            return (New-Office365ConnectionStatus -Service 'PnP (SharePoint)' -StateKey 'ConnectedPnP' -Connected $true -Detail $SiteUrl)
        }
        # Ensure we have an account cached (MSAL) or credentials (legacy)
        if (-not $script:myOffice365Services['Office365UPN'] -and -not $script:myOffice365Services['Office365Credential']) {
            if ($script:myOffice365Services['NoAutoConnect']) {
                return (New-Office365ConnectionStatus -Service 'PnP (SharePoint)' -StateKey 'ConnectedPnP' -Connected $false -ErrorMessage 'No credentials cached. Run Get-Office365Credential first or supply credentials explicitly.')
            }
            try {
                $null = Get-Office365Credential
            }
            catch {
                return (New-Office365ConnectionStatus -Service 'PnP (SharePoint)' -StateKey 'ConnectedPnP' -Connected $false -ErrorMessage ('Could not acquire credentials for PnP: {0}' -f $_.Exception.Message))
            }
        }

        $local:upn = if ($script:myOffice365Services['Office365UPN']) {
            $script:myOffice365Services['Office365UPN']
        }
        else {
            $script:myOffice365Services['Office365Credential'].UserName
        }

        # Scope token to the site's tenant host (e.g. https://contoso.sharepoint.com/.default)
        $local:siteUri     = [uri]$SiteUrl
        $local:tenantScope = '{0}://{1}/.default' -f $local:siteUri.Scheme, $local:siteUri.Host

        try {
            $local:pnpToken = Get-Office365AccessToken -Scope $local:tenantScope
            if ($local:pnpToken) {
                Write-Host ('Connecting to {0} using {1} ..' -f $SiteUrl, $local:upn)
                $null = Connect-PnPOnline -Url $SiteUrl -AccessToken $local:pnpToken -ErrorAction Stop
            }
            else {
                # Fallback: PnP interactive (device code / browser)
                Write-Host ('Connecting to {0} using {1} ..' -f $SiteUrl, $local:upn)
                $null = Connect-PnPOnline -Url $SiteUrl -Interactive -ErrorAction Stop
            }
            if (Get-Command -Name Get-PnPConnection -ErrorAction SilentlyContinue) {
                $local:pnpConnection = Get-PnPConnection -ErrorAction Stop
                if (-not $local:pnpConnection) {
                    return (New-Office365ConnectionStatus -Service 'PnP (SharePoint)' -StateKey 'ConnectedPnP' -Connected $false -ErrorMessage 'Connect-PnPOnline completed without an active PnP connection.')
                }
            }
        }
        catch {
            $script:myOffice365Services['PnPSiteUrl'] = ''
            return (New-Office365ConnectionStatus -Service 'PnP (SharePoint)' -StateKey 'ConnectedPnP' -Connected $false -ErrorMessage ('Cannot connect to PnP: {0}' -f $_.Exception.Message))
        }

        $script:myOffice365Services['PnPSiteUrl']    = $SiteUrl
        New-Office365ConnectionStatus -Service 'PnP (SharePoint)' -StateKey 'ConnectedPnP' -Connected $true -Detail $SiteUrl
    }
}
