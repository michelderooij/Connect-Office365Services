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
        [string]$SiteUrl
    )

    # Resolve site URL — parameter wins, then fall back to tenant root
    if (-not $SiteUrl) {
        $local:tenant = $script:myOffice365Services['Office365Tenant']
        if (-not $local:tenant) {
            Write-Error 'No -SiteUrl supplied and no tenant name is configured. Run Get-Office365Tenant or supply -SiteUrl explicitly.'
            return
        }
        $SiteUrl = 'https://{0}.sharepoint.com' -f $local:tenant
    }

    # Module guard
    if (-not (Get-Module -Name PnP.PowerShell -ErrorAction SilentlyContinue)) {
        Import-Module -Name PnP.PowerShell -ErrorAction SilentlyContinue
    }
    if (-not (Get-Command -Name Connect-PnPOnline -ErrorAction SilentlyContinue)) {
        Write-Error -Message 'Cannot connect via PnP - module not installed or not loading. Install PnP.PowerShell.'
        return
    }

    # Ensure we have an account cached (MSAL) or credentials (legacy)
    if (-not $script:myOffice365Services['Office365UPN'] -and -not $script:myOffice365Services['Office365Credential']) {
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

    # Scope token to the site's tenant host (e.g. https://contoso.sharepoint.com/.default)
    $local:siteUri     = [uri]$SiteUrl
    $local:tenantScope = '{0}://{1}/.default' -f $local:siteUri.Scheme, $local:siteUri.Host

    $local:pnpToken = Get-Office365AccessToken -Scope $local:tenantScope

    if ($local:pnpToken) {
        Write-Host ('Connecting to {0} using {1} ..' -f $SiteUrl, $local:upn)
        Connect-PnPOnline -Url $SiteUrl -AccessToken $local:pnpToken
    }
    else {
        # Fallback: PnP interactive (device code / browser)
        Write-Host ('Connecting to {0} using {1} ..' -f $SiteUrl, $local:upn)
        Connect-PnPOnline -Url $SiteUrl -Interactive
    }

    $script:myOffice365Services['ConnectedPnP']  = $true
    $script:myOffice365Services['PnPSiteUrl']    = $SiteUrl
}
