function Test-Office365Connectivity {
    <#
    .SYNOPSIS
    Tests reachability of key Microsoft 365 service endpoints for the active environment.

    .DESCRIPTION
    Sends a lightweight HTTP GET to each service endpoint and reports whether it
    is reachable. An HTTP response of any status code (2xx–4xx) is treated as
    reachable — it proves the host is alive and responding. Timeouts and DNS
    failures are reported as unreachable.

    Useful for diagnosing proxy, firewall, or DNS issues before attempting to
    connect to a service.

    The SharePoint Online endpoint is skipped when no tenant is configured in
    the current session.

    .EXAMPLE
    Test-Office365Connectivity
    Tests all endpoints for the currently configured environment and displays results.

    .EXAMPLE
    Test-Office365Connectivity | Where-Object { -not $_.Reachable }
    Displays only the unreachable endpoints.
    #>
    [CmdletBinding()]
    param()

    # ── Resolve environment-specific endpoints ────────────────────────────────
    $local:authBase = $script:myOffice365Services['AzureADAuthorizationEndpointUri']
    # Strip to just the scheme+host for a clean HEAD-style test
    $local:authEndpoint = if ($local:authBase) {
        try { ([uri]$local:authBase).GetLeftPart([UriPartial]::Authority) } catch { $local:authBase }
    }
    else { 'https://login.microsoftonline.com' }

    $local:exoRaw = $script:myOffice365Services['ConnectionEndpointUri']
    $local:exoEndpoint = if ($local:exoRaw) {
        try { ([uri]$local:exoRaw).GetLeftPart([UriPartial]::Authority) } catch { $local:exoRaw }
    }
    else { 'https://outlook.office365.com' }

    $local:tenant = $script:myOffice365Services['Office365Tenant']
    $local:spoEndpoint = if ($local:tenant) {
        'https://{0}.sharepoint.com' -f $local:tenant
    }
    else { $null }

    $local:endpoints = [ordered]@{
        'Entra ID / Auth'              = $local:authEndpoint
        'Microsoft Graph'              = 'https://graph.microsoft.com'
        'Exchange Online'              = $local:exoEndpoint
        'SharePoint Online'            = $local:spoEndpoint
        'Microsoft Teams'              = 'https://teams.microsoft.com'
        'Azure Information Protection' = 'https://api.aadrm.com'
    }

    # ── Test each endpoint ────────────────────────────────────────────────────
    foreach ($local:svcName in $local:endpoints.Keys) {
        $local:url = $local:endpoints[$local:svcName]

        if (-not $local:url) {
            [PSCustomObject][ordered]@{
                Service    = $local:svcName
                Endpoint   = '(skipped — tenant name not configured)'
                Reachable  = $null
                StatusCode = $null
                Details    = 'Set tenant via Get-Office365Credential or Get-Office365Tenant'
            }
            continue
        }

        $local:reachable = $false
        $local:statusCode = $null
        $local:details = ''

        try {
            $local:response = Invoke-WebRequest -Uri $local:url -Method Get `
                -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
            $local:statusCode = $local:response.StatusCode
            $local:reachable = $true
        }
        catch [System.Net.WebException] {
            # PS5: a WebException with a response still means the host answered
            $local:webEx = $_.Exception
            if ($local:webEx.Response) {
                $local:statusCode = [int]$local:webEx.Response.StatusCode
                $local:reachable = $true
            }
            else {
                $local:details = $local:webEx.Message
            }
        }
        catch [Microsoft.PowerShell.Commands.HttpResponseException] {
            # PS7: non-2xx responses throw HttpResponseException — the host answered
            $local:statusCode = [int]$_.Exception.Response.StatusCode
            $local:reachable = $true
        }
        catch {
            $local:details = $_.Exception.Message
        }

        [PSCustomObject][ordered]@{
            Service    = $local:svcName
            Endpoint   = $local:url
            Reachable  = $local:reachable
            StatusCode = $local:statusCode
            Details    = $local:details
        }
    }
}
