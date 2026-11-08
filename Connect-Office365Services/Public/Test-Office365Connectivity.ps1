function Test-Office365Connectivity {
    <#
    .SYNOPSIS
    Tests reachability and readiness of key Microsoft 365 service endpoints for the active environment.

    .DESCRIPTION
    Sends a lightweight HTTP GET to each service endpoint and reports whether it
    is reachable and ready for use. Any successful HTTP response indicates the
    host is alive and responding. 401/403 responses are treated as reachable and
    auth-required rather than unreachable. Timeouts and DNS failures are reported
    as unreachable.

    Useful for diagnosing proxy, firewall, DNS, or tenant configuration issues
    before attempting to connect to a service.

    The SharePoint Online endpoint is skipped when no tenant is configured in
    the current session.

    .EXAMPLE
    Test-Office365Connectivity
    Tests all endpoints for the currently configured environment and displays results.

    .EXAMPLE
    Test-Office365Connectivity | Where-Object { -not $_.Ready }
    Displays endpoints that are reachable but not ready for service use.
    #>
    [CmdletBinding()]
    param()

    # ── Resolve environment-specific endpoints ────────────────────────────────
    $local:authBase = $script:myOffice365Services['AzureADAuthorizationEndpointUri']
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

    foreach ($local:svcName in $local:endpoints.Keys) {
        $local:url = $local:endpoints[$local:svcName]

        if (-not $local:url) {
            [PSCustomObject][ordered]@{
                Service    = $local:svcName
                Endpoint   = '(skipped — tenant name not configured)'
                Reachable  = $null
                Ready      = $null
                StatusCode = $null
                Details    = 'Set tenant via Get-Office365Credential or Get-Office365Tenant'
            }
            continue
        }

        $local:reachable = $false
        $local:ready = $false
        $local:statusCode = $null
        $local:details = ''

        try {
            $local:invokeParams = @{
                Uri         = $local:url
                Method      = 'Get'
                TimeoutSec  = 10
                ErrorAction = 'Stop'
            }

            # PowerShell 5.1 supports -UseBasicParsing; PowerShell 7 does not.
            if ($PSVersionTable.PSEdition -eq 'Desktop' -or $PSVersionTable.PSVersion.Major -lt 6) {
                $local:invokeParams['UseBasicParsing'] = $true
            }

            $local:response = Invoke-WebRequest @local:invokeParams
            $local:statusCode = [int]$local:response.StatusCode
            $local:reachable = $true

            if ($local:statusCode -ge 200 -and $local:statusCode -lt 500) {
                $local:ready = $true
                if ($local:statusCode -in 401, 403) {
                    $local:details = 'Service is reachable; authentication or authorization is required.'
                }
            }
            elseif ($local:statusCode -ge 500) {
                $local:ready = $false
                $local:details = 'Service responded with HTTP {0}; service readiness is degraded.' -f $local:statusCode
            }
        }
        catch [System.Net.WebException] {
            $local:webEx = $_.Exception
            if ($local:webEx.Response) {
                $local:statusCode = [int]$local:webEx.Response.StatusCode
                $local:reachable = $true

                if ($local:statusCode -in 401, 403) {
                    $local:ready = $true
                    $local:details = 'Service is reachable; authentication or authorization is required.'
                }
                elseif ($local:statusCode -ge 500) {
                    $local:ready = $false
                    $local:details = 'Service responded with HTTP {0}; service readiness is degraded.' -f $local:statusCode
                }
                else {
                    $local:ready = $true
                }
            }
            else {
                $local:details = $local:webEx.Message
            }
        }
        catch [Microsoft.PowerShell.Commands.HttpResponseException] {
            $local:statusCode = [int]$_.Exception.Response.StatusCode
            $local:reachable = $true

            if ($local:statusCode -in 401, 403) {
                $local:ready = $true
                $local:details = 'Service is reachable; authentication or authorization is required.'
            }
            elseif ($local:statusCode -ge 500) {
                $local:ready = $false
                $local:details = 'Service responded with HTTP {0}; service readiness is degraded.' -f $local:statusCode
            }
            else {
                $local:ready = $true
            }
        }
        catch {
            $local:details = $_.Exception.Message
        }

        [PSCustomObject][ordered]@{
            Service    = $local:svcName
            Endpoint   = $local:url
            Reachable  = $local:reachable
            Ready      = $local:ready
            StatusCode = $local:statusCode
            Details    = $local:details
        }
    }
}
