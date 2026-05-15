function Get-Office365Session {
    <#
    .SYNOPSIS
    Shows the current identity and connection status for all supported services.

    .DESCRIPTION
    Returns two objects:
    - An identity snapshot (UPN, TenantID, Environment) from the module state.
    - A per-service connection status table derived from the Connected* flags set
      by the connect functions.

    For Exchange Online, additional detail is pulled from Get-ConnectionInformation
    (ExchangeOnlineManagement v3+) when available.

    .EXAMPLE
    Get-Office365Session
    Displays the active identity and which services are currently connected.
    #>
    [CmdletBinding()]
    param()

    # ── Identity ──────────────────────────────────────────────────────────────
    $local:upn = $script:myOffice365Services['Office365UPN']
    if (-not $local:upn -and $script:myOffice365Services['Office365Credential']) {
        $local:upn = $script:myOffice365Services['Office365Credential'].UserName
    }

    Write-Host ''
    Write-Host 'Identity'
    Write-Host '--------'
    [PSCustomObject][ordered]@{
        UPN         = if ($local:upn) { $local:upn } else { '(none)' }
        TenantID    = if ($script:myOffice365Services['TenantID']) { $script:myOffice365Services['TenantID'] } else { '(unknown)' }
        Environment = [string]$script:myOffice365Services['AzureEnvironmentName']
    } | Format-List

    # ── Service connections ───────────────────────────────────────────────────
    Write-Host 'Service Connections'
    Write-Host '-------------------'

    # Enrich EXO status via Get-ConnectionInformation when available (EOM v3+)
    $local:exoDetail = $null
    if (Get-Command -Name Get-ConnectionInformation -ErrorAction SilentlyContinue) {
        try {
            $local:connInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
            if ($local:connInfo) {
                $local:exoDetail = $local:connInfo | Select-Object -First 1 -ExpandProperty UserPrincipalName -ErrorAction SilentlyContinue
            }
        }
        catch { }
    }

    # Enrich Graph status via Get-MgContext when available (Microsoft.Graph SDK)
    $local:graphDetail = $null
    if (Get-Command -Name Get-MgContext -ErrorAction SilentlyContinue) {
        try {
            $local:mgCtx = Get-MgContext -ErrorAction SilentlyContinue
            if ($local:mgCtx) {
                $local:graphDetail = if ($local:mgCtx.Account) { $local:mgCtx.Account } else { $local:mgCtx.AppName }
            }
        }
        catch { }
    }

    $local:services = [ordered]@{
        'Exchange Online'              = $script:myOffice365Services['ConnectedEXO']
        'Security & Compliance'        = $script:myOffice365Services['ConnectedSCC']
        'SharePoint Online'            = $script:myOffice365Services['ConnectedSPO']
        'Microsoft Teams'              = $script:myOffice365Services['ConnectedTeams']
        'Azure Information Protection' = $script:myOffice365Services['ConnectedAIP']
        'PowerApps'                    = $script:myOffice365Services['ConnectedPowerApps']
        'Exchange On-Premises'         = $script:myOffice365Services['ConnectedExchange']
        'Microsoft Graph'              = $script:myOffice365Services['ConnectedGraph']
        'Power BI'                     = $script:myOffice365Services['ConnectedPowerBI']
        'PnP (SharePoint)'             = $script:myOffice365Services['ConnectedPnP']
    }

    $local:rows = foreach ($local:svcName in $local:services.Keys) {
        $local:connected = [bool]$local:services[$local:svcName]
        $local:detail    = ''
        if ($local:svcName -eq 'Exchange Online' -and $local:connected -and $local:exoDetail) {
            $local:detail = $local:exoDetail
        }
        if ($local:svcName -eq 'Microsoft Graph' -and $local:connected -and $local:graphDetail) {
            $local:detail = $local:graphDetail
        }
        if ($local:svcName -eq 'PnP (SharePoint)' -and $local:connected -and $script:myOffice365Services['PnPSiteUrl']) {
            $local:detail = $script:myOffice365Services['PnPSiteUrl']
        }
        [PSCustomObject][ordered]@{
            Service   = $local:svcName
            Connected = $local:connected
            Detail    = $local:detail
        }
    }
    $local:rows | Format-Table -AutoSize
}
