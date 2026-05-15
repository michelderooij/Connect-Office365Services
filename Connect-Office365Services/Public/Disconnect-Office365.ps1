function Disconnect-Office365 {
    <#
    .SYNOPSIS
    Disconnects one or more Microsoft 365 / Office 365 service sessions.

    .DESCRIPTION
    When called without -Service, disconnects all services that have an active
    session. Use -Service to target specific services only.

    The optional -ClearCredentials switch removes the cached account information
    (UPN, PSCredential, and MSAL token cache) from the module state.

    .PARAMETER Service
    One or more services to disconnect. Valid values:
    EXO, SCC, SPO, Teams, AIP, PowerApps, Exchange.
    When omitted, all services are disconnected.

    .PARAMETER ClearCredentials
    When specified, clears the cached credentials and MSAL token state from the
    module after disconnecting.

    .EXAMPLE
    Disconnect-Office365
    Disconnects all connected services.

    .EXAMPLE
    Disconnect-Office365 -Service EXO, SCC
    Disconnects Exchange Online and Security & Compliance only.

    .EXAMPLE
    Disconnect-Office365 -ClearCredentials
    Disconnects all services and clears the cached credentials.
    #>
    [CmdletBinding()]
    param(
        [ValidateSet('EXO', 'SCC', 'SPO', 'Teams', 'AIP', 'PowerApps', 'Exchange', 'Graph', 'PowerBI', 'PnP')]
        [string[]]$Service,

        [switch]$ClearCredentials
    )

    # When no services specified, target all
    $local:targets = if ($Service) { $Service } else {
        @('EXO', 'SCC', 'SPO', 'Teams', 'AIP', 'PowerApps', 'Exchange', 'Graph', 'PowerBI', 'PnP')
    }

    foreach ($local:svc in $local:targets) {
        switch ($local:svc) {
            'EXO' {
                if (Get-Command -Name Disconnect-ExchangeOnline -ErrorAction SilentlyContinue) {
                    Write-Host 'Disconnecting from Exchange Online ..'
                    try {
                        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
                    }
                    catch { Write-Verbose ('Disconnect-ExchangeOnline: {0}' -f $_) }
                }
                $script:myOffice365Services['ConnectedEXO'] = $false
            }
            'SCC' {
                if (Get-Command -Name Disconnect-ExchangeOnline -ErrorAction SilentlyContinue) {
                    Write-Host 'Disconnecting from Security & Compliance Center ..'
                    try {
                        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
                    }
                    catch { Write-Verbose ('Disconnect-ExchangeOnline (SCC): {0}' -f $_) }
                }
                $script:myOffice365Services['ConnectedSCC'] = $false
            }
            'SPO' {
                if (Get-Command -Name Disconnect-SPOService -ErrorAction SilentlyContinue) {
                    Write-Host 'Disconnecting from SharePoint Online ..'
                    try {
                        Disconnect-SPOService -ErrorAction SilentlyContinue
                    }
                    catch { Write-Verbose ('Disconnect-SPOService: {0}' -f $_) }
                }
                $script:myOffice365Services['ConnectedSPO'] = $false
            }
            'Teams' {
                if (Get-Command -Name Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue) {
                    Write-Host 'Disconnecting from Microsoft Teams ..'
                    try {
                        Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue
                    }
                    catch { Write-Verbose ('Disconnect-MicrosoftTeams: {0}' -f $_) }
                }
                $script:myOffice365Services['ConnectedTeams'] = $false
            }
            'AIP' {
                if (Get-Command -Name Disconnect-AipService -ErrorAction SilentlyContinue) {
                    Write-Host 'Disconnecting from Azure Information Protection ..'
                    try {
                        Disconnect-AipService -ErrorAction SilentlyContinue
                    }
                    catch { Write-Verbose ('Disconnect-AipService: {0}' -f $_) }
                }
                $script:myOffice365Services['ConnectedAIP'] = $false
            }
            'PowerApps' {
                if (Get-Command -Name Remove-PowerAppsAccount -ErrorAction SilentlyContinue) {
                    Write-Host 'Disconnecting from PowerApps ..'
                    try {
                        Remove-PowerAppsAccount -ErrorAction SilentlyContinue
                    }
                    catch { Write-Verbose ('Remove-PowerAppsAccount: {0}' -f $_) }
                }
                $script:myOffice365Services['ConnectedPowerApps'] = $false
            }
            'Exchange' {
                $local:session = $script:myOffice365Services['SessionExchange']
                if ($local:session) {
                    Write-Host 'Disconnecting from Exchange On-Premises ..'
                    try {
                        Remove-PSSession -Session $local:session -ErrorAction SilentlyContinue
                    }
                    catch { Write-Verbose ('Remove-PSSession (Exchange): {0}' -f $_) }
                    $script:myOffice365Services['SessionExchange'] = $null
                }
                $script:myOffice365Services['ConnectedExchange'] = $false
            }
            'Graph' {
                if (Get-Command -Name Disconnect-MgGraph -ErrorAction SilentlyContinue) {
                    Write-Host 'Disconnecting from Microsoft Graph ..'
                    try {
                        Disconnect-MgGraph -ErrorAction SilentlyContinue
                    }
                    catch { Write-Verbose ('Disconnect-MgGraph: {0}' -f $_) }
                }
                $script:myOffice365Services['ConnectedGraph'] = $false
            }
            'PowerBI' {
                if (Get-Command -Name Disconnect-PowerBIServiceAccount -ErrorAction SilentlyContinue) {
                    Write-Host 'Disconnecting from Power BI ..'
                    try {
                        Disconnect-PowerBIServiceAccount -ErrorAction SilentlyContinue
                    }
                    catch { Write-Verbose ('Disconnect-PowerBIServiceAccount: {0}' -f $_) }
                }
                $script:myOffice365Services['ConnectedPowerBI'] = $false
            }
            'PnP' {
                if (Get-Command -Name Disconnect-PnPOnline -ErrorAction SilentlyContinue) {
                    Write-Host 'Disconnecting from PnP (SharePoint) ..'
                    try {
                        Disconnect-PnPOnline -ErrorAction SilentlyContinue
                    }
                    catch { Write-Verbose ('Disconnect-PnPOnline: {0}' -f $_) }
                }
                $script:myOffice365Services['ConnectedPnP']  = $false
                $script:myOffice365Services['PnPSiteUrl']    = ''
            }
        }
    }

    if ($ClearCredentials) {
        Write-Host 'Clearing cached credentials ..'
        $script:myOffice365Services['Office365UPN']          = ''
        $script:myOffice365Services['Office365Credential']   = $null
        $script:myOffice365Services['MsalAccount']           = $null
        $script:myOffice365Services['MsalApp']               = $null
        $script:myOffice365Services['MsalTokenCacheBytes']   = $null
    }
}
