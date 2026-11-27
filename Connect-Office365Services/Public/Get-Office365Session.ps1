function Get-Office365Session {
    <#
    .SYNOPSIS
    Returns the current identity and structured connection status for all supported services.

    .DESCRIPTION
    Returns one object with an Identity property and a Connections array. Live
    vendor session state is used where the service exposes a reliable query.

    .EXAMPLE
    Get-Office365Session
    Returns structured status for the active identity and service connections.
    #>
    [CmdletBinding()]
    param()

    $local:upn = $script:myOffice365Services['Office365UPN']
    if (-not $local:upn -and $script:myOffice365Services['Office365Credential']) {
        $local:upn = $script:myOffice365Services['Office365Credential'].UserName
    }

    $local:identity = [PSCustomObject][ordered]@{
        UPN         = $local:upn
        TenantID    = $script:myOffice365Services['TenantID']
        Environment = [string]$script:myOffice365Services['AzureEnvironmentName']
    }

    $local:exoConnections = @()
    $local:exoProbeError = $null
    $local:hasExchangeQuery = [bool](Get-Command -Name Get-ConnectionInformation -ErrorAction SilentlyContinue)
    if ($local:hasExchangeQuery) {
        try {
            $local:exoConnections = @(Get-ConnectionInformation -ErrorAction Stop)
        }
        catch {
            $local:exoProbeError = 'Could not query Exchange Online connection state: {0}' -f $_.Exception.Message
        }
    }

    $local:exoInfo = @($local:exoConnections | Where-Object {
        $_.State -eq 'Connected' -and $_.ConnectionUri -notmatch 'compliance|protection'
    })
    if ($local:hasExchangeQuery) {
        $local:exoDetail = if ($local:exoInfo) { $local:exoInfo[0].UserPrincipalName } else { $null }
        $local:exoConnected = [bool]$local:exoInfo
        if ($local:exoProbeError) {
            $local:exoStatus = New-Office365ConnectionStatus -Service 'Exchange Online' -StateKey 'ConnectedEXO' -Connected $false -ErrorMessage $local:exoProbeError
        }
        else {
            $local:exoStatus = New-Office365ConnectionStatus -Service 'Exchange Online' -StateKey 'ConnectedEXO' -Connected $local:exoConnected -Detail $local:exoDetail -IncludeSuccess
        }
    }
    else {
        $local:exoSession = $script:myOffice365Services['Session365']
        $local:exoConnected = if ($local:exoSession -and $local:exoSession.PSObject.Properties['State']) {
            $local:exoSession.State -eq 'Opened'
        }
        else {
            [bool]$script:myOffice365Services['ConnectedEXO']
        }
        $local:exoStatus = [PSCustomObject][ordered]@{
            Service   = 'Exchange Online'
            Connected = [bool]$local:exoConnected
            Detail    = $null
            Error     = $null
        }
    }

    $local:sccConnections = @($local:exoConnections | Where-Object {
        $_.State -eq 'Connected' -and $_.ConnectionUri -match 'compliance|protection'
    })
    $local:sccSession = $script:myOffice365Services['SessionCC']
    if (-not $local:hasExchangeQuery -and $local:sccSession -and $local:sccSession.State -eq 'Opened') {
        $local:sccConnections = @($local:sccSession)
    }
    if ($local:hasExchangeQuery) {
        $local:sccDetail = if ($local:sccConnections) { $local:sccConnections[0].UserPrincipalName } else { $null }
        if ($local:exoProbeError) {
            $local:sccStatus = New-Office365ConnectionStatus -Service 'Security & Compliance' -StateKey 'ConnectedSCC' -Connected $false -ErrorMessage $local:exoProbeError
        }
        else {
            $local:sccStatus = New-Office365ConnectionStatus -Service 'Security & Compliance' -StateKey 'ConnectedSCC' -Connected ([bool]$local:sccConnections) -Detail $local:sccDetail -IncludeSuccess
        }
    }
    else {
        $local:sccConnected = if ($local:sccSession -and $local:sccSession.PSObject.Properties['State']) {
            $local:sccSession.State -eq 'Opened'
        }
        else {
            [bool]$script:myOffice365Services['ConnectedSCC']
        }
        $local:sccStatus = [PSCustomObject][ordered]@{
            Service   = 'Security & Compliance'
            Connected = [bool]$local:sccConnected
            Detail    = $null
            Error     = $null
        }
    }

    $local:graphStatus = $null
    if (Get-Command -Name Get-MgContext -ErrorAction SilentlyContinue) {
        try {
            $local:mgCtx = Get-MgContext -ErrorAction Stop
            $local:graphDetail = if ($local:mgCtx.Account) { $local:mgCtx.Account } else { $local:mgCtx.AppName }
            $local:graphStatus = New-Office365ConnectionStatus -Service 'Microsoft Graph' -StateKey 'ConnectedGraph' -Connected ([bool]$local:mgCtx) -Detail $local:graphDetail -IncludeSuccess
        }
        catch {
            $local:graphError = 'Could not query Microsoft Graph connection state: {0}' -f $_.Exception.Message
            $local:graphStatus = New-Office365ConnectionStatus -Service 'Microsoft Graph' -StateKey 'ConnectedGraph' -Connected $false -ErrorMessage $local:graphError
        }
    }
    else {
        $local:graphStatus = [PSCustomObject][ordered]@{
            Service   = 'Microsoft Graph'
            Connected = [bool]$script:myOffice365Services['ConnectedGraph']
            Detail    = $null
            Error     = $null
        }
    }

    $local:exchangeSession = $script:myOffice365Services['SessionExchange']
    $local:exchangeConnected = $exchangeSession -and $exchangeSession.State -eq 'Opened'
    $local:exchangeDetail = if ($local:exchangeConnected) { $script:myOffice365Services['ExchangeOnPremisesFQDN'] } else { $null }
    $local:exchangeStatus = New-Office365ConnectionStatus -Service 'Exchange On-Premises' -StateKey 'ConnectedExchange' -Connected ([bool]$local:exchangeConnected) -Detail $local:exchangeDetail -IncludeSuccess

    $local:pnpStatus = $null
    if (Get-Command -Name Get-PnPConnection -ErrorAction SilentlyContinue) {
        try {
            $local:pnpConnection = Get-PnPConnection -ErrorAction Stop
            $local:pnpConnected = $null -ne $local:pnpConnection
            $local:pnpDetail = if ($local:pnpConnected -and $local:pnpConnection.Url) { $local:pnpConnection.Url } else { $null }
            $local:pnpStatus = New-Office365ConnectionStatus -Service 'PnP (SharePoint)' -StateKey 'ConnectedPnP' -Connected $local:pnpConnected -Detail $local:pnpDetail -IncludeSuccess
        }
        catch {
            if ($script:myOffice365Services['ConnectedPnP']) {
                $local:pnpError = 'Could not query PnP connection state: {0}' -f $_.Exception.Message
                $local:pnpStatus = New-Office365ConnectionStatus -Service 'PnP (SharePoint)' -StateKey 'ConnectedPnP' -Connected $false -ErrorMessage $local:pnpError
            }
            else {
                $local:pnpStatus = New-Office365ConnectionStatus -Service 'PnP (SharePoint)' -StateKey 'ConnectedPnP' -Connected $false
            }
        }
    }
    else {
        $local:pnpStatus = [PSCustomObject][ordered]@{
            Service   = 'PnP (SharePoint)'
            Connected = [bool]$script:myOffice365Services['ConnectedPnP']
            Detail    = if ($script:myOffice365Services['ConnectedPnP']) { $script:myOffice365Services['PnPSiteUrl'] } else { $null }
            Error     = $null
        }
    }

    $local:connections = @(
        $local:exoStatus
        $local:sccStatus
        [PSCustomObject][ordered]@{
            Service   = 'SharePoint Online'
            Connected = [bool]$script:myOffice365Services['ConnectedSPO']
            Detail    = $null
            Error     = $null
        }
        [PSCustomObject][ordered]@{
            Service   = 'Microsoft Teams'
            Connected = [bool]$script:myOffice365Services['ConnectedTeams']
            Detail    = $null
            Error     = $null
        }
        [PSCustomObject][ordered]@{
            Service   = 'Azure Information Protection'
            Connected = [bool]$script:myOffice365Services['ConnectedAIP']
            Detail    = $null
            Error     = $null
        }
        [PSCustomObject][ordered]@{
            Service   = 'PowerApps'
            Connected = [bool]$script:myOffice365Services['ConnectedPowerApps']
            Detail    = $null
            Error     = $null
        }
        $local:exchangeStatus
        $local:graphStatus
        [PSCustomObject][ordered]@{
            Service   = 'Power BI'
            Connected = [bool]$script:myOffice365Services['ConnectedPowerBI']
            Detail    = $null
            Error     = $null
        }
        $local:pnpStatus
    )

    [PSCustomObject][ordered]@{
        Identity    = $local:identity
        Connections = $local:connections
    }
}
