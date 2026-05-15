function Connect-Office365 {
    <#
    .SYNOPSIS
    Connects to Microsoft 365 services.

    .DESCRIPTION
    When called without -Service, connects to Exchange Online, Microsoft Teams,
    Security & Compliance, and SharePoint Online (the same behaviour as previous versions).

    Use -Service to connect to a specific subset of services.

    .PARAMETER Service
    One or more services to connect. Valid values:
    EXO, SCC, SPO, Teams, AIP, PowerApps, Exchange, Graph, PowerBI, PnP.
    When omitted, EXO, Teams, SCC, and SPO are connected.

    .EXAMPLE
    Connect-Office365
    Connects to Exchange Online, Teams, Security & Compliance, and SharePoint Online.

    .EXAMPLE
    Connect-Office365 -Service EXO, Graph
    Connects to Exchange Online and Microsoft Graph only.
    #>
    [CmdletBinding()]
    param(
        [ValidateSet('EXO', 'SCC', 'SPO', 'Teams', 'AIP', 'PowerApps', 'Exchange', 'Graph', 'PowerBI', 'PnP')]
        [string[]]$Service
    )

    $local:targets = if ($Service) { $Service } else {
        @('EXO', 'Teams', 'SCC', 'SPO')
    }

    $local:dispatch = @{
        EXO       = { Connect-EXO }
        SCC       = { Connect-SCC }
        SPO       = { Connect-SPO }
        Teams     = { Connect-MSTeams }
        AIP       = { Connect-AIP }
        PowerApps = { Connect-PowerApps }
        Exchange  = { Connect-Exchange }
        Graph     = { Connect-MG }
        PowerBI   = { Connect-PowerBI }
        PnP       = { Connect-PnP }
    }

    foreach ($local:svc in $local:targets) {
        & $local:dispatch[$local:svc]
    }
}

