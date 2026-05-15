function Connect-MG {
    <#
    .SYNOPSIS
    Connects to Microsoft Graph using the cached identity or an interactive sign-in.

    .DESCRIPTION
    Loads the Microsoft Graph Authentication module (Microsoft.Graph.Authentication,
    Microsoft.Graph, or Microsoft.Entra — whichever is installed), acquires an OAuth2
    token via MSAL.NET, and calls Connect-MgGraph with that token.

    When no MSAL token is available (e.g. MSAL.NET is not installed) the function falls
    back to Connect-MgGraph's own interactive flow.

    .EXAMPLE
    Connect-MG
    Connects to Microsoft Graph using the currently cached UPN or via an interactive sign-in.
    #>
    [CmdletBinding()]
    param()

    # Module guard — prefer Authentication sub-module for minimal footprint
    foreach ($local:modName in @('Microsoft.Graph.Authentication', 'Microsoft.Graph', 'Microsoft.Entra')) {
        if (-not (Get-Module -Name $local:modName -ErrorAction SilentlyContinue)) {
            Import-Module -Name $local:modName -ErrorAction SilentlyContinue
        }
        if (Get-Command -Name Connect-MgGraph -ErrorAction SilentlyContinue) { break }
    }

    if (-not (Get-Command -Name Connect-MgGraph -ErrorAction SilentlyContinue)) {
        Write-Error -Message 'Cannot connect to Microsoft Graph - module not installed or not loading. Install Microsoft.Graph.Authentication or Microsoft.Graph.'
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

    # Acquire MSAL token for Graph
    $local:graphToken = Get-Office365AccessToken -Scope 'https://graph.microsoft.com/.default'

    if ($local:graphToken) {
        Write-Host ('Connecting to Microsoft Graph using {0} ..' -f $local:upn)
        # Graph SDK v2 requires AccessToken as SecureString
        $local:secureToken = ConvertTo-SecureString -String $local:graphToken -AsPlainText -Force
        Connect-MgGraph -AccessToken $local:secureToken -NoWelcome
    }
    else {
        # Fallback: let the Graph SDK run its own interactive flow
        Write-Host ('Connecting to Microsoft Graph using {0} ..' -f $local:upn)
        Connect-MgGraph -Scopes 'https://graph.microsoft.com/.default' -NoWelcome
    }
    $script:myOffice365Services['ConnectedGraph'] = $true
}
