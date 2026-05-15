function Connect-PowerBI {
    <#
    .SYNOPSIS
    Connects to the Power BI service using the cached identity or an interactive sign-in.

    .DESCRIPTION
    Imports the MicrosoftPowerBIMgmt.Profile module and calls Connect-PowerBIServiceAccount.
    When an MSAL token is available it is passed directly so no additional browser prompt
    is needed. Falls back to Connect-PowerBIServiceAccount's own interactive flow otherwise.

    .EXAMPLE
    Connect-PowerBI
    Connects to Power BI using the currently cached UPN or via an interactive sign-in.
    #>
    [CmdletBinding()]
    param()

    # Module guard — the Profile sub-module is sufficient for Connect-PowerBIServiceAccount
    foreach ($local:modName in @('MicrosoftPowerBIMgmt.Profile', 'MicrosoftPowerBIMgmt')) {
        if (-not (Get-Module -Name $local:modName -ErrorAction SilentlyContinue)) {
            Import-Module -Name $local:modName -ErrorAction SilentlyContinue
        }
        if (Get-Command -Name Connect-PowerBIServiceAccount -ErrorAction SilentlyContinue) { break }
    }

    if (-not (Get-Command -Name Connect-PowerBIServiceAccount -ErrorAction SilentlyContinue)) {
        Write-Error -Message 'Cannot connect to Power BI - module not installed or not loading. Install MicrosoftPowerBIMgmt.'
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

    # Acquire MSAL token for Power BI
    $local:pbiToken = Get-Office365AccessToken -Scope 'https://analysis.windows.net/powerbi/api/.default'

    if ($local:pbiToken) {
        Write-Host ('Connecting to Power BI using {0} ..' -f $local:upn)
        Connect-PowerBIServiceAccount -AccessToken $local:pbiToken
    }
    else {
        # Fallback: let the Power BI module run its own interactive flow
        Write-Host ('Connecting to Power BI using {0} ..' -f $local:upn)
        Connect-PowerBIServiceAccount
    }
    $script:myOffice365Services['ConnectedPowerBI'] = $true
}
