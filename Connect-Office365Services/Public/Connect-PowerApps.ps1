function Connect-PowerApps {
    if (-not (Get-Module -Name Microsoft.PowerApps.PowerShell)) {
        Import-Module -Name Microsoft.PowerApps.PowerShell -ErrorAction SilentlyContinue
    }
    if (-not (Get-Module -Name Microsoft.PowerApps.Administration.PowerShell)) {
        Import-Module -Name Microsoft.PowerApps.Administration.PowerShell -ErrorAction SilentlyContinue
    }
    if (Get-Command -Name Add-PowerAppsAccount -ErrorAction SilentlyContinue) {
        if ( -not $script:myOffice365Services['Office365UPN'] -and -not $script:myOffice365Services['Office365Credential']) {
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
        Write-Host ('Connecting to PowerApps using {0} ..' -f $local:upn)

        # Add-PowerAppsAccount accepts -AccessToken (plain string) to avoid a second browser prompt.
        $local:powerAppsToken = Get-Office365AccessToken -Scope 'https://service.powerapps.com/.default'
        if ($local:powerAppsToken) {
            Add-PowerAppsAccount -AccessToken $local:powerAppsToken -Username $local:upn
        }
        else {
            Add-PowerAppsAccount -Username $local:upn
        }
        $script:myOffice365Services['ConnectedPowerApps'] = $true
    }
    else {
        Write-Error -Message 'Cannot connect to PowerApps - problem loading module.'
    }
}
