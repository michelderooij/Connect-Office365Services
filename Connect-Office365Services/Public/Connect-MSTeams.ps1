function Connect-MSTeams {
    # Import MicrosoftTeams before any MSAL activity to ensure Teams loads its own
    # MSAL.Broker version first and avoids assembly conflicts.
    if (-not (Get-Module -Name MicrosoftTeams)) {
        Import-Module -Name MicrosoftTeams -ErrorAction SilentlyContinue
    }
    if (Get-Command -Name Connect-MicrosoftTeams -ErrorAction SilentlyContinue) {
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
        Write-Host ('Connecting to Microsoft Teams using {0} ..' -f $local:upn)

        $local:teamsParams = @{ AccountId = $local:upn }
        if ($script:myOffice365Services['TenantID']) {
            $local:teamsParams['TenantId'] = $script:myOffice365Services['TenantID']
        }
        # MSAL.Broker assembly conflicts are non-fatal: Teams still authenticates via
        # its own interactive flow. Suppress those specific errors; re-throw anything else.
        try {
            Connect-MicrosoftTeams @local:teamsParams -ErrorAction Stop
        }
        catch {
            if ($_.Exception.Message -match 'Microsoft\.Identity\.Client\.Broker') {
                # Assembly version mismatch is a warning-level issue; Connect-MicrosoftTeams
                # will retry internally without the Broker.
                Write-Verbose ('MSAL.Broker assembly note (non-fatal): {0}' -f $_.Exception.Message)
            }
            else {
                throw
            }
        }
        $script:myOffice365Services['ConnectedTeams'] = $true
    }
    else {
        Write-Error -Message 'Cannot connect to Microsoft Teams - module not installed or not loading.'
    }
}
