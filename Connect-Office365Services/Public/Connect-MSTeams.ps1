function Connect-MSTeams {
    <#
    .SYNOPSIS
    Connects to Microsoft Teams.
    .DESCRIPTION
    Explicit parameters are passed to the installed Connect-MicrosoftTeams command.
    Without parameters, the cached identity is used.
    #>
    [CmdletBinding()]
    param()
    dynamicparam {
        Get-Office365ConnectParameters -Module 'MicrosoftTeams' -Command 'Connect-MicrosoftTeams'
    }
    process {
        # Import MicrosoftTeams before any MSAL activity to ensure Teams loads its own
        # MSAL.Broker version first and avoids assembly conflicts.
        if (-not (Get-Module -Name MicrosoftTeams)) {
            Import-Module -Name MicrosoftTeams -ErrorAction SilentlyContinue
        }
        if (Get-Command -Name Connect-MicrosoftTeams -ErrorAction SilentlyContinue) {
            $local:teamsParams = Get-Office365ExplicitConnectParameters -BoundParameters $PSBoundParameters
            if ($local:teamsParams.Count -gt 0) {
                try {
                    Write-Host 'Connecting to Microsoft Teams using the specified options ..'
                    $null = Connect-MicrosoftTeams @local:teamsParams -ErrorAction Stop
                }
                catch {
                    return (New-Office365ConnectionStatus -Service 'Microsoft Teams' -StateKey 'ConnectedTeams' -Connected $false -ErrorMessage ('Cannot connect to Microsoft Teams: {0}' -f $_.Exception.Message))
                }
                return (New-Office365ConnectionStatus -Service 'Microsoft Teams' -StateKey 'ConnectedTeams' -Connected $true)
            }
            if ( -not $script:myOffice365Services['Office365UPN'] -and -not $script:myOffice365Services['Office365Credential']) {
                if ($script:myOffice365Services['NoAutoConnect']) {
                    return (New-Office365ConnectionStatus -Service 'Microsoft Teams' -StateKey 'ConnectedTeams' -Connected $false -ErrorMessage 'No credentials cached. Run Get-Office365Credential first or supply credentials explicitly.')
                }
                try {
                    $null = Get-Office365Credential
                }
                catch {
                    return (New-Office365ConnectionStatus -Service 'Microsoft Teams' -StateKey 'ConnectedTeams' -Connected $false -ErrorMessage ('Could not acquire credentials for Microsoft Teams: {0}' -f $_.Exception.Message))
                }
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
                $null = Connect-MicrosoftTeams @local:teamsParams -ErrorAction Stop
            }
            catch {
                if ($_.Exception.Message -match 'Microsoft\.Identity\.Client\.Broker') {
                    Write-Verbose ('MSAL.Broker assembly note (non-fatal): {0}' -f $_.Exception.Message)
                    return (New-Office365ConnectionStatus -Service 'Microsoft Teams' -StateKey 'ConnectedTeams' -Connected $false -ErrorMessage ('Microsoft Teams connection did not complete: {0}' -f $_.Exception.Message))
                }
                return (New-Office365ConnectionStatus -Service 'Microsoft Teams' -StateKey 'ConnectedTeams' -Connected $false -ErrorMessage ('Cannot connect to Microsoft Teams: {0}' -f $_.Exception.Message))
            }
            New-Office365ConnectionStatus -Service 'Microsoft Teams' -StateKey 'ConnectedTeams' -Connected $true -Detail $local:upn
        }
        else {
            New-Office365ConnectionStatus -Service 'Microsoft Teams' -StateKey 'ConnectedTeams' -Connected $false -ErrorMessage 'Cannot connect to Microsoft Teams - module not installed or not loading.'
        }
    }
}
