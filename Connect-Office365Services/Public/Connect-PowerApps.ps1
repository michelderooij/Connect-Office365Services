function Connect-PowerApps {
    <#
    .SYNOPSIS
    Connects to Power Apps.
    .DESCRIPTION
    Explicit parameters are passed to the installed Add-PowerAppsAccount command.
    Without parameters, the cached identity is used.
    #>
    [CmdletBinding()]
    param()
    dynamicparam {
        Get-Office365ConnectParameters -Module @('Microsoft.PowerApps.PowerShell', 'Microsoft.PowerApps.Administration.PowerShell') -Command 'Add-PowerAppsAccount'
    }
    process {
        if (-not (Get-Module -Name Microsoft.PowerApps.PowerShell)) {
            Import-Module -Name Microsoft.PowerApps.PowerShell -ErrorAction SilentlyContinue
        }
        if (-not (Get-Module -Name Microsoft.PowerApps.Administration.PowerShell)) {
            Import-Module -Name Microsoft.PowerApps.Administration.PowerShell -ErrorAction SilentlyContinue
        }
        if (Get-Command -Name Add-PowerAppsAccount -ErrorAction SilentlyContinue) {
            $local:connectParams = Get-Office365ExplicitConnectParameters -BoundParameters $PSBoundParameters
            if ($local:connectParams.Count -gt 0) {
                try {
                    Write-Host 'Connecting to PowerApps using the specified options ..'
                    $null = Add-PowerAppsAccount @local:connectParams -ErrorAction Stop
                }
                catch {
                    return (New-Office365ConnectionStatus -Service 'PowerApps' -StateKey 'ConnectedPowerApps' -Connected $false -ErrorMessage ('Cannot connect to PowerApps: {0}' -f $_.Exception.Message))
                }
                return (New-Office365ConnectionStatus -Service 'PowerApps' -StateKey 'ConnectedPowerApps' -Connected $true)
            }
            if ( -not $script:myOffice365Services['Office365UPN'] -and -not $script:myOffice365Services['Office365Credential']) {
                if ($script:myOffice365Services['NoAutoConnect']) {
                    return (New-Office365ConnectionStatus -Service 'PowerApps' -StateKey 'ConnectedPowerApps' -Connected $false -ErrorMessage 'No credentials cached. Run Get-Office365Credential first or supply credentials explicitly.')
                }
                try {
                    $null = Get-Office365Credential
                }
                catch {
                    return (New-Office365ConnectionStatus -Service 'PowerApps' -StateKey 'ConnectedPowerApps' -Connected $false -ErrorMessage ('Could not acquire credentials for PowerApps: {0}' -f $_.Exception.Message))
                }
            }
            $local:upn = if ($script:myOffice365Services['Office365UPN']) {
                $script:myOffice365Services['Office365UPN']
            }
            else {
                $script:myOffice365Services['Office365Credential'].UserName
            }
            Write-Host ('Connecting to PowerApps using {0} ..' -f $local:upn)

            # Add-PowerAppsAccount accepts -AccessToken (plain string) to avoid a second browser prompt.
            try {
                $local:powerAppsToken = Get-Office365AccessToken -Scope 'https://service.powerapps.com/.default'
                if ($local:powerAppsToken) {
                    $null = Add-PowerAppsAccount -AccessToken $local:powerAppsToken -Username $local:upn -ErrorAction Stop
                }
                else {
                    $null = Add-PowerAppsAccount -Username $local:upn -ErrorAction Stop
                }
            }
            catch {
                return (New-Office365ConnectionStatus -Service 'PowerApps' -StateKey 'ConnectedPowerApps' -Connected $false -ErrorMessage ('Cannot connect to PowerApps: {0}' -f $_.Exception.Message))
            }
            New-Office365ConnectionStatus -Service 'PowerApps' -StateKey 'ConnectedPowerApps' -Connected $true -Detail $local:upn
        }
        else {
            New-Office365ConnectionStatus -Service 'PowerApps' -StateKey 'ConnectedPowerApps' -Connected $false -ErrorMessage 'Cannot connect to PowerApps - problem loading module.'
        }
    }
}
