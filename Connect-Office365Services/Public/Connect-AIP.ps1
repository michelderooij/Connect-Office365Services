function Connect-AIP {
    <#
    .SYNOPSIS
    Connects to Azure Information Protection.
    .DESCRIPTION
    Explicit parameters are passed to the installed Connect-AipService command.
    Without parameters, the cached identity is used.
    #>
    [CmdletBinding()]
    param()
    dynamicparam {
        Get-Office365ConnectParameters -Module 'AIPService' -Command 'Connect-AipService'
    }
    process {
        if (-not (Get-Module -Name AIPService)) {
            Import-Module -Name AIPService -ErrorAction SilentlyContinue
        }
        if (Get-Command -Name Connect-AipService -ErrorAction SilentlyContinue) {
            $local:connectParams = Get-Office365ExplicitConnectParameters -BoundParameters $PSBoundParameters
            if ($local:connectParams.Count -gt 0) {
                try {
                    Write-Host 'Connecting to Azure Information Protection using the specified options ..'
                    $null = Connect-AipService @local:connectParams -ErrorAction Stop
                }
                catch {
                    return (New-Office365ConnectionStatus -Service 'Azure Information Protection' -StateKey 'ConnectedAIP' -Connected $false -ErrorMessage ('Cannot connect to Azure Information Protection: {0}' -f $_.Exception.Message))
                }
                return (New-Office365ConnectionStatus -Service 'Azure Information Protection' -StateKey 'ConnectedAIP' -Connected $true)
            }
            # Ensure we have an account cached (MSAL) or credentials (legacy)
            if ( -not $script:myOffice365Services['Office365UPN'] -and -not $script:myOffice365Services['Office365Credential']) {
                if ($script:myOffice365Services['NoAutoConnect']) {
                    return (New-Office365ConnectionStatus -Service 'Azure Information Protection' -StateKey 'ConnectedAIP' -Connected $false -ErrorMessage 'No credentials cached. Run Get-Office365Credential first or supply credentials explicitly.')
                }
                try {
                    $null = Get-Office365Credential
                }
                catch {
                    return (New-Office365ConnectionStatus -Service 'Azure Information Protection' -StateKey 'ConnectedAIP' -Connected $false -ErrorMessage ('Could not acquire credentials for Azure Information Protection: {0}' -f $_.Exception.Message))
                }
            }

            try {
                # Modern auth: acquire AADRM-scoped token and pass via -AccessToken (plain String)
                $local:aipToken = Get-Office365AccessToken -Scope 'https://api.aadrm.com/.default'
                if ($local:aipToken) {
                    $local:displayName = $script:myOffice365Services['Office365UPN']
                    Write-Host ('Connecting to Azure Information Protection using {0}' -f $local:displayName)
                    $null = Connect-AipService -AccessToken $local:aipToken -TenantId $script:myOffice365Services['TenantID'] -ErrorAction Stop
                }
                elseif ($script:myOffice365Services['Office365Credential']) {
                    # Legacy PSCredential fallback
                    Write-Host ('Connecting to Azure Information Protection using {0}' -f $script:myOffice365Services['Office365Credential'].UserName)
                    $null = Connect-AipService -Credential $script:myOffice365Services['Office365Credential'] -ErrorAction Stop
                }
                else {
                    Write-Host 'Connecting to Azure Information Protection ..'
                    $null = Connect-AipService -ErrorAction Stop
                }
            }
            catch {
                return (New-Office365ConnectionStatus -Service 'Azure Information Protection' -StateKey 'ConnectedAIP' -Connected $false -ErrorMessage ('Cannot connect to Azure Information Protection: {0}' -f $_.Exception.Message))
            }
            New-Office365ConnectionStatus -Service 'Azure Information Protection' -StateKey 'ConnectedAIP' -Connected $true -Detail $script:myOffice365Services['Office365UPN']
        }
        else {
            New-Office365ConnectionStatus -Service 'Azure Information Protection' -StateKey 'ConnectedAIP' -Connected $false -ErrorMessage 'Cannot connect to Azure Information Protection - module not installed or not loading.'
        }
    }
}
