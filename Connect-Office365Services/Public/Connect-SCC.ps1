function Connect-SCC {
    <#
    .SYNOPSIS
    Connects to Security & Compliance PowerShell.

    .DESCRIPTION
    Uses the cached identity unless explicit vendor parameters are provided. With no
    cached identity, Exchange handles interactive sign-in unless NoAutoConnect is set.
    Microsoft is removing -Credential from new Exchange Online PowerShell module versions
    released starting December 2026.

    .PARAMETER Credential
    Credentials to pass to Connect-IPPSSession. This authentication method is deprecated.

    .EXAMPLE
    Connect-SCC
    Connects using the cached identity or Exchange's interactive sign-in.

    .EXAMPLE
    Connect-SCC -Credential (Get-Credential)
    Connects using explicit credentials and displays a deprecation warning.
    #>
    [CmdletBinding()]
    param(
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]$Credential
    )

    dynamicparam {
        Get-Office365ConnectParameters -Module 'ExchangeOnlineManagement' -Command 'ExchangeOnlineManagement\Connect-IPPSSession' -Exclude 'Credential'
    }
    process {
        $local:explicitCredential = $PSBoundParameters.ContainsKey('Credential')
        if ($local:explicitCredential) {
            Write-ExchangeCredentialDeprecationWarning
        }
        if (-not (Get-Module -Name ExchangeOnlineManagement -ErrorAction SilentlyContinue)) {
            Import-Module -Name ExchangeOnlineManagement -ErrorAction SilentlyContinue
        }
        $local:command = Get-Command -Name ExchangeOnlineManagement\Connect-IPPSSession -ErrorAction SilentlyContinue
        if ($local:command) {
            if ($local:explicitCredential -and -not $local:command.Parameters.ContainsKey('Credential')) {
                return (New-Office365ConnectionStatus -Service 'Security & Compliance' -StateKey 'ConnectedSCC' -Connected $false -ErrorMessage 'The installed Connect-IPPSSession command does not support the Credential parameter.')
            }
            $local:explicitParams = Get-Office365ExplicitConnectParameters -BoundParameters $PSBoundParameters
            if ($local:explicitParams.Count -gt 0) {
                if ($local:explicitCredential -and -not $local:explicitParams.ContainsKey('PSSessionOption') -and $local:command.Parameters.ContainsKey('PSSessionOption')) {
                    $local:explicitParams['PSSessionOption'] = $script:myOffice365Services['SessionOptions']
                }
                if (-not $local:explicitParams.ContainsKey('ConnectionUri') -and -not $local:explicitParams.ContainsKey('ExchangeEnvironmentName')) {
                    $local:eomEnv = $script:myOffice365Services['EOMEnvironmentName']
                    if ($local:eomEnv -and $local:eomEnv -ne 'O365Default') {
                        if (-not $local:command.Parameters.ContainsKey('ExchangeEnvironmentName')) {
                            return (New-Office365ConnectionStatus -Service 'Security & Compliance' -StateKey 'ConnectedSCC' -Connected $false -ErrorMessage 'The installed Connect-IPPSSession does not support ExchangeEnvironmentName for the selected cloud.')
                        }
                        $local:explicitParams['ExchangeEnvironmentName'] = $local:eomEnv
                    }
                    elseif (-not $local:eomEnv -and $script:myOffice365Services['SCCConnectionEndpointUri']) {
                        $local:explicitParams['ConnectionUri'] = $script:myOffice365Services['SCCConnectionEndpointUri']
                        if (-not $local:explicitParams.ContainsKey('AzureADAuthorizationEndpointUri')) {
                            $local:explicitParams['AzureADAuthorizationEndpointUri'] = $script:myOffice365Services['AzureADAuthorizationEndpointUri']
                        }
                    }
                }
                try {
                    Write-Host 'Connecting to Security & Compliance Center using the specified options ..'
                    $script:myOffice365Services['SessionCC'] = Connect-IPPSSession @local:explicitParams -ErrorAction Stop
                    if ($script:myOffice365Services['SessionCC']) {
                        Import-PSSession -Session $script:myOffice365Services['SessionCC'] -AllowClobber -ErrorAction Stop | Out-Null
                    }
                }
                catch {
                    $script:myOffice365Services['SessionCC'] = $null
                    return (New-Office365ConnectionStatus -Service 'Security & Compliance' -StateKey 'ConnectedSCC' -Connected $false -ErrorMessage ('Cannot connect to Security & Compliance: {0}' -f $_.Exception.Message))
                }
                return (New-Office365ConnectionStatus -Service 'Security & Compliance' -StateKey 'ConnectedSCC' -Connected $true)
            }
            # Let Exchange handle sign-in without loading another module's MSAL assembly.
            if (-not $local:explicitCredential -and -not $script:myOffice365Services['Office365UPN'] -and -not $script:myOffice365Services['Office365Credential']) {
                if ($script:myOffice365Services['NoAutoConnect']) {
                    return (New-Office365ConnectionStatus -Service 'Security & Compliance' -StateKey 'ConnectedSCC' -Connected $false -ErrorMessage 'No credentials cached. Run Get-Office365Credential first or supply credentials explicitly.')
                }
            }

            Write-Host 'Connecting to Security & Compliance Center ..'

            # Same REST-mode logic as Connect-EXO: use ExchangeEnvironmentName for sovereign clouds,
            # omit ConnectionUri for O365Default so EOM v3 stays in REST mode (enables delegated AccessToken).
            $local:connectParams = @{
                PSSessionOption = $script:myOffice365Services['SessionOptions']
            }
            $local:eomEnv = $script:myOffice365Services['EOMEnvironmentName']
            if ($local:eomEnv -and $local:eomEnv -ne 'O365Default') {
                $local:connectParams['ExchangeEnvironmentName'] = $local:eomEnv
            }
            elseif (-not $local:eomEnv -and $script:myOffice365Services['SCCConnectionEndpointUri']) {
                # AzurePPE or unknown: use legacy ConnectionUri
                $local:connectParams['ConnectionUri'] = $script:myOffice365Services['SCCConnectionEndpointUri']
                $local:connectParams['AzureADAuthorizationEndpointUri'] = $script:myOffice365Services['AzureADAuthorizationEndpointUri']
            }
            # O365Default: omit both — EOM v3 auto-discovers the SCC REST endpoint

            # Pass UPN so EOM's own internal MSAL acquires the SCC token silently via WAM SSO.
            # External MSAL clients cannot request EXO/SCC tokens — AAD blocks first-party to
            # first-party token requests from external apps (AADSTS65002).
            if ($local:explicitCredential) {
                $local:connectParams['Credential'] = $Credential
            }
            elseif ($script:myOffice365Services['Office365UPN']) {
                $local:connectParams['UserPrincipalName'] = $script:myOffice365Services['Office365UPN']
            }
            elseif ($script:myOffice365Services['Office365Credential']) {
                $local:connectParams['UserPrincipalName'] = $script:myOffice365Services['Office365Credential'].UserName
            }

            try {
                try {
                    $script:myOffice365Services['SessionCC'] = Connect-IPPSSession @local:connectParams -ErrorAction Stop
                }
                catch {
                    if ($local:explicitCredential) { throw }
                    # WAM (Windows Web Account Manager) broker requires a native window handle that
                    # PowerShell console and terminal hosts never supply, causing a timeout.
                    # Fall back to device code flow which only needs the console.
                    if ($_.Exception.Message -like '*Operation did not start in the allotted time*' -or
                        $_.Exception.Message -like '*Error Acquiring Token*') {
                        Write-Warning 'WAM broker timed out — retrying with device code flow (check console for URL + code) ..'
                        $local:connectParams.Remove('UserPrincipalName') | Out-Null
                        $local:connectParams['Device'] = $true
                        $script:myOffice365Services['SessionCC'] = Connect-IPPSSession @local:connectParams -ErrorAction Stop
                    }
                    else { throw }
                }
                if ($script:myOffice365Services['SessionCC']) {
                    Import-PSSession -Session $script:myOffice365Services['SessionCC'] -AllowClobber -ErrorAction Stop | Out-Null
                }
            }
            catch {
                $script:myOffice365Services['SessionCC'] = $null
                return (New-Office365ConnectionStatus -Service 'Security & Compliance' -StateKey 'ConnectedSCC' -Connected $false -ErrorMessage ('Cannot connect to Security & Compliance: {0}' -f $_.Exception.Message))
            }
            New-Office365ConnectionStatus -Service 'Security & Compliance' -StateKey 'ConnectedSCC' -Connected $true -Detail $script:myOffice365Services['Office365UPN']
        }
        else {
            New-Office365ConnectionStatus -Service 'Security & Compliance' -StateKey 'ConnectedSCC' -Connected $false -ErrorMessage 'Cannot connect to Security & Compliance Center - module not installed or not loading.'
        }
    }
}
