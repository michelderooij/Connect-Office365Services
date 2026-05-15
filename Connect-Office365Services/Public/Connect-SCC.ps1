function Connect-SCC {
    if (-not (Get-Module -Name ExchangeOnlineManagement -ErrorAction SilentlyContinue)) {
        Import-Module -Name ExchangeOnlineManagement -ErrorAction SilentlyContinue
    }
    if (Get-Command -Name ExchangeOnlineManagement\Connect-IPPSSession -ErrorAction SilentlyContinue) {
        # Ensure we have an account cached (MSAL) or credentials (legacy)
        if ( -not $script:myOffice365Services['Office365UPN'] -and -not $script:myOffice365Services['Office365Credential']) {
            if ($script:myOffice365Services['NoAutoConnect']) {
                Write-Error 'No credentials cached. Run Get-Office365Credential first or supply credentials explicitly.'
                return
            }
            Get-Office365Credential
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
        if ($script:myOffice365Services['Office365UPN']) {
            $local:connectParams['UserPrincipalName'] = $script:myOffice365Services['Office365UPN']
        }
        elseif ($script:myOffice365Services['Office365Credential']) {
            $local:connectParams['UserPrincipalName'] = $script:myOffice365Services['Office365Credential'].UserName
        }

        try {
            $script:myOffice365Services['SessionCC'] = Connect-IPPSSession @local:connectParams
        }
        catch {
            # WAM (Windows Web Account Manager) broker requires a native window handle that
            # PowerShell console and terminal hosts never supply, causing a timeout.
            # Fall back to device code flow which only needs the console.
            if ($_.Exception.Message -like '*Operation did not start in the allotted time*' -or
                $_.Exception.Message -like '*Error Acquiring Token*') {
                Write-Warning 'WAM broker timed out — retrying with device code flow (check console for URL + code) ..'
                $local:connectParams.Remove('UserPrincipalName') | Out-Null
                $local:connectParams['Device'] = $true
                $script:myOffice365Services['SessionCC'] = Connect-IPPSSession @local:connectParams
            }
            else { throw }
        }
        if ( $script:myOffice365Services['SessionCC'] ) {
            Import-PSSession -Session $script:myOffice365Services['SessionCC'] -AllowClobber
        }
        $script:myOffice365Services['ConnectedSCC'] = $true
    }
    else {
        Write-Error -Message 'Cannot connect to Security & Compliance Center - module not installed or not loading.'
    }
}
