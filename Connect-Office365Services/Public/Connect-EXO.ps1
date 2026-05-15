function Connect-EXO {
    [CmdletBinding()]
    param()

    dynamicparam {
        if (-not (Get-Module -Name ExchangeOnlineManagement -ErrorAction SilentlyContinue)) {
            Import-Module -Name ExchangeOnlineManagement -ErrorAction SilentlyContinue
        }
        $local:cmd = Get-Command -Name ExchangeOnlineManagement\Connect-ExchangeOnline -ErrorAction SilentlyContinue
        if ($local:cmd) {
            $local:dict = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()
            foreach ($local:p in $local:cmd.Parameters.Values) {
                if ([System.Management.Automation.Cmdlet]::CommonParameters -contains $local:p.Name) { continue }
                $local:dict.Add($local:p.Name,
                    [System.Management.Automation.RuntimeDefinedParameter]::new(
                        $local:p.Name, $local:p.ParameterType, $local:p.Attributes))
            }
            return $local:dict
        }
    }

    process {
        if (-not (Get-Module -Name ExchangeOnlineManagement -ErrorAction SilentlyContinue)) {
            Import-Module -Name ExchangeOnlineManagement -ErrorAction SilentlyContinue
        }
        if (-not (Get-Command -Name ExchangeOnlineManagement\Connect-ExchangeOnline -ErrorAction SilentlyContinue)) {
            Write-Error -Message 'Cannot connect to Exchange Online - module not installed or not loading.'
            return
        }

        # EOM v3 uses REST mode by default; do NOT inject the legacy /PowerShell-LiveId ConnectionUri —
        # that forces RPS mode where delegated AccessToken auth returns 401.
        # For sovereign clouds use -ExchangeEnvironmentName so EOM v3 picks the right REST endpoint.
        # AzurePPE has no standard EOM environment name, so fall back to ConnectionUri there.
        if (-not $PSBoundParameters.ContainsKey('ConnectionUri') -and
            -not $PSBoundParameters.ContainsKey('ExchangeEnvironmentName')) {
            $local:eomEnv = $script:myOffice365Services['EOMEnvironmentName']
            if ($local:eomEnv -and $local:eomEnv -ne 'O365Default') {
                $PSBoundParameters['ExchangeEnvironmentName'] = $local:eomEnv
            }
            elseif (-not $local:eomEnv -and $script:myOffice365Services['ConnectionEndpointUri']) {
                # AzurePPE or unknown: use legacy ConnectionUri
                $PSBoundParameters['ConnectionUri'] = $script:myOffice365Services['ConnectionEndpointUri']
                $PSBoundParameters['AzureADAuthorizationEndpointUri'] = $script:myOffice365Services['AzureADAuthorizationEndpointUri']
            }
            # O365Default (worldwide + GCC): omit both — EOM v3 auto-discovers the REST endpoint
        }
        if (-not $PSBoundParameters.ContainsKey('PSSessionOption')) {
            $PSBoundParameters['PSSessionOption'] = $script:myOffice365Services['SessionOptions']
        }

        # Credential handling — skip when modern/cert/app auth params were supplied
        if ( $PSBoundParameters.ContainsKey('UserPrincipalName') -or $PSBoundParameters.ContainsKey('Certificate') -or
            $PSBoundParameters.ContainsKey('CertificateFilePath') -or $PSBoundParameters.ContainsKey('CertificateThumbprint') -or
            $PSBoundParameters.ContainsKey('AppId')) {
            Write-Host ('Connecting to Exchange Online ..')
        }
        else {
            if ( $PSBoundParameters.ContainsKey('Credential')) {
                Write-Host ('Connecting to Exchange Online using {0} ..' -f $PSBoundParameters['Credential'].UserName)
                $script:myOffice365Services['Office365Credential'] = $PSBoundParameters['Credential']
            }
            else {
                # Ensure we have an account cached (MSAL) or credentials (legacy)
                if ( -not $script:myOffice365Services['Office365UPN'] -and -not $script:myOffice365Services['Office365Credential']) {
                    if ($script:myOffice365Services['NoAutoConnect']) {
                        Write-Error 'No credentials cached. Run Get-Office365Credential first or supply credentials explicitly.'
                        return
                    }
                    Get-Office365Credential
                }

                if ( $script:myOffice365Services['Office365UPN']) {
                    # Pass UPN so EOM's own internal MSAL acquires the EXO token silently via WAM SSO.
                    # External MSAL clients cannot request EXO tokens — AAD blocks first-party to
                    # first-party token requests from external apps (AADSTS65002).
                    $PSBoundParameters['UserPrincipalName'] = $script:myOffice365Services['Office365UPN']
                    Write-Host ('Connecting to Exchange Online using {0} ..' -f $script:myOffice365Services['Office365UPN'])
                }
                elseif ( $script:myOffice365Services['Office365Credential']) {
                    # Legacy PSCredential path
                    Write-Host ('Connecting to Exchange Online using {0} ..' -f $script:myOffice365Services['Office365Credential'].UserName)
                    $PSBoundParameters['Credential'] = $script:myOffice365Services['Office365Credential']
                }
                else {
                    Write-Host ('Connecting to Exchange Online ..')
                }
            }
        }

        try {
            $script:myOffice365Services['Session365'] = Connect-ExchangeOnline @PSBoundParameters
        }
        catch {
            # WAM (Windows Web Account Manager) broker requires a native window handle that
            # PowerShell console and terminal hosts never supply, causing a timeout.
            # Fall back to device code flow which only needs the console.
            if ($_.Exception.Message -like '*Operation did not start in the allotted time*' -or
                $_.Exception.Message -like '*Error Acquiring Token*') {
                Write-Warning 'WAM broker timed out — retrying with device code flow (check console for URL + code) ..'
                $PSBoundParameters.Remove('UserPrincipalName') | Out-Null
                $PSBoundParameters['Device'] = $true
                $script:myOffice365Services['Session365'] = Connect-ExchangeOnline @PSBoundParameters
            }
            else { throw }
        }
        if ( $script:myOffice365Services['Session365']) {
            Import-PSSession -Session $script:myOffice365Services['Session365'] -AllowClobber
        }
        $script:myOffice365Services['ConnectedEXO'] = $true
    }
}
