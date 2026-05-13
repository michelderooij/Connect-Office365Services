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

        # Inject module defaults when caller did not supply them
        if (-not $PSBoundParameters.ContainsKey('ConnectionUri')) {
            $PSBoundParameters['ConnectionUri'] = $script:myOffice365Services['ConnectionEndpointUri']
        }
        if (-not $PSBoundParameters.ContainsKey('AzureADAuthorizationEndpointUri')) {
            $PSBoundParameters['AzureADAuthorizationEndpointUri'] = $script:myOffice365Services['AzureADAuthorizationEndpointUri']
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
                    Get-Office365Credential
                }

                # Modern auth: acquire EXO-scoped token and pass via -AccessToken + -UserPrincipalName
                $local:exoToken = Get-Office365AccessToken -Scope 'https://outlook.office365.com/.default'
                if ($local:exoToken) {
                    $PSBoundParameters['AccessToken'] = ConvertTo-SecureString $local:exoToken -AsPlainText -Force
                    $PSBoundParameters['UserPrincipalName'] = $script:myOffice365Services['Office365UPN']
                    Write-Host ('Connecting to Exchange Online using {0} ..' -f $script:myOffice365Services['Office365UPN'])
                }
                elseif ( $script:myOffice365Services['Office365UPN']) {
                    # MSAL auth done but EXO scope unavailable — use UPN for WAM SSO
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

        $script:myOffice365Services['Session365'] = Connect-ExchangeOnline @PSBoundParameters
        if ( $script:myOffice365Services['Session365']) {
            Import-PSSession -Session $script:myOffice365Services['Session365'] -AllowClobber
        }
    }
}
