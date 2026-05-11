function Connect-ExchangeOnline {
    [CmdletBinding()]
    param()

    DynamicParam {
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
        If ( $PSBoundParameters.ContainsKey('UserPrincipalName') -or $PSBoundParameters.ContainsKey('Certificate') -or
            $PSBoundParameters.ContainsKey('CertificateFilePath') -or $PSBoundParameters.ContainsKey('CertificateThumbprint') -or
            $PSBoundParameters.ContainsKey('AppId')) {
            Write-Host ('Connecting to Exchange Online ..')
        }
        Else {
            If ( $PSBoundParameters.ContainsKey('Credential')) {
                Write-Host ('Connecting to Exchange Online using {0} ..' -f $PSBoundParameters['Credential'].UserName)
                $script:myOffice365Services['Office365Credentials'] = $PSBoundParameters['Credential']
            }
            Else {
                If ( $script:myOffice365Services['Office365Credentials']) {
                    Write-Host ('Connecting to Exchange Online using {0} ..' -f $script:myOffice365Services['Office365Credentials'].UserName)
                    $PSBoundParameters['Credential'] = $script:myOffice365Services['Office365Credentials']
                }
                Else {
                    Get-Office365Credentials
                    If ( $script:myOffice365Services['Office365Credentials']) {
                        Write-Host ('Connecting to Exchange Online using {0} ..' -f $script:myOffice365Services['Office365Credentials'].UserName)
                        $PSBoundParameters['Credential'] = $script:myOffice365Services['Office365Credentials']
                    }
                    Else {
                        Write-Host ('Connecting to Exchange Online ..')
                    }
                }
            }
        }

        $script:myOffice365Services['Session365'] = ExchangeOnlineManagement\Connect-ExchangeOnline @PSBoundParameters
        If ( $script:myOffice365Services['Session365']) {
            Import-PSSession -Session $script:myOffice365Services['Session365'] -AllowClobber
        }
    }
}
