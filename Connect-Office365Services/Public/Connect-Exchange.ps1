function Connect-Exchange {
    <#
    .SYNOPSIS
    Connects to Exchange On-Premises.
    .DESCRIPTION
    Credential, Authentication, CertificateThumbprint and ConnectionUri are passed
    to New-PSSession when specified. Certificate authentication requires an explicit
    HTTPS ConnectionUri. Without options, the cached Kerberos connection is used.
    #>
    [CmdletBinding()]
    param()

    dynamicparam {
        Get-Office365ConnectParameters -Command 'New-PSSession' -Include @('Credential', 'Authentication', 'CertificateThumbprint', 'ConnectionUri')
    }
    process {
        $local:connectParams = Get-Office365ExplicitConnectParameters -BoundParameters $PSBoundParameters
        if ($local:connectParams.Count -gt 0) {
            try {
                if ($local:connectParams.ContainsKey('CertificateThumbprint') -and
                    (-not $local:connectParams.ContainsKey('ConnectionUri') -or $local:connectParams['ConnectionUri'].Scheme -ne 'https')) {
                    throw 'Certificate authentication requires an explicit HTTPS ConnectionUri.'
                }
                if (-not $local:connectParams.ContainsKey('ConnectionUri')) {
                    if (-not $script:myOffice365Services['ExchangeOnPremisesFQDN']) {
                        $null = Get-ExchangeOnPremisesFQDN
                    }
                    if (-not $script:myOffice365Services['ExchangeOnPremisesFQDN']) {
                        throw 'No Exchange On-Premises FQDN is configured.'
                    }
                    $local:connectParams['ConnectionUri'] = 'http://{0}/PowerShell' -f $script:myOffice365Services['ExchangeOnPremisesFQDN']
                }
                $local:connectParams['ConfigurationName'] = 'Microsoft.Exchange'
                $local:connectParams['AllowRedirection'] = $true
                $local:connectParams['SessionOption'] = $script:myOffice365Services['SessionOptions']
                Write-Host 'Connecting to Exchange On-Premises using the specified options ..'
                $script:myOffice365Services['SessionExchange'] = New-PSSession @local:connectParams -ErrorAction Stop
                if (-not $script:myOffice365Services['SessionExchange'] -or $script:myOffice365Services['SessionExchange'].State -ne 'Opened') {
                    throw 'The Exchange On-Premises session was not opened.'
                }
                Import-PSSession -Session $script:myOffice365Services['SessionExchange'] -AllowClobber -ErrorAction Stop | Out-Null
            }
            catch {
                $script:myOffice365Services['SessionExchange'] = $null
                return (New-Office365ConnectionStatus -Service 'Exchange On-Premises' -StateKey 'ConnectedExchange' -Connected $false -ErrorMessage ('Cannot connect to Exchange On-Premises: {0}' -f $_.Exception.Message))
            }
            return (New-Office365ConnectionStatus -Service 'Exchange On-Premises' -StateKey 'ConnectedExchange' -Connected $true -Detail ([string]$local:connectParams['ConnectionUri']))
        }

        try {
            if (-not $script:myOffice365Services['OnPremisesCredential']) {
                if ($script:myOffice365Services['NoAutoConnect']) {
                    return (New-Office365ConnectionStatus -Service 'Exchange On-Premises' -StateKey 'ConnectedExchange' -Connected $false -ErrorMessage 'No credentials cached. Run Get-OnPremisesCredential first.')
                }
                $null = Get-OnPremisesCredential
            }
            if (-not $script:myOffice365Services['ExchangeOnPremisesFQDN']) {
                $null = Get-ExchangeOnPremisesFQDN
            }

            if (-not $script:myOffice365Services['OnPremisesCredential']) {
                return (New-Office365ConnectionStatus -Service 'Exchange On-Premises' -StateKey 'ConnectedExchange' -Connected $false -ErrorMessage 'No on-premises credentials were provided.')
            }
            if (-not $script:myOffice365Services['ExchangeOnPremisesFQDN']) {
                return (New-Office365ConnectionStatus -Service 'Exchange On-Premises' -StateKey 'ConnectedExchange' -Connected $false -ErrorMessage 'No Exchange On-Premises FQDN is configured.')
            }

            Write-Host ('Connecting to Exchange On-Premises {0} using {1} ..' -f $script:myOffice365Services['ExchangeOnPremisesFQDN'], $script:myOffice365Services['OnPremisesCredential'].UserName)
            $script:myOffice365Services['SessionExchange'] = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$($script:myOffice365Services['ExchangeOnPremisesFQDN'])/PowerShell" -Credential $script:myOffice365Services['OnPremisesCredential'] -Authentication Kerberos -AllowRedirection -SessionOption $script:myOffice365Services['SessionOptions'] -ErrorAction Stop
            if (-not $script:myOffice365Services['SessionExchange'] -or $script:myOffice365Services['SessionExchange'].State -ne 'Opened') {
                return (New-Office365ConnectionStatus -Service 'Exchange On-Premises' -StateKey 'ConnectedExchange' -Connected $false -ErrorMessage 'The Exchange On-Premises session was not opened.')
            }
            Import-PSSession -Session $script:myOffice365Services['SessionExchange'] -AllowClobber -ErrorAction Stop | Out-Null
        }
        catch {
            $script:myOffice365Services['SessionExchange'] = $null
            return (New-Office365ConnectionStatus -Service 'Exchange On-Premises' -StateKey 'ConnectedExchange' -Connected $false -ErrorMessage ('Cannot connect to Exchange On-Premises: {0}' -f $_.Exception.Message))
        }

        New-Office365ConnectionStatus -Service 'Exchange On-Premises' -StateKey 'ConnectedExchange' -Connected $true -Detail $script:myOffice365Services['ExchangeOnPremisesFQDN']
    }
}
