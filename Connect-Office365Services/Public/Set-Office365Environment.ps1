function Set-Office365Environment {
    param(
        [ValidateSet('Germany', 'China', 'AzurePPE', 'GCC', 'GCCHigh', 'DoD', 'Default')]
        [string]$Environment
    )
    Switch ( $Environment) {
        'Germany' {
            # Microsoft Cloud Germany (T-Systems trustee) was decommissioned Oct 2021.
            # New German datacenter regions use worldwide AzureCloud infrastructure.
            $script:myOffice365Services['AzureEnvironmentName'] = 'Germany'
            $script:myOffice365Services['ConnectionEndpointUri'] = 'https://outlook.office365.com/PowerShell-LiveId'
            $script:myOffice365Services['SCCConnectionEndpointUri'] = 'https://ps.compliance.protection.outlook.com/PowerShell-LiveId'
            $script:myOffice365Services['EOPConnectionEndpointUri'] = 'https://ps.protection.protection.outlook.com/PowerShell-LiveId'
            $script:myOffice365Services['AzureADAuthorizationEndpointUri'] = 'https://login.microsoftonline.com/common'
            $script:myOffice365Services['EOMEnvironmentName'] = 'O365GermanyCloud'
            $script:myOffice365Services['SharePointRegion'] = 'Germany'
            $script:myOffice365Services['AzureEnvironment'] = 'AzureCloud'
            $script:myOffice365Services['TeamsEnvironment'] = ''
        }
        'China' {
            # China operated by 21Vianet uses separate AzureChinaCloud infrastructure and login endpoints:
            $script:myOffice365Services['AzureEnvironmentName'] = 'China'
            $script:myOffice365Services['ConnectionEndpointUri'] = 'https://partner.outlook.cn/PowerShell-LiveID'
            $script:myOffice365Services['SCCConnectionEndpointUri'] = 'https://ps.compliance.protection.outlook.com/PowerShell-LiveId'
            $script:myOffice365Services['EOPConnectionEndpointUri'] = 'https://ps.protection.protection.outlook.com/PowerShell-LiveId'
            $script:myOffice365Services['AzureADAuthorizationEndpointUri'] = 'https://login.chinacloudapi.cn/common'
            $script:myOffice365Services['EOMEnvironmentName'] = 'O365China'
            $script:myOffice365Services['SharePointRegion'] = 'China'
            $script:myOffice365Services['AzureEnvironment'] = 'AzureChinaCloud'
            $script:myOffice365Services['TeamsEnvironment'] = ''
        }
        'AzurePPE' {
            # Azure Public Preview environment — no standard EOM environment name; uses ConnectionUri fallback.
            $script:myOffice365Services['AzureEnvironmentName'] = 'AzurePPE'
            $script:myOffice365Services['ConnectionEndpointUri'] = ''
            $script:myOffice365Services['SCCConnectionEndpointUri'] = 'https://ps.compliance.protection.outlook.com/PowerShell-LiveId'
            $script:myOffice365Services['EOPConnectionEndpointUri'] = 'https://ps.protection.protection.outlook.com/PowerShell-LiveId'
            $script:myOffice365Services['AzureADAuthorizationEndpointUri'] = ''
            $script:myOffice365Services['EOMEnvironmentName'] = ''
            $script:myOffice365Services['SharePointRegion'] = ''
            $script:myOffice365Services['AzureEnvironment'] = 'AzurePPE'
            $script:myOffice365Services['TeamsEnvironment'] = ''
        }
        'GCC' {
            # Standard Government Community Cloud, uses worldwide commercial infrastructure but with government tenant isolation:
            $script:myOffice365Services['AzureEnvironmentName'] = 'GCC'
            $script:myOffice365Services['ConnectionEndpointUri'] = 'https://outlook.office365.com/PowerShell-LiveId'
            $script:myOffice365Services['SCCConnectionEndpointUri'] = 'https://ps.compliance.protection.outlook.com/PowerShell-LiveId'
            $script:myOffice365Services['EOPConnectionEndpointUri'] = 'https://ps.protection.protection.outlook.com/PowerShell-LiveId'
            $script:myOffice365Services['AzureADAuthorizationEndpointUri'] = 'https://login.microsoftonline.com/common'
            $script:myOffice365Services['EOMEnvironmentName'] = 'O365Default'
            $script:myOffice365Services['SharePointRegion'] = 'ITAR'
            $script:myOffice365Services['AzureEnvironment'] = 'AzureCloud'
            $script:myOffice365Services['TeamsEnvironment'] = 'TeamsGCC'
        }
        'GCCHigh' {
            # GCC High uses the same .us sovereign infrastructure as DoD but on shared GCC High hostnames and with less stringent compliance controls than DoD:
            $script:myOffice365Services['AzureEnvironmentName'] = 'GCCHigh'
            $script:myOffice365Services['ConnectionEndpointUri'] = 'https://outlook.office365.us/PowerShell-LiveId'
            $script:myOffice365Services['SCCConnectionEndpointUri'] = 'https://ps.compliance.protection.office365.us/PowerShell-LiveId'
            $script:myOffice365Services['EOPConnectionEndpointUri'] = 'https://ps.protection.protection.office365.us/PowerShell-LiveId'
            $script:myOffice365Services['AzureADAuthorizationEndpointUri'] = 'https://login.microsoftonline.us/common'
            $script:myOffice365Services['EOMEnvironmentName'] = 'O365USGovGCCHigh'
            $script:myOffice365Services['SharePointRegion'] = 'ITAR'
            $script:myOffice365Services['AzureEnvironment'] = 'AzureUSGovernment'
            $script:myOffice365Services['TeamsEnvironment'] = 'TeamsGCCHigh'
        }
        'DoD' {
            # DoD uses the same .us sovereign infrastructure as GCCHigh but on dedicated DoD-specific hostnames and with additional compliance controls:
            $script:myOffice365Services['AzureEnvironmentName'] = 'DoD'
            $script:myOffice365Services['ConnectionEndpointUri'] = 'https://outlook-dod.office365.us/PowerShell-LiveId'
            $script:myOffice365Services['SCCConnectionEndpointUri'] = 'https://l5.ps.compliance.protection.office365.us/PowerShell-LiveId'
            $script:myOffice365Services['EOPConnectionEndpointUri'] = 'https://ps.protection.protection.office365.us/PowerShell-LiveId'
            $script:myOffice365Services['AzureADAuthorizationEndpointUri'] = 'https://login.microsoftonline.us/common'
            $script:myOffice365Services['EOMEnvironmentName'] = 'O365USGovDoD'
            $script:myOffice365Services['SharePointRegion'] = 'USGovernmentDoD'
            $script:myOffice365Services['AzureEnvironment'] = 'AzureUSGovernment'
            $script:myOffice365Services['TeamsEnvironment'] = 'TeamsGCCHigh'
        }
        default {
            # WWW/global commercial:
            $script:myOffice365Services['AzureEnvironmentName'] = 'Default'
            $script:myOffice365Services['ConnectionEndpointUri'] = 'https://outlook.office365.com/PowerShell-LiveId'
            $script:myOffice365Services['SCCConnectionEndpointUri'] = 'https://ps.compliance.protection.outlook.com/PowerShell-LiveId'
            $script:myOffice365Services['EOPConnectionEndpointUri'] = 'https://ps.protection.protection.outlook.com/PowerShell-LiveId'
            $script:myOffice365Services['AzureADAuthorizationEndpointUri'] = 'https://login.microsoftonline.com/common'
            $script:myOffice365Services['EOMEnvironmentName'] = 'O365Default'
            $script:myOffice365Services['SharePointRegion'] = 'Default'
            $script:myOffice365Services['AzureEnvironment'] = 'AzureCloud'
            $script:myOffice365Services['TeamsEnvironment'] = ''
        }
    }
}
