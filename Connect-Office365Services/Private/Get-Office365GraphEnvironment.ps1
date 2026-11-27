function Get-Office365GraphEnvironment {
    $local:environmentName = [string]$script:myOffice365Services['AzureEnvironmentName']
    switch ($local:environmentName) {
        { $_ -in @('Default', 'GCC', 'Germany') } {
            [PSCustomObject]@{
                Name         = 'Global'
                GraphEndpoint = 'https://graph.microsoft.com'
            }
        }
        'China' {
            [PSCustomObject]@{
                Name         = 'China'
                GraphEndpoint = 'https://microsoftgraph.chinacloudapi.cn'
            }
        }
        'GCCHigh' {
            [PSCustomObject]@{
                Name         = 'USGov'
                GraphEndpoint = 'https://graph.microsoft.us'
            }
        }
        'DoD' {
            [PSCustomObject]@{
                Name         = 'USGovDoD'
                GraphEndpoint = 'https://dod-graph.microsoft.us'
            }
        }
        'AzurePPE' {
            throw 'AzurePPE has no built-in Microsoft Graph environment. Supply -Environment for a registered Graph SDK environment.'
        }
        default {
            throw ('No Microsoft Graph environment mapping is configured for {0}.' -f $local:environmentName)
        }
    }
}
