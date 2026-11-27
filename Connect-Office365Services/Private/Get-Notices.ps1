$script:retiredModuleReplacements = [ordered]@{
    'Microsoft.Graph.Entra' = 'Microsoft.Entra'
    'Microsoft.Graph.Entra.Beta' = 'Microsoft.Entra.Beta'
    'Microsoft.Graph.Intune' = 'Microsoft.Graph.DeviceManagement'
}
$script:retiredModules = @('MSOnline', 'AzureAD', 'AzureADPreview') + @($script:retiredModuleReplacements.Keys)

function Get-NoticeModuleNames {
    $script:retiredModules
    if ($PSVersionTable.PSVersion.Major -eq 5) {
        'Microsoft.Graph'
        'Microsoft.Graph.*'
    }
}

function Write-ModuleRetirementNotices {
    param(
        $AllInstalled
    )

    foreach ($local:retiredModule in $script:retiredModules) {
        if ($AllInstalled | Where-Object { $_.Name -eq $local:retiredModule }) {
            if ($script:retiredModuleReplacements.Contains($local:retiredModule)) {
                $local:notice = 'The {0} PowerShell module is retired and is no longer managed. It has been replaced by {1}.' -f $local:retiredModule, $script:retiredModuleReplacements[$local:retiredModule]
                if ($local:retiredModule -eq 'Microsoft.Graph.Intune') {
                    $local:notice = 'Microsoft.Graph.Intune has not been updated since 2019 and is no longer managed. Microsoft.Graph.DeviceManagement is recommended.'
                }
                Write-Warning $local:notice
            }
            else {
                Write-Warning ('The {0} PowerShell module is retired and is no longer managed.' -f $local:retiredModule)
            }
        }
    }
    if ($PSVersionTable.PSVersion.Major -eq 5 -and
        ($AllInstalled | Where-Object {
            ($_.Name -eq 'Microsoft.Graph' -or $_.Name -like 'Microsoft.Graph.*') -and
            $_.Name -notin @('Microsoft.Graph.Intune', 'Microsoft.Graph.Entra', 'Microsoft.Graph.Entra.Beta', 'Microsoft.Graph.Compatibility.AzureAD') -and
            $_.Version.Major -eq 2
        })) {
        Write-Warning 'Microsoft Graph PowerShell v2 on Windows PowerShell 5.x is in its retirement period. Microsoft Graph PowerShell v3 will be released in September 2027 and will support PowerShell 7.x only.'
    }
}

function Write-ExchangeCredentialDeprecationWarning {
    Write-Warning 'The -Credential parameter is being removed from new Exchange Online Management modules starting December 2026. Consider interactive sign-in, certificate-based app-only authentication, or managed identity. See https://techcommunity.microsoft.com/blog/exchange/deprecation-of-the--credential-parameter-in-exchange-online-powershell/4494584'
}
