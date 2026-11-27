function Connect-MG {
    <#
    .SYNOPSIS
    Connects to Microsoft Graph using the cached identity or an interactive sign-in.

    .DESCRIPTION
    Loads the Microsoft Graph Authentication module (Microsoft.Graph.Authentication,
    Microsoft.Graph, or Microsoft.Entra — whichever is installed), acquires an OAuth2
    token via MSAL.NET, and calls Connect-MgGraph with that token.

    All authentication parameters supported by the installed Connect-MgGraph command
    are available. When any are supplied, they are passed through to the vendor command.
    With no options, the cached-identity token flow is retained and uses the selected
    Microsoft cloud.

    .EXAMPLE
    Connect-MG
    Connects to Microsoft Graph using the currently cached UPN or via an interactive sign-in.

    .EXAMPLE
    Connect-MG -Scopes 'User.Read.All' -TenantId '00000000-0000-0000-0000-000000000000'
    Connects with the specified delegated permission and tenant.

    .EXAMPLE
    Connect-MG -Identity -Environment USGov
    Connects using managed identity in Azure US Government.

    .EXAMPLE
    Connect-MG -ClientId '00000000-0000-0000-0000-000000000000' -TenantId '00000000-0000-0000-0000-000000000000' -CertificateThumbprint 'CERTIFICATE_THUMBPRINT'
    Connects app-only with a certificate.
    #>
    [CmdletBinding()]
    param()

    dynamicparam {
        foreach ($local:modName in @('Microsoft.Graph.Authentication', 'Microsoft.Graph', 'Microsoft.Entra')) {
            if (-not (Get-Module -Name $local:modName -ErrorAction SilentlyContinue)) {
                Import-Module -Name $local:modName -ErrorAction SilentlyContinue
            }
            if (Get-Command -Name Connect-MgGraph -ErrorAction SilentlyContinue) { break }
        }

        $local:cmd = Get-Command -Name Connect-MgGraph -ErrorAction SilentlyContinue
        if ($local:cmd) {
            $local:dict = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()
            foreach ($local:p in $local:cmd.Parameters.Values) {
                if ([System.Management.Automation.Cmdlet]::CommonParameters -contains $local:p.Name -or
                    [System.Management.Automation.Cmdlet]::OptionalCommonParameters -contains $local:p.Name) {
                    continue
                }
                $local:dict.Add($local:p.Name,
                    [System.Management.Automation.RuntimeDefinedParameter]::new(
                        $local:p.Name, $local:p.ParameterType, $local:p.Attributes))
            }
            return $local:dict
        }
    }

    process {
        # Module guard — prefer Authentication sub-module for minimal footprint
        foreach ($local:modName in @('Microsoft.Graph.Authentication', 'Microsoft.Graph', 'Microsoft.Entra')) {
            if (-not (Get-Module -Name $local:modName -ErrorAction SilentlyContinue)) {
                Import-Module -Name $local:modName -ErrorAction SilentlyContinue
            }
            if (Get-Command -Name Connect-MgGraph -ErrorAction SilentlyContinue) { break }
        }

        $local:command = Get-Command -Name Connect-MgGraph -ErrorAction SilentlyContinue
        if (-not $local:command) {
            return (New-Office365ConnectionStatus -Service 'Microsoft Graph' -StateKey 'ConnectedGraph' -Connected $false -ErrorMessage 'Cannot connect to Microsoft Graph - module not installed or not loading. Install Microsoft.Graph.Authentication or Microsoft.Graph.')
        }

        $local:graphParams = @{}
        foreach ($local:paramName in $PSBoundParameters.Keys) {
            if ($local:command.Parameters.ContainsKey($local:paramName) -and
                [System.Management.Automation.Cmdlet]::CommonParameters -notcontains $local:paramName -and
                [System.Management.Automation.Cmdlet]::OptionalCommonParameters -notcontains $local:paramName) {
                $local:graphParams[$local:paramName] = $PSBoundParameters[$local:paramName]
            }
        }

        if ($local:graphParams.Count -gt 0) {
            try {
                if (-not $local:graphParams.ContainsKey('Environment') -and
                    -not $local:graphParams.ContainsKey('EnvironmentVariable')) {
                    $local:graphEnvironment = Get-Office365GraphEnvironment
                    if ($local:command.Parameters.ContainsKey('Environment')) {
                        $local:graphParams['Environment'] = $local:graphEnvironment.Name
                    }
                    elseif ($local:graphEnvironment.Name -ne 'Global') {
                        throw 'The installed Graph SDK does not support cloud selection with -Environment; update Microsoft.Graph.Authentication.'
                    }
                }
                if (-not $local:graphParams.ContainsKey('NoWelcome') -and $local:command.Parameters.ContainsKey('NoWelcome')) {
                    $local:graphParams['NoWelcome'] = $true
                }
                $local:graphParams['ErrorAction'] = 'Stop'
                Write-Host 'Connecting to Microsoft Graph using the specified authentication options ..'
                $null = Connect-MgGraph @local:graphParams
                $local:mgContext = if (Get-Command -Name Get-MgContext -ErrorAction SilentlyContinue) {
                    Get-MgContext -ErrorAction Stop
                }
                if ((Get-Command -Name Get-MgContext -ErrorAction SilentlyContinue) -and -not $local:mgContext) {
                    return (New-Office365ConnectionStatus -Service 'Microsoft Graph' -StateKey 'ConnectedGraph' -Connected $false -ErrorMessage 'Connect-MgGraph completed without an active Graph context.')
                }
            }
            catch {
                return (New-Office365ConnectionStatus -Service 'Microsoft Graph' -StateKey 'ConnectedGraph' -Connected $false -ErrorMessage ('Cannot connect to Microsoft Graph: {0}' -f $_.Exception.Message))
            }
            $local:detail = if ($local:mgContext.Account) { $local:mgContext.Account } elseif ($local:mgContext.AppName) { $local:mgContext.AppName } else { $null }
            return (New-Office365ConnectionStatus -Service 'Microsoft Graph' -StateKey 'ConnectedGraph' -Connected $true -Detail $local:detail)
        }

        try {
            $local:graphEnvironment = Get-Office365GraphEnvironment
        }
        catch {
            return (New-Office365ConnectionStatus -Service 'Microsoft Graph' -StateKey 'ConnectedGraph' -Connected $false -ErrorMessage $_.Exception.Message)
        }

        # Ensure we have an account cached (MSAL) or credentials (legacy)
        if (-not $script:myOffice365Services['Office365UPN'] -and -not $script:myOffice365Services['Office365Credential']) {
            if ($script:myOffice365Services['NoAutoConnect']) {
                return (New-Office365ConnectionStatus -Service 'Microsoft Graph' -StateKey 'ConnectedGraph' -Connected $false -ErrorMessage 'No credentials cached. Run Get-Office365Credential first or supply credentials explicitly.')
            }
            try {
                $null = Get-Office365Credential
            }
            catch {
                return (New-Office365ConnectionStatus -Service 'Microsoft Graph' -StateKey 'ConnectedGraph' -Connected $false -ErrorMessage ('Could not acquire credentials for Microsoft Graph: {0}' -f $_.Exception.Message))
            }
        }

        $local:upn = if ($script:myOffice365Services['Office365UPN']) {
            $script:myOffice365Services['Office365UPN']
        }
        else {
            $script:myOffice365Services['Office365Credential'].UserName
        }

        try {
            # Acquire MSAL token for Graph
            $local:graphScope = '{0}/.default' -f $local:graphEnvironment.GraphEndpoint
            if (-not $local:command.Parameters.ContainsKey('Environment') -and $local:graphEnvironment.Name -ne 'Global') {
                return (New-Office365ConnectionStatus -Service 'Microsoft Graph' -StateKey 'ConnectedGraph' -Connected $false -ErrorMessage 'The installed Graph SDK does not support cloud selection with -Environment; update Microsoft.Graph.Authentication.')
            }
            $local:graphToken = Get-Office365AccessToken -Scope $local:graphScope
            if ($local:graphToken) {
                Write-Host ('Connecting to Microsoft Graph using {0} ..' -f $local:upn)
                # Graph SDK v2 requires AccessToken as SecureString
                $local:secureToken = ConvertTo-SecureString -String $local:graphToken -AsPlainText -Force
                $local:graphParams = @{
                    AccessToken = $local:secureToken
                    ErrorAction = 'Stop'
                }
                if ($local:command.Parameters.ContainsKey('Environment')) {
                    $local:graphParams['Environment'] = $local:graphEnvironment.Name
                }
                if ($local:command.Parameters.ContainsKey('NoWelcome')) {
                    $local:graphParams['NoWelcome'] = $true
                }
                $null = Connect-MgGraph @local:graphParams
            }
            else {
                # Fallback: let the Graph SDK run its own interactive flow
                Write-Host ('Connecting to Microsoft Graph using {0} ..' -f $local:upn)
                $local:graphParams = @{
                    Scopes      = @($local:graphScope)
                    ErrorAction = 'Stop'
                }
                if ($local:command.Parameters.ContainsKey('Environment')) {
                    $local:graphParams['Environment'] = $local:graphEnvironment.Name
                }
                if ($local:command.Parameters.ContainsKey('NoWelcome')) {
                    $local:graphParams['NoWelcome'] = $true
                }
                $null = Connect-MgGraph @local:graphParams
            }
            if (Get-Command -Name Get-MgContext -ErrorAction SilentlyContinue) {
                $local:mgContext = Get-MgContext -ErrorAction Stop
                if (-not $local:mgContext) {
                    return (New-Office365ConnectionStatus -Service 'Microsoft Graph' -StateKey 'ConnectedGraph' -Connected $false -ErrorMessage 'Connect-MgGraph completed without an active Graph context.')
                }
            }
        }
        catch {
            return (New-Office365ConnectionStatus -Service 'Microsoft Graph' -StateKey 'ConnectedGraph' -Connected $false -ErrorMessage ('Cannot connect to Microsoft Graph: {0}' -f $_.Exception.Message))
        }
        New-Office365ConnectionStatus -Service 'Microsoft Graph' -StateKey 'ConnectedGraph' -Connected $true -Detail $local:upn
    }
}
