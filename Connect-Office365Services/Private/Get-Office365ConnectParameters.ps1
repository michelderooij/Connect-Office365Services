function Get-Office365ConnectParameters {
    param(
        [string[]]$Module,
        [string]$Command,
        [string[]]$Exclude = @(),
        [string[]]$Include = @()
    )

    $local:commandInfo = Get-Command -Name $Command -ErrorAction SilentlyContinue | Select-Object -First 1
    foreach ($local:moduleName in $Module) {
        if ($local:commandInfo) { break }
        Import-Module -Name $local:moduleName -ErrorAction SilentlyContinue
        $local:commandInfo = Get-Command -Name $Command -ErrorAction SilentlyContinue | Select-Object -First 1
    }
    if (-not $local:commandInfo) { return }
    if ($local:commandInfo.CommandType -eq 'Alias') {
        $local:commandInfo = $local:commandInfo.ResolvedCommand
    }

    $local:parameters = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()
    $local:usedAliases = @()
    foreach ($local:parameter in $local:commandInfo.Parameters.Values) {
        if (($Include.Count -gt 0 -and $local:parameter.Name -notin $Include) -or
            $local:parameter.Name -in $Exclude -or
            $local:parameter.Name -in [System.Management.Automation.Cmdlet]::CommonParameters -or
            $local:parameter.Name -in [System.Management.Automation.Cmdlet]::OptionalCommonParameters -or
            $local:parameter.Name -in @('WhatIf', 'Confirm')) {
            continue
        }
        $local:attributes = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
        $local:attributes.Add([System.Management.Automation.ParameterAttribute]::new())
        $local:aliases = @($local:parameter.Aliases | Where-Object {
            $_ -notin $Exclude -and $_ -notin $local:usedAliases -and
            -not $local:commandInfo.Parameters.ContainsKey($_) -and
            $_ -notin [System.Management.Automation.Cmdlet]::CommonParameters -and
            $_ -notin [System.Management.Automation.Cmdlet]::OptionalCommonParameters -and
            $_ -notin @('WhatIf', 'Confirm')
        })
        if ($local:aliases.Count -gt 0) {
            $local:attributes.Add([System.Management.Automation.AliasAttribute]::new([string[]]$local:aliases))
            $local:usedAliases += $local:aliases
        }
        $local:parameters.Add($local:parameter.Name,
            [System.Management.Automation.RuntimeDefinedParameter]::new(
                $local:parameter.Name, $local:parameter.ParameterType, $local:attributes))
    }
    return $local:parameters
}

function Get-Office365ExplicitConnectParameters {
    param(
        [System.Collections.IDictionary]$BoundParameters,
        [string[]]$Exclude = @()
    )

    $local:parameters = @{}
    foreach ($local:name in $BoundParameters.Keys) {
        if ($local:name -notin $Exclude -and
            $local:name -notin [System.Management.Automation.Cmdlet]::CommonParameters -and
            $local:name -notin [System.Management.Automation.Cmdlet]::OptionalCommonParameters -and
            $local:name -notin @('WhatIf', 'Confirm')) {
            $local:parameters[$local:name] = $BoundParameters[$local:name]
        }
    }
    return $local:parameters
}
