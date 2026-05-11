function Get-Office365Services {
    <#
    .SYNOPSIS
    Returns the current module configuration and session state.
    .DESCRIPTION
    Provides read access to the module-scoped state hashtable that holds environment
    endpoints, credentials, session handles, and install preferences.
    Returns a copy so callers cannot accidentally mutate module state directly.
    #>
    [CmdletBinding()]
    param()
    # Return a shallow copy to prevent external mutation of module state
    $copy = @{}
    foreach ($key in $script:myOffice365Services.Keys) {
        $copy[$key] = $script:myOffice365Services[$key]
    }
    [PSCustomObject]$copy
}
