function Get-OnlineModuleVersion {
    <#
    .SYNOPSIS
    Returns the latest online version string for a module, using a 60-minute
    in-session cache to avoid redundant PSGallery round-trips.
    The cache ($script:myOffice365Services['OnlineVersionCache']) is pre-populated
    in parallel by Show-Office365Modules and Update-Office365Modules (PS 7+).
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$Name,
        [int]$MaxAgeMinutes = 60
    )
    $local:cache = $script:myOffice365Services['OnlineVersionCache']
    $local:entry = $local:cache[$Name]
    if ($null -ne $local:entry -and ([datetime]::Now - $local:entry.Fetched).TotalMinutes -lt $MaxAgeMinutes) {
        return $local:entry.Version
    }
    $local:online = Find-myModule -Name $Name -ErrorAction SilentlyContinue
    $local:ver    = if ($local:online) { [string]$local:online.Version } else { $null }
    $local:cache[$Name] = [PSCustomObject]@{ Version = $local:ver; Fetched = [datetime]::Now }
    return $local:ver
}
