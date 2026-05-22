function Get-ModuleVersionInfo {
    param(
        $Module
    )
    $Module = $Module | Select-Object -First 1

    # Fast path: prerelease tag already available in the PSModuleInfo object loaded
    # by PSResourceGet / PowerShellGet (avoids a Get-Content disk read per module).
    $local:prerelease = $null
    if ($null -ne $Module.PrivateData -and $Module.PrivateData -is [hashtable]) {
        $local:psData = $Module.PrivateData['PSData']
        if ($null -ne $local:psData -and $local:psData -is [hashtable]) {
            $local:prerelease = $local:psData['Prerelease']
        }
    }
    if ($local:prerelease) {
        return ('{0}-{1}' -f $Module.Version.ToString(), $local:prerelease)
    }

    # Slow path: parse the .psd1 file for modules not installed via the Gallery
    # (e.g. manual installs, modules whose PrivateData uses a non-hashtable shape).
    $ModuleManifestPath = $Module.Path
    if ($ModuleManifestPath) {
        if (!(Test-Path -Path $ModuleManifestPath)) {
            return $Module.Version.ToString()
        }
        $ModuleManifestContent = Get-Content -Path $ModuleManifestPath
        $preReleaseInfo = $ModuleManifestContent -match "Prerelease = '(.*)'"
        if ($preReleaseInfo) {
            $preReleaseVersion = $preReleaseInfo[0].Split('=')[1].Trim().Trim("'")
            if ($preReleaseVersion) {
                return ('{0}-{1}' -f $Module.Version.ToString(), $preReleaseVersion)
            }
        }
    }
    return $Module.Version.ToString()
}
