function Get-ModuleVersionInfo {
    param(
        $Module
    )
    $Module= $Module | Select-Object -First 1
    $ModuleManifestPath = $Module.Path
    If( $ModuleManifestPath) {
        $isModuleManifestPathValid = Test-Path -Path $ModuleManifestPath
        If(!( $isModuleManifestPathValid)) {
            # Module manifest path invalid, skipping extracting prerelease info
            $ModuleVersion= $Module.Version.ToString()
        }
        Else {
            $ModuleManifestContent = Get-Content -Path $ModuleManifestPath
            $preReleaseInfo = $ModuleManifestContent -match "Prerelease = '(.*)'"
            If( $preReleaseInfo) {
                $preReleaseVersion= $preReleaseInfo[0].Split('=')[1].Trim().Trim("'")
                If( $preReleaseVersion) {
                    $ModuleVersion= ('{0}-{1}' -f $Module.Version.ToString(), $preReleaseVersion)
                }
                Else {
                    $ModuleVersion= $Module.Version.ToString()
                }
            }
            Else {
                $ModuleVersion= $Module.Version.ToString()
            }
        }
    }
    Else {
        $ModuleVersion= $Module.Version.ToString()
    }
    $ModuleVersion
}
