function Get-ModuleInstallPath {
    <#
    .SYNOPSIS
    Returns the installation folder path(s) for a given module name and optional version
    by scanning all directories in $env:PSModulePath.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$Name,
        [string]$Version
    )
    $env:PSModulePath -split [System.IO.Path]::PathSeparator | ForEach-Object {
        $local:ModuleFolderBase = Join-Path -Path $_ -ChildPath $Name
        If( $Version) {
            $local:Candidate = Join-Path -Path $local:ModuleFolderBase -ChildPath $Version
            If( Test-Path -Path $local:Candidate -PathType Container) {
                $local:Candidate
            }
        }
        Else {
            If( Test-Path -Path $local:ModuleFolderBase -PathType Container) {
                $local:ModuleFolderBase
            }
        }
    }
}
