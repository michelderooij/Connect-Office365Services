function Get-InstalledRepoModule {
    <#
    .SYNOPSIS
    Returns the highest installed version of a module whose repository matches
    the supplied Repo URI authority. Accepts a pre-fetched module list to avoid
    repeated Get-Module -ListAvailable filesystem scans inside loops.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$Name,
        [Parameter(Mandatory=$true)]
        [string]$Repo,
        # Optional: pass result of (Get-Module -ListAvailable) to avoid repeated scans.
        $AllInstalled = $null
    )
    $local:candidates = If( $AllInstalled) {
        $AllInstalled | Where-Object { $_.Name -eq $Name }
    } Else {
        Get-Module -Name $Name -ListAvailable -ErrorAction SilentlyContinue
    }
    $local:candidates |
        Sort-Object -Property Version -Descending |
        Where-Object { $_.RepositorySourceLocation -and ([System.Uri]($_.RepositorySourceLocation)).Authority -ieq ([System.Uri]($Repo)).Authority } |
        Select-Object -First 1
}
