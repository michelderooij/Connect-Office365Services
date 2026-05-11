function Show-Office365Modules {
    param(
        [switch]$AllowPrerelease
    )

    $local:Functions= Get-Office365ModuleInfo
    if ($PSBoundParameters.ContainsKey('AllowPrerelease')) {
        $script:myOffice365Services['AllowPrerelease'] = $AllowPrerelease.IsPresent
    }

    ForEach ( $local:Item in $local:Functions) {

        # Use Get-Module directly for RepositorySourceLocation check (PSResourceInfo lacks this property)
        $local:Module= Get-Module -Name ('{0}' -f $local:Item.Module) -ListAvailable | Sort-Object -Property Version -Descending
        $local:Module= $local:Module | Where-Object { $_.RepositorySourceLocation -and ([System.Uri]($_.RepositorySourceLocation)).Authority -ieq ([System.Uri]($local:Item.Repo)).Authority } | Select-Object -First 1

        If( $local:Module) {

            $local:Version = Get-ModuleVersionInfo -Module $local:Module
            Write-Host ('{0}: Local v{1}' -f $local:Item.Description, $Local:Version) -NoNewline
            $OnlineModule = Find-myModule -Name $local:Item.Module -ErrorAction SilentlyContinue

            If( $OnlineModule) {
                Write-Host (', Online v{0}' -f $OnlineModule.version) -NoNewline
            }
            Else {
                Write-Host (', Online N/A') -NoNewline
            }
            Write-Host (', Scope:{0} Status:' -f (Get-ModuleScope -Module $local:Module)) -NoNewline

            If( [string]::IsNullOrEmpty( $local:Version) -or [string]::IsNullOrEmpty( $OnlineModule.version)) {
                Write-Host ('Unknown')
            }
            Else {
                If( (Compare-TextVersionNumber -Version $local:Version -CompareTo $OnlineModule.version) -eq 1) {
                    Write-Host ('Outdated') -ForegroundColor Red
                }
                Else {
                    Write-Host ('OK') -ForegroundColor Green
                }
            }
            If( $local:Item.ReplacedBy) {
                Write-Warning ('{0} has been replaced by {1}' -f $local:Item.Module, $local:Item.ReplacedBy)
            }
        }
        Else {
            Write-Host ('{0} not found ({1})' -f $local:Item.Description, $local:Item.Repo) -ForegroundColor DarkGray
        }
    }
}
