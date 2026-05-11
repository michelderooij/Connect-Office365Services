function Get-ModuleScope {
    param(
        $Module
    )
    If( $Module.ModuleBase -ilike ('{0}*' -f (Join-Path -Path $ENV:HOMEDRIVE -ChildPath $ENV:HOMEPATH))) {
        'CurrentUser'
    }
    Else {
        'AllUsers'
    }
}
