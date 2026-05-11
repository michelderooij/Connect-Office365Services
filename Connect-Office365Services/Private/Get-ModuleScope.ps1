function Get-ModuleScope {
    param(
        $Module
    )
    If( $Module.ModuleBase -ilike ('{0}*' -f [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::UserProfile))) {
        'CurrentUser'
    }
    Else {
        'AllUsers'
    }
}
