function Get-TenantID {
    $script:myOffice365Services['TenantID']= Get-TenantIDfromMail $script:myOffice365Services['Office365Credentials'].UserName
    If( $script:myOffice365Services['TenantID']) {
        Write-Host ('TenantID: {0}' -f $script:myOffice365Services['TenantID'])
    }
}
