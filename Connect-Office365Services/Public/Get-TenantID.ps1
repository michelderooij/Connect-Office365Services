function Get-TenantID {
    # Prefer modern auth UPN; fall back to legacy PSCredential username
    $local:upn = if ($script:myOffice365Services['Office365UPN']) {
        $script:myOffice365Services['Office365UPN']
    } elseif ($script:myOffice365Services['Office365Credentials']) {
        $script:myOffice365Services['Office365Credentials'].UserName
    }

    if ($local:upn) {
        # Skip HTTP lookup when TenantID already populated (e.g. from MSAL token)
        if (-not $script:myOffice365Services['TenantID']) {
            $script:myOffice365Services['TenantID'] = Get-TenantIDfromMail $local:upn
        }
        if ($script:myOffice365Services['TenantID']) {
            Write-Host ('TenantID: {0}' -f $script:myOffice365Services['TenantID'])
        }
    }
}
