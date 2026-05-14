function Connect-SPO {
    if (-not (Get-Module -Name Microsoft.Online.SharePoint.PowerShell -ListAvailable)) {
        Import-Module -Name Microsoft.Online.SharePoint.PowerShell -ErrorAction SilentlyContinue
    }
    if (Get-Command -Name Connect-SPOService -ErrorAction SilentlyContinue) {
        if ( -not $script:myOffice365Services['Office365UPN'] -and -not $script:myOffice365Services['Office365Credential']) {
            Get-Office365Credential
        }
        $local:upn = if ($script:myOffice365Services['Office365UPN']) {
            $script:myOffice365Services['Office365UPN']
        }
        else {
            $script:myOffice365Services['Office365Credential'].UserName
        }
        # Derive the tenant name from the UPN when not already known
        if ($local:upn -like '*.onmicrosoft.com') {
            $script:myOffice365Services['Office365Tenant'] = $local:upn.Substring($local:upn.IndexOf('@') + 1).Replace('.onmicrosoft.com', '')
        }
        elseif (-not $script:myOffice365Services['Office365Tenant']) {
            Get-Office365Tenant
        }
        Write-Host 'Connecting to SharePoint Online ..'
        $Parms = @{
            Url    = 'https://{0}-admin.sharepoint.com' -f $script:myOffice365Services['Office365Tenant']
            Region = $script:myOffice365Services['SharePointRegion']
        }
        Connect-SPOService @Parms
    }
    else {
        Write-Error -Message 'Cannot connect to SharePoint Online - module not installed or not loading.'
    }
}
