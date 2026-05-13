function Connect-SPO {
    If (-not (Get-Module -Name Microsoft.Online.SharePoint.PowerShell -ListAvailable)) {
        Import-Module -Name Microsoft.Online.SharePoint.PowerShell -ErrorAction SilentlyContinue
    }
    If (Get-Command -Name Connect-SPOService -ErrorAction SilentlyContinue) {
        If ( -not $script:myOffice365Services['Office365UPN'] -and -not $script:myOffice365Services['Office365Credentials']) {
            Get-Office365Credentials
        }
        $local:upn = if ($script:myOffice365Services['Office365UPN']) {
            $script:myOffice365Services['Office365UPN']
        } else {
            $script:myOffice365Services['Office365Credentials'].UserName
        }
        # Derive the tenant name from the UPN when not already known
        If ($local:upn -like '*.onmicrosoft.com') {
            $script:myOffice365Services['Office365Tenant'] = $local:upn.Substring($local:upn.IndexOf('@') + 1).Replace('.onmicrosoft.com', '')
        }
        ElseIf (-not $script:myOffice365Services['Office365Tenant']) {
            Get-Office365Tenant
        }
        Write-Host 'Connecting to SharePoint Online ..'
        $Parms = @{
            Url    = 'https://{0}-admin.sharepoint.com' -f $script:myOffice365Services['Office365Tenant']
            Region = $script:myOffice365Services['SharePointRegion']
        }
        Connect-SPOService @Parms
    }
    Else {
        Write-Error -Message 'Cannot connect to SharePoint Online - module not installed or not loading.'
    }
}
