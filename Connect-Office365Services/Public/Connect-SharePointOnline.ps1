function Connect-SharePointOnline {
    If(!( Get-Module -Name Microsoft.Online.SharePoint.PowerShell -ListAvailable)) {
        Import-Module -Name Microsoft.Online.SharePoint.PowerShell -ErrorAction SilentlyContinue
    }
    If ( Get-Command -Name Connect-SPOService -ErrorAction SilentlyContinue) {
        If ( !($script:myOffice365Services['Office365Credentials'])) { Get-Office365Credentials }
        If (($script:myOffice365Services['Office365Credentials']).UserName -like '*.onmicrosoft.com') {
            $script:myOffice365Services['Office365Tenant'] = ($script:myOffice365Services['Office365Credentials']).UserName.Substring(($script:myOffice365Services['Office365Credentials']).UserName.IndexOf('@') + 1).Replace('.onmicrosoft.com', '')
        }
        Else {
            If ( !($script:myOffice365Services['Office365Tenant'])) { Get-Office365Tenant }
        }
        Write-Host 'Connecting to SharePoint Online  ..'
        $Parms = @{
            url= 'https://{0}-admin.sharepoint.com' -f $($script:myOffice365Services['Office365Tenant'])
            region= $script:myOffice365Services['SharePointRegion']
        }
        Connect-SPOService @Parms
    }
    Else {
        Write-Error -Message 'Cannot connect to SharePoint Online - module not installed or not loading.'
    }
}
