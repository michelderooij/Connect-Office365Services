function Connect-PowerApps {
    If(!( Get-Module -Name Microsoft.PowerApps.PowerShell -ListAvailable)) {
        Import-Module -Name Microsoft.PowerApps.PowerShell -ErrorAction SilentlyContinue
    }
    If(!( Get-Module -Name Microsoft.PowerApps.Administration.PowerShell -ListAvailable)) {
        Import-Module -Name Microsoft.PowerApps.Administration.PowerShell -ErrorAction SilentlyContinue
    }
    If ( Get-Command -Name Add-PowerAppsAccount -ErrorAction SilentlyContinue) {
        If ( !($script:myOffice365Services['Office365Credentials'])) { Get-Office365Credentials }
        Write-Host "Connecting to PowerApps using $($script:myOffice365Services['Office365Credentials'].UserName) .."
        $Parms = @{'Username' = $script:myOffice365Services['Office365Credentials'].UserName }
        Add-PowerAppsAccount @Parms
    }
    Else {
        Write-Error -Message 'Cannot connect to PowerApps - problem loading module.'
    }
}
