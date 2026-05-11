function Connect-AIP {
    If(!( Get-Module -Name AIPService -ListAvailable)) {
        Import-Module -Name AIPService -ErrorAction SilentlyContinue
    }
    If ( Get-Command -Name Connect-AipService -ErrorAction SilentlyContinue) {
        If ( !($script:myOffice365Services['Office365Credentials'])) { Get-Office365Credentials }
        Write-Host ('Connecting to Azure Information Protection using {0}' -f $script:myOffice365Services['Office365Credentials'].UserName)
        Connect-AipService -Credential $script:myOffice365Services['Office365Credentials']
    }
    Else {
        Write-Error -Message 'Cannot connect to Azure Information Protection - module not installed or not loading.'
    }
}
