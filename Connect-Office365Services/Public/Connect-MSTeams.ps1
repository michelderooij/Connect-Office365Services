function Connect-MSTeams {
    If(!( Get-Module -Name MicrosoftTeams -ListAvailable)) {
        Import-Module -Name MicrosoftTeams -ErrorAction SilentlyContinue
    }
    If ( Get-Command -Name Connect-MicrosoftTeams -ErrorAction SilentlyContinue) {
        If ( !($script:myOffice365Services['Office365Credentials'])) { Get-Office365Credentials }
        Write-Host ('Connecting to Microsoft Teams using {0} ..' -f $script:myOffice365Services['Office365Credentials'].UserName)
        # Fixed: 'TenantId' corrected to 'TenantID' to match the actual hashtable key
        Connect-MicrosoftTeams -AccountId ($script:myOffice365Services['Office365Credentials']).UserName -TenantId $script:myOffice365Services['TenantID']
    }
    Else {
        Write-Error -Message 'Cannot connect to Microsoft Teams - module not installed or not loading.'
    }
}
