function Connect-MSTeams {
    If (-not (Get-Module -Name MicrosoftTeams -ListAvailable)) {
        Import-Module -Name MicrosoftTeams -ErrorAction SilentlyContinue
    }
    If (Get-Command -Name Connect-MicrosoftTeams -ErrorAction SilentlyContinue) {
        If ( -not $script:myOffice365Services['Office365UPN'] -and -not $script:myOffice365Services['Office365Credential']) {
            Get-Office365Credential
        }
        $local:upn = if ($script:myOffice365Services['Office365UPN']) {
            $script:myOffice365Services['Office365UPN']
        } else {
            $script:myOffice365Services['Office365Credential'].UserName
        }
        Write-Host ('Connecting to Microsoft Teams using {0} ..' -f $local:upn)
        Connect-MicrosoftTeams -AccountId $local:upn -TenantId $script:myOffice365Services['TenantID']
    }
    Else {
        Write-Error -Message 'Cannot connect to Microsoft Teams - module not installed or not loading.'
    }
}
