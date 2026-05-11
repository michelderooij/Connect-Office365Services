function Connect-PowerApps {
    If (-not (Get-Module -Name Microsoft.PowerApps.PowerShell -ListAvailable)) {
        Import-Module -Name Microsoft.PowerApps.PowerShell -ErrorAction SilentlyContinue
    }
    If (-not (Get-Module -Name Microsoft.PowerApps.Administration.PowerShell -ListAvailable)) {
        Import-Module -Name Microsoft.PowerApps.Administration.PowerShell -ErrorAction SilentlyContinue
    }
    If (Get-Command -Name Add-PowerAppsAccount -ErrorAction SilentlyContinue) {
        If ( -not $script:myOffice365Services['Office365UPN'] -and -not $script:myOffice365Services['Office365Credential']) {
            Get-Office365Credential
        }
        $local:upn = if ($script:myOffice365Services['Office365UPN']) {
            $script:myOffice365Services['Office365UPN']
        } else {
            $script:myOffice365Services['Office365Credential'].UserName
        }
        Write-Host ('Connecting to PowerApps using {0} ..' -f $local:upn)
        Add-PowerAppsAccount -Username $local:upn
    }
    Else {
        Write-Error -Message 'Cannot connect to PowerApps - problem loading module.'
    }
}
