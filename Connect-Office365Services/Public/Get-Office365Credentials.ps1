function Get-Office365Credentials {
    $script:myOffice365Services['Office365Credentials'] = $host.ui.PromptForCredential('Office 365 Credentials', 'Please enter your Office 365 credentials', $script:myOffice365Services['Office365Credentials'].UserName, '')
    Get-TenantID
    Set-WindowTitle
}
