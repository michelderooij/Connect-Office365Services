function Get-OnPremisesCredentials {
    $script:myOffice365Services['OnPremisesCredentials'] = $host.ui.PromptForCredential('On-Premises Credentials', 'Please Enter Your On-Premises Credentials', '', '')
}
