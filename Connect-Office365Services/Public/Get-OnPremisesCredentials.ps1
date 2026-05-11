function Get-OnPremisesCredentials {
    $local:prevuser= if ($script:myOffice365Services['OnPremisesCredentials']) {
        $script:myOffice365Services['OnPremisesCredentials'].UserName
    } 
    else { 
        ''
    }
    $script:myOffice365Services['OnPremisesCredentials'] = Get-Credential -UserName $local:prevUser -Message 'Please enter your on-premises credentials' -Title 'On-Premises Credentials' -ErrorAction SilentlyContinue
}
