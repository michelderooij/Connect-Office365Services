function Get-OnPremisesCredential {
    $local:prevuser= if ($script:myOffice365Services['OnPremisesCredential']) {
        $script:myOffice365Services['OnPremisesCredential'].UserName
    }
    else {
        ''
    }
    $script:myOffice365Services['OnPremisesCredential'] = Get-Credential -UserName $local:prevUser -Message 'Please enter your on-premises credentials' -Title 'On-Premises Credentials' -ErrorAction SilentlyContinue
}
