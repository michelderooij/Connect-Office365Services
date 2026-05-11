function Set-WindowTitle {
    If ($host.ui.RawUI.WindowTitle -and $script:myOffice365Services['TenantID']) {
        $local:PromptPrefix = ''
        $ThisPrincipal = New-Object System.Security.Principal.WindowsPrincipal([System.Security.Principal.WindowsIdentity]::GetCurrent())
        if ($ThisPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
            $local:PromptPrefix = 'Administrator:'
        }
        # Use modern auth UPN when available; fall back to legacy credential username
        $local:displayName = if ($script:myOffice365Services['Office365UPN']) {
            $script:myOffice365Services['Office365UPN']
        } elseif ($script:myOffice365Services['Office365Credential']) {
            $script:myOffice365Services['Office365Credential'].UserName
        } else { '' }
        $local:Title = '{0}{1} connected to Tenant ID {2}' -f $local:PromptPrefix, $local:displayName, $script:myOffice365Services['TenantID']
        $host.ui.RawUI.WindowTitle = $local:Title
    }
}
