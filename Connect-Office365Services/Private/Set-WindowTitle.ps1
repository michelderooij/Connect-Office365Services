function Set-WindowTitle {
    If( $host.ui.RawUI.WindowTitle -and $script:myOffice365Services['TenantID']) {
        $local:PromptPrefix= ''
        $ThisPrincipal= new-object System.Security.principal.windowsprincipal( [System.Security.Principal.WindowsIdentity]::GetCurrent())
        if( $ThisPrincipal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator)) {
            $local:PromptPrefix= 'Administrator:'
        }
        $local:Title= '{0}{1} connected to Tenant ID {2}' -f $local:PromptPrefix, $script:myOffice365Services['Office365Credentials'].UserName, $script:myOffice365Services['TenantID']
        $host.ui.RawUI.WindowTitle = $local:Title
    }
}
