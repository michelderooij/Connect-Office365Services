function Get-Office365AccessToken {
    <#
    .SYNOPSIS
    Acquires an OAuth2 access token for the specified resource scope using MSAL.NET directly.

    .DESCRIPTION
    Loads Microsoft.Identity.Client from whichever installed module ships it (Az.Accounts,
    Microsoft.Graph.Authentication, or MSAL.PS). If none is found, emits a one-time warning
    and returns $null so callers fall back to legacy PSCredential authentication.

    On first call a PublicClientApplication is built and cached in the module state so the
    MSAL token cache persists across calls within the same session. Subsequent calls for the
    same account are satisfied silently without a browser prompt.

    On successful interactive login the function populates:
        Office365UPN  — user principal name extracted from the token account
        MsalAccount   — MSAL IAccount object used for subsequent silent requests
        TenantID      — tenant GUID from the token (only when not already set)

    .PARAMETER Scope
    The OAuth2 scope / resource URI to request, e.g. 'https://graph.microsoft.com/.default'
    or 'https://outlook.office365.com/.default'.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Scope
    )

    # Ensure Microsoft.Identity.Client is loaded. It ships inside Az.Accounts,
    # Microsoft.Graph.Authentication, and the archived MSAL.PS module; load it from
    # whichever is installed rather than requiring a separate download.
    if (-not ([System.AppDomain]::CurrentDomain.GetAssemblies() |
              Where-Object { $_.GetName().Name -eq 'Microsoft.Identity.Client' })) {

        $local:searchBases = @(
            (Get-Module Az.Accounts               -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1).ModuleBase,
            (Get-Module Microsoft.Graph.Authentication -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1).ModuleBase,
            (Get-Module MSAL.PS                   -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1).ModuleBase
        ) | Where-Object { $_ }

        foreach ($local:base in $local:searchBases) {
            $local:dll = Get-ChildItem -Path $local:base -Filter 'Microsoft.Identity.Client.dll' `
                             -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($local:dll) {
                try { Add-Type -Path $local:dll.FullName -ErrorAction Stop } catch { }
                break
            }
        }
    }

    # If still unavailable, warn once and return $null so callers use legacy auth.
    if (-not ([System.AppDomain]::CurrentDomain.GetAssemblies() |
              Where-Object { $_.GetName().Name -eq 'Microsoft.Identity.Client' })) {
        if (-not $script:myOffice365Services['MsalNetWarned']) {
            Write-Warning ('MSAL.NET (Microsoft.Identity.Client) not detected. ' +
                'Install Az, Microsoft.Graph, or MSAL.PS for modern authentication; ' +
                'you may be prompted to sign in multiple times.')
            $script:myOffice365Services['MsalNetWarned'] = $true
        }
        return $null
    }

    $local:clientId = $script:myOffice365Services['MsalClientId']

    # Build or reuse the PublicClientApplication (the MSAL token cache lives on the app instance).
    if (-not $script:myOffice365Services['MsalApp']) {
        $script:myOffice365Services['MsalApp'] =
            [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($local:clientId).
                WithAuthority('https://login.microsoftonline.com/common').
                WithDefaultRedirectUri().
                Build()
    }

    $local:app     = $script:myOffice365Services['MsalApp']
    $local:account = $script:myOffice365Services['MsalAccount']
    $local:token   = $null

    # Try silent acquisition first (uses the MSAL token cache; no browser prompt).
    if ($local:account) {
        try {
            $local:token = $local:app.AcquireTokenSilent([string[]]@($Scope), $local:account).
                               ExecuteAsync().GetAwaiter().GetResult()
        }
        catch { }
    }

    # Interactive fallback — opens the system browser for sign-in.
    if (-not $local:token) {
        try {
            $local:token = $local:app.AcquireTokenInteractive([string[]]@($Scope)).
                               ExecuteAsync().GetAwaiter().GetResult()
        }
        catch {
            Write-Warning ('Modern auth token acquisition failed: {0}' -f $_.Exception.Message)
            return $null
        }

        if ($local:token) {
            $script:myOffice365Services['MsalAccount']  = $local:token.Account
            $script:myOffice365Services['Office365UPN'] = $local:token.Account.Username
            if (-not $script:myOffice365Services['TenantID']) {
                $script:myOffice365Services['TenantID'] = $local:token.TenantId
            }
        }
    }

    if ($local:token) { return $local:token.AccessToken }
    return $null
}

