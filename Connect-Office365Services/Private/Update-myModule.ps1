function Update-myModule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string[]]$Name,
        [ValidateSet('CurrentUser', 'AllUsers')]
        [string]$Scope = $script:myOffice365Services['Scope'],
        # Explicit -Prerelease switch overrides the persisted AllowPrerelease preference.
        # Pass $false when targeting a stable release so that Update-PSResource / Update-Module
        # cannot silently re-install the same prerelease instead of the stable release.
        [switch]$Prerelease
    )
    process {
        $local:usePre = if ($PSBoundParameters.ContainsKey('Prerelease')) { [bool]$Prerelease } else { $script:myOffice365Services['AllowPrerelease'] }
        if ( $script:myOffice365Services['PSResourceGet']) {
            # Use Install-PSResource -Reinstall instead of Update-PSResource because
            # Update-PSResource may silently no-op when upgrading a prerelease to the
            # stable release of the same base version (PSResourceGet compares base versions
            # and considers them equal). -Reinstall bypasses that check entirely.
            # -SkipDependencyCheck prevents reinstalling dependencies that are already
            # present and potentially locked (e.g. PackageManagement in OneDrive paths).
            Install-PSResource -Name $Name -Scope $Scope -Reinstall -AcceptLicense -Prerelease:$local:usePre -TrustRepository -SkipDependencyCheck -ErrorAction Stop
        }
        else {
            # Note: Update-Module does not support -Scope, -AllowClobber, or -AcceptLicense.
            Update-Module -Name $Name -Force -AllowPrerelease:$local:usePre -ErrorAction Stop
        }
    }
}
