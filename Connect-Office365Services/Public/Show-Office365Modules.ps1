function Show-Office365Modules {
    param(
        [switch]$AllowPrerelease,
        [switch]$Refresh
    )

    $local:Functions= Get-Office365ModuleInfo
    $local:UsePre = if ($AllowPrerelease.IsPresent) { $true } else { [bool]$script:myOffice365Services['AllowPrerelease'] }
    if ($Refresh) { $script:myOffice365Services['OnlineVersionCache'].Clear() }

    $local:AllInstalled = Get-Module -ListAvailable -ErrorAction SilentlyContinue

    # Single pass: build a module-name→PSModuleInfo map for installed modules.
    # This avoids calling Get-InstalledRepoModule twice per module (once here for
    # the parallel pre-fetch name list, and again in the display loop).
    $local:InstalledMap = @{}
    ForEach ($local:Item in $local:Functions) {
        $local:m = Get-InstalledRepoModule -Name $local:Item.Module -Repo $local:Item.Repo -AllInstalled $local:AllInstalled
        if ($local:m) { $local:InstalledMap[$local:Item.Module] = $local:m }
    }

    # Pre-fetch online versions.
    # PS 7+: run lookups in parallel (ThrottleLimit 10) — reduces ~22 s to ~3 s.
    # PS 5.1: sequential; each lookup populates the cache so repeat calls are instant.
    $local:Cache   = $script:myOffice365Services['OnlineVersionCache']
    $local:ToFetch = $local:InstalledMap.Keys | Where-Object {
        $local:e = $local:Cache[$_]
        $null -eq $local:e -or ([datetime]::Now - $local:e.Fetched).TotalMinutes -ge 60
    }

    if ($local:ToFetch) {
        if ($PSVersionTable.PSVersion.Major -ge 7) {
            $local:UsePSRG = $script:myOffice365Services['PSResourceGet']
            $local:ToFetch | ForEach-Object -Parallel {
                $local:n = $_
                $local:o = if ($using:UsePSRG) {
                    Find-PSResource -Name $local:n -Prerelease:$using:UsePre -ErrorAction SilentlyContinue
                } else {
                    Find-Module -Name $local:n -AllowPrerelease:$using:UsePre -ErrorAction SilentlyContinue
                }
                [PSCustomObject]@{ Name=$local:n; Version=if($local:o){[string]$local:o.Version}else{$null} }
            } -ThrottleLimit 10 | ForEach-Object {
                # Write back to cache in the parent runspace (serial, thread-safe)
                $script:myOffice365Services['OnlineVersionCache'][$_.Name] = [PSCustomObject]@{
                    Version = $_.Version; Fetched = [datetime]::Now
                }
            }
        } else {
            # PS 5.1 sequential path — Get-OnlineModuleVersion caches each result
            foreach ($local:n in $local:ToFetch) { $null = Get-OnlineModuleVersion -Name $local:n }
        }
    }

    # Display loop — all online versions come from the cache (instant lookups)
    ForEach ($local:Item in $local:Functions) {

        $local:Module = $local:InstalledMap[$local:Item.Module]

        If( $local:Module) {

            $local:Version    = Get-ModuleVersionInfo -Module $local:Module
            $local:OnlineVer  = Get-OnlineModuleVersion -Name $local:Item.Module

            Write-Host ('{0} v{1}' -f $local:Item.Description, $local:Version) -NoNewline

            If( $local:OnlineVer) {
                Write-Host (' (Online v{0})' -f $local:OnlineVer) -NoNewline
            }
            Else {
                Write-Host (' (Online N/A)') -NoNewline
            }
            Write-Host (', Scope:{0} - Status is ' -f (Get-ModuleScope -Module $local:Module)) -NoNewline

            If( [string]::IsNullOrEmpty( $local:Version) -or [string]::IsNullOrEmpty( $local:OnlineVer)) {
                Write-Host ('Unknown') -ForegroundColor $script:myConsoleColors.Warning
            }
            Else {
                If( (Compare-TextVersionNumber -Version $local:Version -CompareTo $local:OnlineVer) -eq 1) {
                    Write-Host ('Outdated') -ForegroundColor $script:myConsoleColors.Error
                }
                Else {
                    Write-Host ('OK') -ForegroundColor $script:myConsoleColors.OK
                }
            }
            If( $local:Item.ReplacedBy) {
                Write-Warning ('{0} has been replaced by {1}' -f $local:Item.Module, $local:Item.ReplacedBy)
            }
        }
        Else {
            Write-Host ('{0} not installed' -f $local:Item.Description) -ForegroundColor $script:myConsoleColors.Muted
        }
    }
}

