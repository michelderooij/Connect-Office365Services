#Requires -Version 5.0

$local:ModuleVersion = '4.0'

# ── Load Private functions ────────────────────────────────────────────────────
$local:PrivateFunctions = Get-ChildItem -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Private') -Filter '*.ps1' -Recurse -ErrorAction SilentlyContinue
foreach ($local:Function in $local:PrivateFunctions) {
    try {
        . $local:Function.FullName
    }
    catch {
        Write-Error ('Failed to import private function {0}: {1}' -f $local:Function.BaseName, $_.Exception.Message)
    }
}

# ── Load Public functions ─────────────────────────────────────────────────────
$local:PublicFunctions = Get-ChildItem -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Public') -Filter '*.ps1' -Recurse -ErrorAction SilentlyContinue
foreach ($local:Function in $local:PublicFunctions) {
    try {
        . $local:Function.FullName
    }
    catch {
        Write-Error ('Failed to import public function {0}: {1}' -f $local:Function.BaseName, $_.Exception.Message)
    }
}

# ── Initialize module state ───────────────────────────────────────────────────
Initialize-ServicesState

# ── Detect PSResourceGet availability ────────────────────────────────────────
Get-myPSResourceGetInstalled

# ── Banner ────────────────────────────────────────────────────────────────────
$local:PSGetModule = Get-Module -Name Microsoft.PowerShell.PSResourceGet -ListAvailable -ErrorAction SilentlyContinue |
    Sort-Object -Property Version -Descending | Select-Object -First 1
$local:PSGetVer = If ($local:PSGetModule) { $local:PSGetModule.Version } Else { 'N/A' }

$local:PackageManagementModule = Get-Module -Name PackageManagement -ListAvailable -ErrorAction SilentlyContinue |
    Sort-Object -Property Version -Descending | Select-Object -First 1
$local:PMMVer = If ($local:PackageManagementModule) { $local:PackageManagementModule.Version } Else { 'N/A' }

$local:IsAdmin = [System.Security.principal.windowsprincipal]::new(
    [System.Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)

Write-Host ('*' * 78)
Write-Host ('Connect-Office365Services v{0}' -f $local:ModuleVersion)
Write-Host ('Source: https://github.com/michelderooij/Connect-Office365Services')
Write-Host ('Environment:{0}, Administrator:{1}, Scope:{2}' -f $script:myOffice365Services['AzureEnvironment'], $local:IsAdmin, $script:myOffice365Services['Scope'])
Write-Host ('PS:{0}, PSResourceGet:{1}, PackageManagement:{2}' -f ($PSVersionTable).PSVersion, $local:PSGetVer, $local:PMMVer)
Write-Host ('*' * 78)

# ── List installed modules ────────────────────────────────────────────────────
$local:Functions = Get-Office365ModuleInfo
$local:Functions | ForEach-Object -Process {
    $local:Item = $_
    # Use Get-Module directly (PSResourceInfo lacks RepositorySourceLocation)
    $local:Module = Get-Module -Name ('{0}' -f $local:Item.Module) -ListAvailable -ErrorAction SilentlyContinue |
        Sort-Object -Property Version -Descending
    $local:Module = $local:Module |
        Where-Object { $_.RepositorySourceLocation -and ([System.Uri]($_.RepositorySourceLocation)).Authority -ieq ([System.Uri]($local:Item.Repo)).Authority } |
        Select-Object -First 1
    If ($local:Module) {
        $local:Version = Get-ModuleVersionInfo -Module $local:Module
        Write-Host ('Found {0} (v{1})' -f $local:Item.Description, $local:Version)
        If ($local:Item.ReplacedBy) {
            Write-Warning ('{0} replaced by {1}' -f $local:Item.Module, $local:Item.ReplacedBy)
        }
    }
}

# ── Random quote ──────────────────────────────────────────────────────────────
$local:Quotes = @(
    "You are standing in an open field west of a white house, with a boarded front door. There is a small mailbox here.",
    "You wake up. The room is spinning very gently round your head.`nOr at least it would be if you could see it which you can't. It is pitch black.",
    "You are in a comfortable tunnel like hall. To the east there is the round green door.",
    "You are standing at the end of a road before a small brick building. Around you is a forest.`nA small stream flows out of the building and down a gully.",
    "Shall we play a game?",
    "Request access to CLU program.",
    "You are in a clearing, with a forest surrounding you on all sides. A path leads north."
)
Write-Host ('{0}{1}' -f [System.Environment]::NewLine, ($local:Quotes | Get-Random))
