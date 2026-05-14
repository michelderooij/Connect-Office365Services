#Requires -Version 5.0

$local:ModuleVersion = '4.0.4'

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

# ── Console color struct ─────────────────────────────────────────────────────
$local:PrivateData = $Host.PrivateData
$script:myConsoleColors = [PSCustomObject]@{
    Warning = if ($local:PrivateData -and $local:PrivateData.WarningForegroundColor -is [System.ConsoleColor]) {
        $local:PrivateData.WarningForegroundColor
    }
    else { [System.ConsoleColor]::Yellow }
    Error   = if ($local:PrivateData -and $local:PrivateData.ErrorForegroundColor -is [System.ConsoleColor]) {
        $local:PrivateData.ErrorForegroundColor
    }
    else { [System.ConsoleColor]::Red }
    OK      = [System.ConsoleColor]::Green
    Muted   = [System.ConsoleColor]::White
}

# ── Detect PSResourceGet availability ────────────────────────────────────────
# NOTE: called after the AllInstalled pre-fetch so it re-uses the cached list
#       instead of triggering a second filtered Get-Module -ListAvailable scan.

# ── Banner ────────────────────────────────────────────────────────────────────
$local:AllInstalled = Get-Module -ListAvailable -ErrorAction SilentlyContinue
Get-myPSResourceGetInstalled -AllInstalled $local:AllInstalled
$local:PSGetModule = $local:AllInstalled | Where-Object { $_.Name -eq 'Microsoft.PowerShell.PSResourceGet' } |
Sort-Object -Property Version -Descending | Select-Object -First 1
$local:PSGetVer = if ($local:PSGetModule) { $local:PSGetModule.Version } else { 'N/A' }

$local:PackageManagementModule = $local:AllInstalled | Where-Object { $_.Name -eq 'PackageManagement' } |
Sort-Object -Property Version -Descending | Select-Object -First 1
$local:PMMVer = if ($local:PackageManagementModule) { $local:PackageManagementModule.Version } else { 'N/A' }

$local:IsAdmin = Test-IsAdministrator

if (-not $script:myOffice365Services['NoBanner']) {
    Write-Host ('***********************************************************************')
    Write-Host (' _____                     _       _____ ___ ___ _         ___ ___ ___')
    Write-Host ('|     |___ ___ ___ ___ ___| |_ ___|     |  _|  _|_|___ ___|_  |  _|  _|')
    Write-Host ('|   --| . |   |   | -_|  _|  _|___|  |  |  _|  _| |  _| -_|_  | . |_  |')
    Write-Host ('|_____|___|_|_|_|_|___|___|_|     |_____|_| |_| |_|___|___|___|___|___|')
    Write-Host ('***********************************************************************')
    Write-Host ('Connect-Office365Services v{0}' -f $local:ModuleVersion)
    Write-Host ('https://github.com/michelderooij/Connect-Office365Services')
    Write-Host ('Environment:{0}, Administrator: {1}, Scope: {2}' -f $script:myOffice365Services['AzureEnvironmentName'], $local:IsAdmin, $script:myOffice365Services['Scope'])
    Write-Host ('PS:{0}, PSResourceGet: {1}, PackageManagement: {2}' -f ($PSVersionTable).PSVersion, $local:PSGetVer, $local:PMMVer)
}

# ── List installed modules ────────────────────────────────────────────────────
$local:Functions = Get-Office365ModuleInfo
$local:Functions | ForEach-Object -Process {
    $local:Item = $_
    $local:Module = Get-InstalledRepoModule -Name $local:Item.Module -Repo $local:Item.Repo -AllInstalled $local:AllInstalled
    if ($local:Module) {
        $local:Version = Get-ModuleVersionInfo -Module $local:Module
        Write-Host ('Found {0} (v {1})' -f $local:Item.Description, $local:Version)
        if ($local:Item.ReplacedBy) {
            Write-Warning ('{0} replaced by {1}' -f $local:Item.Module, $local:Item.ReplacedBy)
        }
    }
}

# ── Random quote ──────────────────────────────────────────────────────────────
if (-not $script:myOffice365Services['NoQuote']) {
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
}
