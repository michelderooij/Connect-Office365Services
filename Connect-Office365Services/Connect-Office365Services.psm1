#Requires -Version 5.0

$local:ModuleVersion = '4.0.8'

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

# ── Detect PSResourceGet + gather module inventory ────────────────────────────
# Get-Module -ListAvailable scans every module directory (~6-12 s with 400+ modules).
# Instead, pre-fetch the report module list (cheap in-memory JSON parse) and pass its
# names as a -Name filter so PowerShell resolves only the specific module folders.
#   - NoReport=false  → targeted query for ~30 report modules (+ banner modules when needed)
#   - NoBanner=false, NoReport=true → targeted 2-module query (banner version numbers only)
#   - Both true       → no scan at all

# Pre-load the Office 365 module list now (JSON literal, zero I/O) so the name array
# drives the targeted Get-Module query and the result is reused in the report loop.
$local:ReportFunctions = $null
if (-not $script:myOffice365Services['NoReport']) {
    $local:ReportFunctions = Get-Office365ModuleInfo
}

$local:AllInstalled = $null
if (-not $script:myOffice365Services['NoReport']) {
    # Targeted scan: only the module names the report loop will query.
    # Get-Module -Name resolves specific module folders rather than enumerating all directories.
    $local:ScanNames = @($local:ReportFunctions.Module)
    if (-not $script:myOffice365Services['NoBanner']) {
        # Fold banner modules into the same single query to avoid a second scan.
        $local:ScanNames += 'Microsoft.PowerShell.PSResourceGet', 'PackageManagement'
    }
    $local:AllInstalled = Get-Module -Name $local:ScanNames -ListAvailable -ErrorAction SilentlyContinue
}
elseif (-not $script:myOffice365Services['NoBanner']) {
    # Targeted query — sufficient for banner version numbers.
    $local:AllInstalled = Get-Module -Name 'Microsoft.PowerShell.PSResourceGet', 'PackageManagement' `
        -ListAvailable -ErrorAction SilentlyContinue
}

Get-myPSResourceGetInstalled -AllInstalled $local:AllInstalled
$local:PSGetModule = if ($local:AllInstalled) {
    $local:AllInstalled | Where-Object { $_.Name -eq 'Microsoft.PowerShell.PSResourceGet' } |
    Sort-Object -Property Version -Descending | Select-Object -First 1
}
$local:PSGetVer = if ($local:PSGetModule) { $local:PSGetModule.Version } else { 'N/A' }

$local:PackageManagementModule = if ($local:AllInstalled) {
    $local:AllInstalled | Where-Object { $_.Name -eq 'PackageManagement' } |
    Sort-Object -Property Version -Descending | Select-Object -First 1
}
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
if (-not $script:myOffice365Services['NoReport']) {
    $local:ReportFunctions | ForEach-Object -Process {
        $local:Item = $_
        $local:Module = Get-InstalledRepoModule -Name $local:Item.Module -Repo $local:Item.Repo -AllInstalled $local:AllInstalled
        if ($local:Module) {
            $local:Version = Get-ModuleVersionInfo -Module $local:Module
            Write-Host ('Found {0} (v{1})' -f $local:Item.Description, $local:Version)
            if ($local:Item.ReplacedBy) {
                Write-Warning ('{0} replaced by {1}' -f $local:Item.Module, $local:Item.ReplacedBy)
            }
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
