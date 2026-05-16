# Connect-Office365Services

![PowerShell Gallery Downloads](https://img.shields.io/powershellgallery/dt/Connect-Office365Services)
![GitHub Release](https://img.shields.io/github/v/release/michelderooij/Connect-Office365Services)
![GitHub Repo stars](https://img.shields.io/github/stars/michelderooij/Connect-Office365Services)
![GitHub forks](https://img.shields.io/github/forks/michelderooij/Connect-Office365Services)

PowerShell module providing functions to connect to Microsoft 365 online services and Exchange On-Premises, as well as perform module management.

## Getting Started

After importing the module, the following functions are available:

**Connect**
* `Connect-EXO`                     Connects to Exchange Online (was `Connect-ExchangeOnline`)
* `Connect-Exchange`                Connects to Exchange On-Premises (was `Connect-ExchangeOnPremises`)
* `Connect-SCC`                     Connects to Security & Compliance (was `Connect-IPPSSession`)
* `Connect-MSTeams`                 Connects to Microsoft Teams
* `Connect-AIP`                     Connects to Azure Information Protection
* `Connect-SPO`                     Connects to SharePoint Online (was `Connect-SharePointOnline`)
* `Connect-PowerApps`               Connects to Power Apps
* `Connect-MG`                      Connects to Microsoft Graph (`Connect-MgGraph`)
* `Connect-PowerBI`                 Connects to Power BI (`Connect-PowerBIServiceAccount`)
* `Connect-PnP`                     Connects to a SharePoint site via PnP PowerShell; use `-SiteUrl` to specify a site, defaults to tenant root
* `Connect-Office365`               Connects to EXO, Teams, SCC, and SPO by default; use `-Service` to target specific services

**Disconnect & session**
* `Disconnect-Office365`            Disconnects one or more services
* `Get-Office365Session`            Shows the active identity and per-service connection status
* `Test-Office365Connectivity`      Tests reachability of key Microsoft 365 endpoints

**Credentials & identity**
* `Get-Office365Credential`         Gets Microsoft 365 credentials
* `Get-OnPremisesCredential`        Gets On-Premises credentials
* `Get-Office365Tenant`             Gets Microsoft 365 tenant name
* `Get-ExchangeOnPremisesFQDN`      Gets FQDN for Exchange On-Premises
* `Get-TenantID`                    Resolves tenant ID from a domain or UPN

**Environment & preferences**
* `Set-Office365Environment`        Configures endpoints and region (Default, GCC, GCCHigh, DoD, China, Germany, AzurePPE)
* `Get-Office365Services`           Returns current module state
* `Get-Office365ServicesPreferences` Displays all current preference values and the preferences file location
* `Set-Office365ServicesPreferences` Sets persistent user preferences (see [Preferences](#preferences) below)

**Module management**
* `Select-Office365Modules`         Interactively install/uninstall Office 365 PowerShell modules
* `Update-Office365Modules`         Updates installed Office 365 modules
* `Optimize-Office365Modules`       Removes old module versions
* `Show-Office365Modules`           Lists available and installed modules
* `Save-Office365ModuleState`       Saves installed module versions to the preferences file
* `Restore-Office365ModuleState`    Reinstalls modules from the saved state; use `-Recent` for latest versions
* `Export-Office365ModuleConfig`    Exports the preferences file as JSON (`-File`)
* `Import-Office365ModuleConfig`    Imports a previously exported preferences file (`-File`)

### Prerequisites

PowerShell 5.1 or higher. Modules for the services you want to connect to must be installed — use `Select-Office365Modules` to manage them.

### Installing

```powershell
Install-Module -Name Connect-Office365Services
```

Or with [PSResourceGet](https://learn.microsoft.com/powershell/module/microsoft.powershell.psresourceget/):

```powershell
Install-PSResource -Name Connect-Office365Services
```

Then import the module (or add this to your PowerShell profile):

```powershell
Import-Module Connect-Office365Services
```

## Exporting and Importing Preferences

You can back up your preferences and installed module versions to a JSON file, and restore them on another machine or after a clean OS installation.

### Export

`Export-Office365ModuleConfig` first snapshots the currently installed module versions (`Save-Office365ModuleState`) and then writes the full config — preferences and module state — to the specified file:

```powershell
Export-Office365ModuleConfig -File C:\Backup\O365Config.json
```

### Import

Copy the exported file to the new machine (with this module already installed), then import it:

```powershell
Import-Office365ModuleConfig -File C:\Backup\O365Config.json
```

This imports the saved preferences and module configuration from the file.

### Restore modules

After importing the config, reinstall the saved modules at their exact recorded versions:

```powershell
Restore-Office365ModuleState
```

Or install the latest available version of each saved module instead:

```powershell
Restore-Office365ModuleState -Recent
```

## Preferences

Use `Set-Office365ServicesPreferences` to change settings and `Get-Office365ServicesPreferences` to view the current values. Preferences are persisted to `%APPDATA%\Office365Services\config.json`.

| Preference | Default | Description |
|---|---|---|
| `AllowPrerelease` | `$false` | Allow pre-release module versions during install/update |
| `AzureEnvironment` | `Default` | Target cloud environment (see `Set-Office365Environment`) |
| `Scope` | `AllUsers` | Module install scope (`AllUsers` or `CurrentUser`) |
| `ProxyAccessType` | `None` | WinHTTP proxy access type for module downloads |
| `NoBanner` | `$false` | Suppress the ASCII art banner on module import |
| `NoQuote` | `$false` | Suppress the random quote on module import |
| `NoReport` | `$false` | Suppress the installed-module list on module import |
| `NoAutoConnect` | `$false` | Suppress automatic credential prompts in connect functions; credentials must be cached first via `Get-Office365Credential` |

```powershell
# View current preferences
Get-Office365ServicesPreferences

# Suppress banner and the module list on import
Set-Office365ServicesPreferences -NoBanner $true -NoReport $true

# Require credentials to be cached before connecting
Set-Office365ServicesPreferences -NoAutoConnect $true
```

## Breaking Changes Per v4.0

Version 4.0 converted Connect-Office365Services from a standalone script into a proper PowerShell module. The following changes may require updating existing scripts or profiles that used the older `.ps1` file:
* The script was previously downloaded and to be dot-sourced. A module is installed from PowerShell Gallery and imported.
* Internal helper functions are now private, e.g. Get-myModule, Find-myModule a.o.
* The module state variable is private to the module (myOffice365Services).
* The `USGovernment` environment was renamed to `GCC`.
* `Connect-ComplianceCenter` was removed in v3.45; use `Connect-IPPSSession` instead.

The following connect functions had to be renamed to avoid collisions with cmdlets of the same name from first party modules, due to how
modules handle conflicts compared to dot-sourcing scripts (eg. need to explicitly specify AllowClobber versus silent overloading):

| New cmdlet | Replaces | Connects to |
|---|---|---|
| `Connect-EXO` | `Connect-ExchangeOnline` | Exchange Online |
| `Connect-Exchange` | `Connect-ExchangeOnPremises` | Exchange On-Premises |
| `Connect-SCC` | `Connect-IPPSSession` | Security & Compliance (IPPS) |
| `Connect-SPO` | `Connect-SharePointOnline` | SharePoint Online |


## Changelog

See [CHANGELOG.md](CHANGELOG.md) for the full version history.


## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) for details.
