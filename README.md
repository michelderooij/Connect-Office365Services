# Connect-Office365Services

![PowerShell Gallery Downloads](https://img.shields.io/powershellgallery/dt/Connect-Office365Services)
![GitHub Release](https://img.shields.io/github/v/release/michelderooij/Connect-Office365Services)
![GitHub Repo stars](https://img.shields.io/github/stars/michelderooij/Connect-Office365Services?style=flat-square)
![GitHub forks](https://img.shields.io/github/forks/michelderooij/Connect-Office365Services?style=flat-square)

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

Explicit `-Credential` use with `Connect-EXO` or `Connect-SCC` displays a warning
when the installed Exchange Online Management module supports that parameter.
Microsoft announced that new Exchange Online PowerShell module versions released
starting December 2026 will no longer support this parameter. Use interactive
sign-in, certificate-based app-only authentication, or managed identity as appropriate.
See the [Microsoft announcement](https://techcommunity.microsoft.com/blog/exchange/deprecation-of-the--credential-parameter-in-exchange-online-powershell/4494584).
On newer versions without `-Credential`, `Connect-EXO` rejects the parameter before
it can warn; `Connect-SCC` warns and reports that the parameter is unsupported.

When no identity is cached, `Connect-EXO` and `Connect-SCC` use Exchange Online
Management's native interactive sign-in rather than acquiring a Graph token first.
`NoAutoConnect` still blocks this automatic sign-in. Native Exchange sign-in does not
populate this module's MSAL cache for Graph, SharePoint, or other services.

`Connect-SPO` accepts `-Url` to override the cached tenant's administration URL and
`-Credential` to use explicit credentials instead of a cached identity or token:

```powershell
Connect-SPO -Url https://contoso-admin.sharepoint.com -Credential (Get-Credential)
```

If `-Credential` uses a custom domain and no tenant is cached, also supply `-Url`.

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
* `Move-Office365Modules`           Moves installed modules between AllUsers and CurrentUser scope
* `Show-Office365Modules`           Lists available and installed modules
* `Save-Office365ModuleState`       Saves installed module versions to the preferences file
* `Restore-Office365ModuleState`    Reinstalls modules from the saved state; use `-Recent` for latest versions
* `Export-Office365ModuleConfig`    Exports the preferences file as JSON (`-File`)
* `Import-Office365ModuleConfig`    Imports a previously exported preferences file (`-File`)

### Prerequisites

PowerShell 5.1 or higher. Modules for the services you want to connect to must be installed — use `Select-Office365Modules` to manage them.

On import, installed MSOnline, AzureAD, and AzureADPreview modules are reported as retired. On Windows PowerShell 5.x, Microsoft Graph PowerShell v2 also triggers a retirement notice. Microsoft announced a roughly 12-month retirement period beginning in September 2026 and a v3 release planned for Q4 2026 that supports PowerShell 7.x only. See the [Microsoft announcement](https://devblogs.microsoft.com/microsoft365dev/investing-in-a-more-reliable-microsoft-graph-powershell-experience/).

Installed Microsoft.Graph.Entra and Microsoft.Graph.Entra.Beta modules are reported as retired and replaced by Microsoft.Entra and Microsoft.Entra.Beta, respectively.

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

### Troubleshooting MSAL assembly conflicts

If an Exchange connection reports that `Microsoft.Identity.Client.dll` has a
manifest mismatch, restart PowerShell. Removing a module does not unload its
assemblies. In a fresh `pwsh -NoProfile` process, connect to Exchange before loading
other authentication modules:

```powershell
Import-Module Connect-Office365Services
Connect-EXO -UserPrincipalName 'admin@example.com'
```

EXO/SCC no longer preload external MSAL assemblies for sign-in, but incompatible
assemblies loaded by other code can still conflict. Use separate PowerShell
processes for incompatible module combinations; do not replace vendor DLLs.

On PowerShell 7, when the shared token helper needs to load MSAL, it first tries the
netCore DLL bundled with ExchangeOnlineManagement, preferring the loaded Exchange module
version over other installed versions. It falls back to Az.Accounts, Graph
Authentication, or MSAL.PS if necessary. An already-loaded MSAL assembly is reused;
this preference cannot resolve conflicts already present in the process. PowerShell
5.1 retains the original Az.Accounts, Graph Authentication, and MSAL.PS provider order.

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

MSOnline, AzureAD, AzureADPreview, Microsoft.Graph.Entra and Microsoft.Graph.Entra.Beta are no longer managed. Entries in older saved states are skipped with a warning; installed copies are not removed.

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

Suppressing the banner or module report only affects displayed output; PSResourceGet detection still runs. If a private or public function file fails to load, module import stops and reports the failure.

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
