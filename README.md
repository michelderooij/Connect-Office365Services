# Connect-Office365Services

PowerShell module providing functions to connect to Microsoft 365 online services and Exchange On-Premises, as well as perform module management.

## Getting Started

After importing the module, the following functions are available:

**Connect**
* `Connect-ExchangeOnline`          Connects to Exchange Online
* `Connect-ExchangeOnPremises`      Connects to Exchange On-Premises
* `Connect-IPPSSession`             Connects to Security & Compliance (IPPS)
* `Connect-MSTeams`                 Connects to Microsoft Teams
* `Connect-AIP`                     Connects to Azure Information Protection
* `Connect-SharePointOnline`        Connects to SharePoint Online
* `Connect-PowerApps`               Connects to Power Apps
* `Connect-Office365`               Connects to all configured services

**Credentials & identity**
* `Get-Office365Credentials`        Gets Microsoft 365 credentials
* `Get-OnPremisesCredentials`       Gets On-Premises credentials
* `Get-Office365Tenant`             Gets Microsoft 365 tenant name
* `Get-ExchangeOnPremisesFQDN`      Gets FQDN for Exchange On-Premises
* `Get-TenantID`                    Resolves tenant ID from a domain or UPN

**Environment & preferences**
* `Set-Office365Environment`        Configures endpoints and region (Default, GCC, GCCHigh, DoD, China, Germany, AzurePPE)
* `Get-Office365Services`           Returns current module state
* `Set-Office365ServicesPreferences` Views or sets user preferences

**Module management**
* `Select-Office365Modules`         Interactively install/uninstall Office 365 PowerShell modules
* `Update-Office365Modules`         Updates installed Office 365 modules
* `Optimize-Office365Modules`       Removes old module versions
* `Show-Office365Modules`           Lists available and installed modules

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

## Changelog

See [CHANGELOG.md](CHANGELOG.md) for the full version history.


## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) for details.
