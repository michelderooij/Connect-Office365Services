# Changelog

## v4.0.2
- Changed: `Connect-ExchangeOnline` renamed to `Connect-EXO` to avoid conflict with ExchangeOnlineManagement.
- Changed: `Connect-IPPSSession` renamed to `Connect-SCC` to avoid conflict with ExchangeOnlineManagement.
- Changed: `Connect-SharePointOnline` renamed to `Connect-SPO` for consistency.
- Changed: `Connect-ExchangeOnPremises` renamed to `Connect-Exchange` for consistency.
- Added: `NoBanner` preference to suppress the ASCII art banner on module import.
- Added: `NoQuote` preference to suppress the random quote on module import.
- Added: `Save-Office365ModuleState` to snapshot installed module versions into the preferences file.
- Added: `Restore-Office365ModuleState` to reinstall modules from the saved snapshot; use `-Recent` to install the latest version instead of the pinned version.
- Added: `Export-Office365ModuleConfig` to export the preferences file as JSON using the `-File` parameter.
- Added: `Import-Office365ModuleConfig` to import a previously exported preferences file using the `-File` parameter; updates live session state immediately.
- Fixed: banner module list output producing `InvalidOperation` errors.
- Fixed: `Save-Office365ServicesPreferences` issue when updating preferences.

## v4.0.1
- Fixed: module install path resolved incorrectly on systems with a non-default home drive.
- Fixed: invalid Exchange On-Premises hostnames are now rejected; prompt repeats until valid.
- Fixed: tenant ID lookup no longer proceeds when the e-mail address is malformed.
- Fixed: `Update-Office365Modules` did not detect modules installed in the AllUsers scope.
- Fixed: `Update-Office365Modules` used the configured scope preference instead of the module's actual install scope.
- Fixed: update errors were silently swallowed when the update cmdlet emitted non-terminating errors.
- Fixed: error messages in update output were blank; now show the actual exception text.
- Fixed: "Updated to version" showed the wrong version after an update.
- Fixed: `Update-Office365Modules` failed for modules installed via `Install-Module` when PSResourceGet is active; falls back to `Install-PSResource -Reinstall`.
- Performance: module import faster; installed module list now read from disk only once.
- Performance: `Show-Office365Modules` and `Update-Office365Modules` run version checks concurrently on PS 7 (~3 s vs ~20 s).
- Performance: online version info cached for 60 minutes; repeat calls within a session are instant.
- Added `-Refresh` switch to `Show-Office365Modules` and `Update-Office365Modules` to force a fresh online version check.
- Get-Office365Credential attempts modern auth using MSAL.NET (Microsoft.Identity.Client) when found, fall back to `PSCredential`

## v4.0
- The script has been converted to a proper PowerShell module and published to PowerShell Gallery.
- Removing old module versions now more reliable.
- Optimize-Office365Modules will retry stuck uninstalls and, as a final step, hard-delete any leftover version folders it finds in the module installation path.
- Fixed: connecting to Exchange On-Premises failed when credentials had already been provided.
- Fixed: connecting to Microsoft Teams failed when a Tenant ID was configured.
- Fixed: opening a Security & Compliance (IPPS) session crashed if no credentials had been entered yet.
- Fixed: looking up a tenant ID by e-mail address crashed on DNS or network errors instead of reporting the failure cleanly.
- Fixed: Connect-Office365 attempting to call two functions that no longer exist
- Fixed: the module list and interactive selection screen failed when PSResourceGet was active.
- Fixed: a startup error when initialising the default Azure environment setting.
- Fixed: the Microsoft365DSC repository URL in the module catalogue contained a typo.
- Added Get-Office365Services return a snapshot of the current module state.
- Added Set-Office365ServicesPreferences to view or set user preferences (AllowPrerelease, AzureEnvironment, Scope, ProxyAccessType).
- Preferences are stored in "%APPDATA%\Office365Services\config.json"
- Select-Office365Modules: replaced two-column layout with a single-column list and no screen clearing.
- Select-Office365Modules: deprecated modules can only be removed when installed; they cannot be freshly installed.
- Renamed USGovernment to GCC (Government Community Cloud on worldwide infrastructure) and fixed endpoints
- Added GCCHigh environment
- added DoD environment (Department of Defense)
- Updated Germany environment

## v3.46
- Changed module check before import to catch issues
- Small cleanup Connect-SPO
- Small textual corrections in synopsis

## v3.45
- Fixed Connect-IPPSSession
- Corrected Connect-ComplianceCenter references, changed to Connect-IPPSSession
- Some cosmetic changes
- Removed redundant module check/import pairs

## v3.44
- Added Quote of the Day like message
- Minor cosmetic changes

## v3.43
- Fixed Connect-ExchangeOnline

## v3.42
- Added error handling to Uninstall-MyModule output error handling

## v3.41
- Fixed parameter usage issue with not using PSResourceGet

## v3.4
- Added using Microsoft.PowerShell.PSResourceGet when available (performance)
- Removed obsolete repository code
- Code cleanup
- Cosmetic changes in output

## v3.3
- Added notice for module replacement (e.g. microsoft.graph.entra > microsoft.entra)
- Module information now stored in JSON for maintainability
- Changed Microsoft.Graph.Entra to Microsoft.Entra
- Changed Microsoft.Graph.Entra.Beta to Microsoft.Entra.Beta
- Removed old ISE entries from module information

## v3.231
- Made dependency checking silent when nothing found

## v3.23
- Updated Optimize-Office365Modules to process dependencies (e.g. Graph)
- Removed Compatibility Adapter for AzureAD PowerShell (predecessor Entra PowerShell)

## v3.22
- Removed MFA/Non-MFA code
- Removed Connect-AzureAD helper function
- Modified Get-TenantId to use OpenId endpoint to read ID using credentials username when available
- Removed ISE menu creation code

## v3.21
- Added Places module
- Added Microsoft.Graph.Entra module
- Added Microsoft.Graph.Entra.Beta module

## v3.20
- Added Optimize-Office365Modules

## v3.19
- Removed SkypeOnlineConnector & ExoPowerShellModule related code

## v3.18
- Added Microsoft.Graph.Beta

## v3.17
- Added Microsoft.Graph.Compatibility.AzureAD (Preview)

## v3.16
- Fixed duplicate module processing (Connect-ComplianceCenter/EXO is in same module)

## v3.15
- Fixed creating ISE menu options for local functions
- Removed Connect-EOP

## v3.14
- Added O365CentralizedAddInDeployment to set of supported modules

## v3.13
- Added ORCA to set of supported modules

## v3.12
- Replaced 'Prerelease' questions with switch

## v3.11
- Fixed header not displaying correct script version

## v3.10
- Added support for WhiteboardAdmin
- Added support for MSIdentityTools
- Removed Microsoft Teams (Test) support (from poshtestgallery)
- Renamed Azure AD v1 to MSOnline to prevent confusion

## v3.01
- Added Preview info when reporting local module info

## v3.00
- Fixed wrongly detecting old modules because mixed native PS module and PSGet cmdlets
- Back to using native PS module management cmdlets
- Startup only reports installed modules; Report now also reports not installed modules
- Some cosmetics

## v2.99
- Added 2 connect helper functions to description

## v2.98
- Fixed ConnectionUri in EXO connection method

## v2.97
- Fixed title for admin roles

## v2.96
- Added Microsoft365DSC module
- Fixed determining current module scope (CurrentUser/AllUsers)

## v2.95
- Added UseRPSSession switch for Connect-ExchangeOnline

## v2.94
- Added AllowClobber to ignore existing cmdlet conflicts when updating modules

## v2.93
- Added cleaning up of module dependencies (e.g. Az)
- Updating will use same scope of installed module
- Showing warning during update when running multiple PowerShell sessions

## v2.92
- Removed duplicate MSCommerce checking

## v2.91
- Removed Microsoft.Graph.Teams.Team module (unlisted at PSGallery)

## v2.90
- Added MSCommerce module
- Added MicrosoftPowerBIMgmt module
- Added Az module

## v2.80
- Fixed updating module using install-package when existing package comes from different repo
- Fixed removal of old modules logic (.100 is newer than .81)
- Improved version handling to properly evaluate Preview modules
- Versions reported are now showing their textual representation, including tags like PreviewX
- Show-Office365Modules output is now more condense

## v2.71

- Revised module updating using Install-Package when available

## v2.70

- Added support for all overloaded Connect-ExchangeOnline parameters from ExchangeOnlineManagement module
- Added PnP.PowerShell module support
- Updated AzureADAuthorizationEndpointUri for Common/GCC
- Removed SharePointPnPPowerShellOnline support
- Removed obsolete code for MFA module presence check

## v2.66

- Reporting change in number of cmdlets after updating

## v2.65

- Fixed connecting to AzureAD using MFA not using provided Username

## v2.64

- Structured Connect-MsTeams

## v2.63

- Changed default ProxyAccessType to None

## v2.62

- Added -ProxyAccessType AutoDetect to default SessionOptions

## v2.61

- Updated connecting to EOP and S&C center using EXOPSv2 module
- Removed needless passing of AzureADAuthorizationEndpointUri when specifying UserPrincipalName

## v2.60

- Removed Skype Online Connector support (retired 15 Feb 2021; use MSTeams instead)
- Removed obsolete Connect-ExchangeOnlinev2 helper function
- Connect-ExchangeOnline will use ExchangeOnlineManagement
- Replaced variable-substitution strings with -f formatted versions
- Replaced aliases with full verbs

## v2.58

- Replaced web call to retrieve tenant ID with much quicker REST call

## v2.56

- Added PowerShell 7.x support (rewrite of some module management calls)

## v2.55

- Fixed updating module when it's loaded
- Fixed removal of old modules logic

## v2.51

- Added ConvertTo-SystemVersion helper function to deal with N.N-PreviewN

## v2.5

- Switched to using PowerShellGet 2.x cmdlets (Get-InstalledModule) for performance
- Added mention of PowerShell, PowerShellGet and PackageManagement version in header

## v2.45

- Improved loading speed by collecting Module information once
- Added AllowPrerelease to uninstall-module operation

## v2.44

- Fixed unneeded update of module in Update-Office365Modules

## v2.43

- Added support for MSCommerce

## v2.42

- Fixed bugs in reporting on and updating modules

## v2.41

- Made Elevated check language-independent

## v2.40

- Added code to detect Exchange Online module version
- Added code to update Exchange Online module
- Speedup loading by skipping version checks (use Show-Office365Modules & Update-Office365Modules)
- Only online version checks are performed; removed 'offline' version data

## v2.31

- Added Microsoft.Graph.Teams.Team module

## v2.30

- Added pre-release modules support

## v2.29

- Updated Exchange Online Management v2 (1.0.1)

## v2.24

- Added Show-Office365Modules to report on known vs online versions

## v2.23

- Added PowerShell Graph module (0.1.1)

## v2.20

- Updated Update-Office365Modules detection logic
- Updated Update-Office365Modules to skip non-repo installed modules

## v2.13

- Removed OnlineAutoUpdate option
- Added notice to use Update-Office365Modules
- Fixed updating of binary modules

## v2.12

- Fixed module processing bug
- Added module upgrading with 'AcceptLicense' switch

## v2.10

- Added Update-Office365Modules

## v2.00

- Added Exchange Online Management v2 (0.3374.4)

## v1.99.92 – 1.99.91

- Updated various module versions (SharePoint, AzureAD, Exchange Online)

## v1.99.90

- Added Microsoft.Intune.Graph module

## v1.99.89

- Updated various module versions

## v1.99.88

- Added PowerApps modules (preview)
- Replaced AADRM module functionality with AIPModule

## v1.99.86

- Updated Exchange Online info

## v1.98.85

- Fixed setting Tenant Name for Connect-SharePointOnline

## v1.98.82 – 1.98.83

- Updated module versions (Teams, AzureAD, SharePoint)
- Revised module auto-updating

## v1.98.81

- Updated Exchange Online info

## v1.98.8

- Added changing console title to Tenant info
- Rewrite initializing to make it manageable from profile

## v1.98.5

- Added display of Tenant ID after providing credentials

## v1.98.4 – 1.98.2

- Updated various module versions

## v1.98.1

- Fixed Connect-ComplianceCenter function

## v1.98

- Added SharePointPnP Online (detection only)
- Fixed Azure RMS location and info

## v1.97

- Updated module versions

## v1.96

- Fixed Skype & SharePoint Module version checking

## v1.95

- Fixed version checking issue in Get-Office365Credentials

## v1.94

- Moved all global vars into one global hashtable (myOffice365Services)

## v1.91

- Fixed removal of old module(s) when updating

## v1.90

- Renamed 'Multi-Factor Authentication' to 'Modern Authentication'

## v1.86

- Added automatic module updating (Admin mode, OnlineModuleAutoUpdate & OnlineModuleVersionChecks)

## v1.85

- Fixed menu creation in ISE

## v1.80

- Added Microsoft Teams PowerShell Module support
- Added Connect-MSTeams function
- Cleared default PSSessionOptions
- Some code rewrite (splatting)

## v1.75

- Added support for MFA-enabled Security & Compliance Center
- Added Set-Office365Environment to switch to other region (Germany, China, etc.)

## v1.7

- Added AzureAD PowerShell Module support
- For disambiguation, renamed Connect-AzureAD to Connect-AzureActiveDirectory

## v1.6

- Added support for the Skype for Business PowerShell module w/MFA
- Added support for the SharePoint Online PowerShell module w/MFA

## v1.5

- Added support for Exchange Online PowerShell module w/MFA
- Added IE proxy config support

## v1.4

- Added (in-code) AzureEnvironment (Connect-AzureAD)
- Added version reporting for modules

## v1.3

- Updated required version of Online Sign-In Assistant

## v1.2

- Community release

