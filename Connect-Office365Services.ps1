<#
    .SYNOPSIS
    Connect-Office365Services

    PowerShell script defining functions to connect to Office 365 online services
    or Exchange On-Premises.

    Michel de Rooij
    michel@eightwone.com
    http://eightwone.com

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    Version 3.54, September 14th, 2025

    Get the latest version from GitHub:
    https://github.com/michelderooij/Connect-Office365Services

    KNOWN LIMITATIONS:
    - When specifying PSSessionOptions for Modern Authentication, authentication fails (OAuth).
      Therefore, no PSSessionOptions are used for Modern Authentication.

    .DESCRIPTION
    The functions are listed below. Note that functions may call each other, for example, to
    connect to Exchange Online, the Office 365 Credentials the user is prompted to enter these credentials.
    Also, the credentials are persistent in the current session; there is no need to re-enter credentials
    when connecting to Exchange Online Protection for example. Should different credentials be required,
    call Get-Office365Credentials or Get-OnPremisesCredentials again.

    Helper Functions:
    =================
    - Connect-AzureRMS              Connects to Azure Rights Management
    - Connect-ExchangeOnline        Connects to Exchange Online (Graph module)
    - Connect-AIP                   Connects to Azure Information Protection
    - Connect-PowerApps             Connects to PowerApps
    - Connect-IPPSSession           Connects to Compliance Center
    - Connect-SharePointOnline      Connects to SharePoint Online
    - Connect-MSTeams               Connects to Microsoft Teams
    - Get-Office365Credentials      Gets Office 365 credentials
    - Connect-ExchangeOnPremises    Connects to Exchange On-Premises
    - Get-OnPremisesCredentials     Gets On-Premises credentials
    - Get-ExchangeOnPremisesFQDN    Gets FQDN for Exchange On-Premises
    - Get-Office365Tenant           Gets Office 365 tenant name
    - Set-Office365Environment	    Configures Uri's and region to use
    - Update-Office365Modules       Updates supported Office 365 modules
    - Report-Office365Modules       Report on known vs online module versions
    - Clean-Office365Modules        Cleanup old versions of supported modules
    - Select-Office365Modules       Interactive module (un)installation menu

    Functions to connect to other services provided by the module, e.g. Connect-MSGraph or Connect-MSTeams.

    To load the helper functions from your PowerShell profile, put Connect-Office365Services.ps1 in the same location
    as your $profile file, and edit $profile as follows:
    & (Join-Path $PSScriptRoot "Connect-Office365Services.ps1")

    .HISTORY
    1.2     Community release
    1.3     Updated required version of Online Sign-In Assistant
    1.4     Added (in-code) AzureEnvironment (Connect-AzureAD)
            Added version reporting for modules
    1.5     Added support for Exchange Online PowerShell module w/MFA
            Added IE proxy config support
            Small cosmetic changes in output
    1.51    Fixed PSSession for non-MFA EXO logons
            Changed credential entering logic for MFA
    1.6     Added support for the Skype for Business PowerShell module w/MFA
            Added support for the SharePoint Online PowerShell module w/MFA
    1.61    Fixed MFA choice bug
    1.7     Added AzureAD PowerShell Module support
            For disambiguation, renamed Connect-AzureAD to Connect-AzureActiveDirectory
    1.71    Added AzureADPreview PowerShell Module Support
    1.72    Changed credential non-prompting condition for AzureAD
    1.75    Added support for MFA-enabled Security & Compliance Center
            Added module version checks (online when possible) setting OnlineModuleVersionChecks
            Switched AzureADv1 to PS gallery version
            Removed Sign-In Assistant checks
            Added Set-Office365Environment to switch to other region, e.g. Germany, China etc.
    1.76    Fixed version number checks for SfB & SP
    1.77    Fixed version number checks for determining MFA status
            Changed default for online version checks to $false
    1.78    Added usage of most recently dated ExO MFA module found (in case multiple are found)
            Changed connecting to SCC with MFA to using the underlying New-ExO cmdlet
    1.80    Added Microsoft Teams PowerShell Module support
            Added Connect-MSTeams function
            Cleared default PSSessionOptions
            Some code rewrite (splatting)
    1.81    Added support for ExO module 16.00.2020.000 w/MFA -Credential support
    1.82    Bug fix SharePoint module version check
    1.83    Removed Credentials option for ExO/MFA connect
    1.84    Added Exchange ADAL loading support
    1.85    Fixed menu creation in ISE
    1.86    Updated version check for AzureADPreview (2.0.0.154)
            Added automatic module updating (Admin mode, OnlineModuleAutoUpdate & OnlineModuleVersionChecks)
    1.87    Small bug fixes in outdated logic
            Added showing OnlineChecks/AutoUpdate/IsAdmin info
    1.88    Updated module updating routine
            Updated SkypeOnlineConnector reference (PSGallery)
            Updated versions for Teams
    1.89    Reverted back to installable SkypeOnlineConnector
    1.90    Updated info for Azure Active Directory Preview module
            Updated info for Exchange Online Modern Authentication module
            Renamed 'Multi-Factor Authentication' to 'Modern Authentication'
    1.91    Updated info for SharePoint Online module
            Fixed removal of old module(s) when updating
    1.92    Updated AzureAD module 2.0.1.6
    1.93    Updated Teams module 0.9.3
            Fixed typo in uninstall of old module when upgrading
    1.94    Moved all global vars into one global hashtable (myOffice365Services)
            Updated AzureAD preview info (v2.0.1.11)
            Updated AzureAD info (v2.0.1.10)
    1.95    Fixed version checking issue in Get-Office365Credentials
    1.96    Updated AzureADv1 (MSOnline) info (v1.1.183.8)
            Fixed Skype & SharePoint Module version checking in Get-Office365Credentials()
    1.97    Updated AzureAD Preview info (v2.0.1.17)
    1.98    Updated Exchange Online info (16.0.2440.0)
            Updated AzureAD Preview info (v2.0.1.18)
            Updated AzureAD info (v2.0.1.16)
            Fixed Azure RMS location + info (v2.13.1.0)
            Added SharePointPnP Online (detection only)
    1.98.1  Fixed Connect-ComplianceCenter function
    1.98.2  Updated Exchange Online info (16.0.2433.0 - 2440 seems pulled)
            Added x86 notice (not all modules available for x86 platform)
    1.98.3  Updated Exchange Online info (16.00.2528.000)
            Updated SharePoint Online info (v16.0.8029.0)
            Updated Microsoft Online info (1.1.183.17)
    1.98.4  Updated AzureAD Preview info (2.0.2.3)
            Updated SharePoint Online info (16.0.8119.0)
            Updated Exchange Online info (16.00.2603.000)
            Updated MSOnline info (1.1.183.17)
            Updated AzureAD info (2.2.2.2)
            Updated SharePointPnP Online info (3.1.1809.0)
    1.98.5  Added display of Tenant ID after providing credentials
    1.98.6  Updated Teams info (0.9.5)
            Updated AzureAD Preview info (2.0.2.5)
            Updated SharePointPnP Online info (3.2.1810.0)
    1.98.7  Modified Module Updating routing
    1.98.8  Updated SharePoint Online info (16.0.8212.0)
            Added changing console title to Tenant info
            Rewrite initializing to make it manageable from profile
    1.98.81 Updated Exchange Online info (16.0.2642.0)
    1.98.82 Updated AzureAD info (2.0.2.4)
            Updated MicrosoftTeams info (0.9.6)
            Updated SharePoint Online info (16.0.8525.1200)
            Revised module auto-updating
    1.98.83 Updated Teams info (1.0.0)
            Updated AzureAD v2 Preview info (2.0.2.17)
            Updated SharePoint Online info (16.0.8715.1200)
    1.98.84 Updated Skype for Business Online info (7.0.1994.0)
    1.98.85 Updated SharePoint Online info (16.0.8924.1200)
            Fixed setting Tenant Name for Connect-SharePointOnline
    1.99.86 Updated Exchange Online info (16.0.3054.0)
    1.99.87 Replaced 'not detected' with 'not found' for esthetics
    1.99.88 Replaced AADRM module functionality with AIPModule
            Updated AzureAD v2 info (2.0.2.31)
            Added PowerApps modules (preview)
            Fixed handling when ExoPS module isn't installed
    1.99.89 Updated AzureAD v2 Preview info (2.0.2.32)
            Updated SPO Online info (16.0.9119.1200)
            Updated Teams info (1.0.1)
    1.99.90 Added Microsoft.Intune.Graph module
            Updated AzureAD v2 info (2.0.2.50)
            Updated AzureAD v2 Preview info (2.0.2.51)
            Updated SharePoint Online info (16.0.19223.12000)
            Updated MSTeams info (1.0.2)
    1.99.91 Updated Exchange Online info (16.0.3346.0)
            Updated AzureAD v2 info (2.0.2.52)
            Updated AzureAD v2 Preview info (2.0.2.53)
            Updated SharePoint Online info (16.0.19404.12000)
    1.99.92 Updated SharePoint Online info (16.0.19418.12000)
    2.00    Added Exchange Online Management v2 (0.3374.4)
    2.10    Added Update-Office365Modules
            Updated MSOnline info (1.1.183.57)
            Updated AzureAD v2 info (2.0.2.61)
            Updated AzureAD v2 Preview info (2.0.2.62)
            Updated PowerApps-Admin-PowerShell info (2.0.21)
    2.11    Added MSTeams info from Test Gallery (1.0.18)
            Updated MSTeams info (1.0.3)
            Updated PowerApps-Admin-PowerShell info (2.0.24)
    2.12    Fixed module processing bug
            Added module upgrading with 'AcceptLicense' switch
    2.13    Removed OnlineAutoUpdate option
            Added notice to use Update-Office365Modules
            Fixed updating of binary modules
            Updated ExchangeOnlineManagement v2 info (0.3374.9)
            Splash header cosmetics
    2.14    Fixed bug in Update-Office365Modules
    2.15    Fixed module detection installed side-by-side
    2.20    Updated ExchangeOnlineManagement v2 info (0.3374.10)
            Updated Azure AD v2 info (2.0.2.76)
            Updated Azure AD v2 Preview info (2.0.2.77)
            Updated SharePoiunt Online info (16.0.19515.12000)
            Updated Update-Office365Modules detection logic
            Updated Update-Office365Modules to skip non-repo installed modules
    2.21    Updated ExchangeOnlineManagement v2 info (0.3374.11)
            Updated PowerApps-Admin-PowerShell info (2.0.34)
            Updated SharePoint PnP Online info (3.17.2001.2)
    2.22    Updated ExchangeOnlineManagement v2 info (0.3555.1)
            Updated MSTeams (Test) info (1.0.19)
    2.23    Added PowerShell Graph module (0.1.1)
            Updated Exchange Online info (16.00.3527.000)
            Updated SharePoint Online info (16.0.19724.12000)
    2.24    Updated ExchangeOnlineManagement v2 info (0.3582.0)
            Updated Microsoft Teams (Test) info (1.0.20)
            Added Report-Office365Modules to report on known vs online versions
    2.25    Updated Microsoft Teams info (1.0.5)
            Updated Azure AD v2 Preview info (2.0.2.85)
            Updated SharePoint Online info (16.0.19814.12000)
            Updated MSTeams (Test) info (1.0.21)
            Updated SharePointPnP Online (3.19.2003.0)
            Updated PowerApps-Admin-PowerShell (2.0.45)
            Updated PowerApps-PowerShell (1.0.9)
            Updated Report-Office365Modules (cosmetic, repository checks)
            Improved loading speed a bit (for repository checks)
    2.26    Added setting Window title to include current account
    2.27    Updated ExchangeOnlineManagement to v0.4578.0
            Updated Azure AD v2 Preview info (2.0.2.89)
            Updated Azure Information Protection info (1.0.0.2)
            Updated SharePoint Online info (16.0.20017.12000)
            Updated MSTeams (Test) info (1.0.22)
            Updated SharePointPnP Online info (3.20.2004.0)
            Updated PowerApps-Admin-PowerShell info (2.0.60)
    2.28    Updated Azure AD v2 Preview info (2.0.2.102)
            Updated SharePointPnP Online info (3.21.2005.1)
            Updated PowerApps-Admin-PowerShell info (2.0.63)
    2.29    Updated Exchange Online Management v2 (1.0.1)
            Updated SharePoint Online (16.0.20122.12000)
            Updated SharePointPnP Online (3.21.2005.2)
            Updated PowerApps-Admin-PowerShell (2.0.64)
            Updated PowerApps-PowerShell (1.0.13)
    2.30    Updated Exchange Online Management Pre-release (2.0.3)
            Updated Azure Active Directory (v2) (2.0.2.104)
            Updated SharePoint Online updated to (16.0.20212.12000)
            Updated Microsoft Teams (Test) (1.0.25)
            Updated Microsoft Teams (2.0.7)
            Updated SharePointPnP Online (3.22.2006.2)
            Updated PowerApps-Admin-PowerShell (2.0.66)
            Updated Microsoft.Graph (0.7.0)
            Added pre-release modules support
    2.31    Added Microsoft.Graph.Teams.Team module
            Updated Azure Active Directory (v2 Preview) (2.0.2.105)
            Updated PowerApps-Admin-PowerShell (2.0.67)
    2.32    Updated Exchange Online info (16.0.3724.0)
            Updated Azure AD (v2) (2.0.2.106)
            Updated SharePoint PnP Online (2.0.72)
            Updated Microsoft Teams (GA) (1.1.4)
            Updated SharePoint PnP Online (3.23.2007.1)
            Updated PowerApps-Admin-PowerShell (2.0.72)
    2.40    Added code to detect Exchange Online module version
            Added code to update Exchange Online module
            Speedup loading by skipping version checks (use Report-Office365Modules & Update-Office365Modules)
            Only online version checks are performed (removes 'offline' version data)
            Some visual cosmetics and simplifications
    2.41    Made Elevated check language-independent
    2.42    Fixed bugs in reporting on and updating modules
            Cosmetics when reporting
    2.43    Added support for MSCommerce
    2.44    Fixed unneeded update of module in Update-Office365Modules
            Slightly speed up updating and reporting routine
    2.45    Improved loading speed by collecting Module information once
            Added AllowPrerelease to uninstall-module operation
    2.5     Switched to using PowerShellGet 2.x cmdlets (Get-InstalledModule) for performance
            Added mention of PowerShell, PowerShellGet and PackageManagement version in header
            Removed InternetAccess mention in header
    2.51    Added ConvertTo-SystemVersion helper function to deal with N.N-PreviewN
    2.52    Added NoClobber and AcceptLicense to update
    2.53    Fixed reporting of installed verion during update
    2.54    Improved module updating
    2.55    Fixed updating updating module when it's loaded
            Fixed removal of old modules logic (.100 is newer than .81)
            Set default response of MFA question to Yes
    2.56    Added PowerShell 7.x support (rewrite of some module management calls)
    2.57    Corrected SessionOption to PSSessionOption for Connect-ExchangeOnline (@ladewig)
    2.58    Replaced web call to retrieve tenant ID with much quicker REST call
    2.60    Changes due to Skype Online Connector retirement per 15Feb2021 (use MSTeams instead)
            Changes due to deprecation of ExoPowershellModule (use EXOPSv2 instead)
            Connect-ExchangeOnline will use ExchangeOnlineManagement
            Removed obsolete Connect-ExchangeOnlinev2 helper function
            Replaced variable-substitution strings "$(..)" with -f formatted versions
            Replaced aliases with full verbs. Happy PSScriptAnalyzer :)
            Due to removal of non-repository module checks, significant loading speed reduction.
    2.61    Updated connecting to EOP and S&C center using EXOPSv2 module
            Removed needless passing of AzureADAuthorizationEndpointUri when specifying UserPrincipalName
    2.62    Added -ProxyAccessType AutoDetect to default SessionOptions
    2.63    Changed default ProxyAccessType to None
    2.64    Structured Connect-MsTeams
    2.65    Fixed connecting to AzureAD using MFA not using provided Username
    2.66    Reporting change in #cmdlets after updating
    2.70    Added support for all overloaded Connect-ExchangeOnline parameters from ExchangeOnlineManagement module
            Added PnP.PowerShell module support
            Removed SharePointPnPPowerShellOnline support
            Removed obsolete code for MFA module presence check
            Updated AzureADAuthorizationEndpointUri for Common/GCC
    2.71    Revised module updating using Install-Package when available
    2.80    Improved version handling to properly evaluate Preview modules
            Fixed updating module using install-package when existing package comes from different repo
            Versions reported are now showing their textual representation, including tags like PreviewX
            Report-Office365Modules output is now more condense
    2.90    Added MSCommerce module
            Added MicrosoftPowerBIMgmt module
            Added Az module
    2.91    Removed Microsoft.Graph.Teams.Team module (unlisted at PSGallery)
    2.92    Removed duplicate MSCommerce checking
    2.93    Added cleaning up of module dependencies (e.g. Az)
            Updating will use same scope of installed module
            Showing warning during update when running multiple PowerShell sessions
    2.94    Added AllowClubber to ignore existing cmdlet conflicts when updating modules
    2.95    Added UseRPSSession switch for Connect-ExchangeOnline
    2.96    Added Microsoft36DSC module
            Fixed determing current module scope (CurrentUser/AllUsers)
    2.97    Fixed title for admin roles
    2.98    Fixed ConnectionUri in EXO connection method
    2.99    Added 2 connect helper functions to description
    3.00    Fixed wrongly detecting old modules because mixed native PS module and PSGet cmdlets
            Back to using native PS module management cmdlets
            Some cosmetics
            Startup only reports installed modules, not "not installed"
            Report now also reports not installed modules
            Removed PSGet check
    3.01    Added Preview info when reporting local module info
    3.10    Removed Microsoft Teams (Test) support (from poshtestgallery)
            Renamed Azure AD v1 to MSOnline to prevent confusion
            Added support for WhiteboardAdmin
            Added support for MSIdentityTools
    3.11    Fixed header not displaying correction script version
    3.12    Replaced 'Prerelease' questions with switch - specify if you want, otherwise default is unspecified (=GA)
    3.13    Added ORCA to set of supported modules
    3.14    Added O365CentralizedAddInDeployment to set of supported modules
    3.15    Fixed creating ISE menu options for local functions
            Removed Connect-EOP
    3.16    Fixed duplicate module processing as connect ComplianceCenter/EXO is in same module
    3.17    Added Microsoft.Graph.Compatibility.AzureAD (Preview)
    3.18    Added Microsoft.Graph.Beta
    3.19    Removed SkypeOnlineConnector & ExoPowerShellModule related code
    3.20    Added Clean-Office365Modules
    3.21    Added Places module
            Added Microsoft.Graph.Entra module
            Added Microsoft.Graph.Entra.Beta module
    3.22    Removed MFA/Non-MFA code
            Removed Connect-AzureAD helper function
            Modified Get-TenantId to use OpenId endpoint to read ID using credentials' username when available
            Removed ISE menu creation code
    3.23    Updated Clean-Office365Modules to process dependencies (eg Graph)
            Removed Compatibility Adapter for AzureAD PowerShell (predecessor Entra PowerShell)
    3.231   Made dependency checking silent when nothing found
    3.3     Changed Microsoft.Graph.Entra to Microsoft.Entra
            Changed Microsoft.Graph.Entra.Beta to Microsoft.Entra.Beta
            Added notice for module replacement, eg microsoft.graph.entra > microsoft.entra
            Module information now stored in JSON for maintainability
            Removed old ISE entries from module information
    3.4     Added using Microsoft.PowerShell.PSResourceGet when available (performance)
            Removed obsolete repository code
            Code cleanup
            Cosmetic changes in output
    3.41    Fixed parameter usage issue with not using PSResourceGet
    3.42    Added error handling to Uninstall-MyModule output error handling
    3.43    Fixed Connect-ExchangeOnline
    3.44    Minor cosmetic changes
            Added Quote of the Day like message
    3.45    Fixed Connect-IPPSSession
            Corrected Connect-ComplianceCenter references, changed to Connect-IPPSSession
            Some cosmetic changes
            Removed redundant module check/import pairs
    3.46    Changed module check before import to catch issues
            Small cleanup Connect-SPO
            Small textual corrections in synopsis
    3.50    Added Select-Office365Modules
            Report-Office365Modules only reports on installed modules
            Changed Install-MyModule to accommodate Select-Office365Modules
    3.51    Fixed version argument issue when removing modules
    3.52    Cleanup-Office365Modules will not consider AllUsers & CurrentUser
    3.53    Removed dependency check when installing modules so it will install dependencies (eg Graph.*)
            Bumped required PowerShell version to 5.1
    3.54    Fixed uninstalling dependencies when uninstalling deselected modules
#>

#Requires -Version 5.1

$local:ScriptVersion = '3.54'

Function global:Get-myPSResourceGetInstalled {
    If( $global:myOffice365Services['PSResourceGet']) {
        # Already determined
    }
    Else {
        $global:myOffice365Services['PSResourceGet']= $null -ne (Get-Module -Name Microsoft.PowerShell.PSResourceGet -ListAvailable -ErrorAction SilentlyContinue)
    }
}

Function global:Get-myModule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false, ValueFromPipeline=$true)]
        [string[]]$Name,
        [switch]$ListAvailable,
        [switch]$AllowPrerelease,
        [switch]$AllScopes
    )
    Process {
        If( $AllScopes) {
            If( $global:myOffice365Services['PSResourceGet']) {
                Get-PSResource -Name $Name -Scope AllUsers,CurrentUser -ErrorAction SilentlyContinue
            }
            Else {
                Get-Module -Name $Name -ListAvailable:$ListAvailable -All -Refresh -ErrorAction SilentlyContinue | Sort Path -Unique
            }
        }
        Else {
            If( $global:myOffice365Services['PSResourceGet']) {
                Get-PSResource -Name $Name -Scope $global:myOffice365Services['Scope'] -ErrorAction SilentlyContinue
            }
            Else {
                Get-Module -Name $Name -ListAvailable:$ListAvailable -Refresh -ErrorAction SilentlyContinue
            }
        }
    }
}

Function global:Find-myModule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false, ValueFromPipeline=$true)]
        [string[]]$Name
    )
    Process {
        If( $global:myOffice365Services['PSResourceGet']) {
            Find-PSResource -Name $Name -Prerelease:$global:myOffice365Services['AllowPrerelease'] -ErrorAction SilentlyContinue
        }
        Else {
            Find-Module -Name $Name -AllowPrerelease:$global:myOffice365Services['AllowPrerelease'] -ErrorAction SilentlyContinue
        }
    }
}

Function global:Update-myModule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string[]]$Name
    )
    Process {
        # Unload module if loaded
        Remove-Module -Name $Name -Force -ErrorAction SilentlyContinue

        If( $global:myOffice365Services['PSResourceGet']) {
            Update-PSResource -Name $Name -Scope $global:myOffice365Services['Scope'] -Force -AcceptLicense:$true -Prerelease:$global:myOffice365Services['AllowPrerelease']
        }
        Else {
            Update-Module -Name $Name -Scope $global:myOffice365Services['Scope'] -Force -AllowClobber -AcceptLicense:True -AllowPrerelease:$global:myOffice365Services['AllowPrerelease']
        }
    }
}

Function global:Uninstall-myModule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string]$Name,
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string]$Version='All',
        [switch]$IsPrerelease
        )
    Process {

        # Unload module if loaded
        Remove-Module -Name $Name -Force -ErrorAction SilentlyContinue

        Try {
            If( $global:myOffice365Services['PSResourceGet']) {
                If( $Version -eq 'All') {
                    Uninstall-PSResource -Name $Name -Scope $global:myOffice365Services['Scope'] -SkipDependencyCheck -Prerelease:$IsPrerelease
                }
                Else {
                    Uninstall-PSResource -Name $Name -Version $Version -Scope $global:myOffice365Services['Scope'] -Prerelease:$IsPrerelease
                }
            }
            Else {
                If( $Version -eq 'All') {
                    Uninstall-Module -Name $Name -AllVersions  -Scope $global:myOffice365Services['Scope'] -AllowPrerelease:$IsPrerelease -Force:$AllVersions
                }
                Else {
                    Uninstall-Module -Name $Name -RequiredVersion [string]$Version -Scope $global:myOffice365Services['Scope'] -AllowPrerelease:$IsPrerelease -Force:$AllVersions
                }
            }
        }
        Catch {
            Switch -Regex ($PSItem.FullyQualifiedErrorId) {
                '^AdminPrivilegesRequiredForUninstall,' {
                    Write-Warning ('Unable to uninstall module without Administrator rights: {0} v{1}' -f $Name, $Version)
                }

                '^ModuleIsInUse,' {
                    # Module throws own error
                }

                '^(UnableToUninstallAsOtherModulesNeedThisModule|UninstallPSResourcePackageIsaDependency),' {
                    Write-Warning ('Unable to uninstall module {0} v{1} due to dependencies' -f $Name, $Version)
                }

                Default {
                    Write-Warning ('Problem uninstalling module {0} v{1}: {2}' -f $Name, $Version, $Error[0].Exception.Message)
                }
            }
        }
    }
}

Function global:Install-myModule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string[]]$Name,
        [switch]$AllowPrerelease,
        [switch]$AllowClobber
    )
    Process {
        If( $global:myOffice365Services['PSResourceGet']) {
            Install-PSResource -Name $Name -Prerelease:$AllowPrerelease -Scope $global:myOffice365Services['Scope'] -NoClobber:(-not $AllowClobber) -Confirm:$false -TrustRepository:$true -AcceptLicense:$true -Reinstall:$true
        }
        Else {
            Install-Module -Name $Name -Force -AllowClobber:$AllowClobber -AllowPrerelease:$AllowPrerelease -Scope $global:myOffice365Services['Scope']  -Confirm:$false  -SkipPublisherCheck:$true -AcceptLicense:$true
        }
    }
}

function global:Set-WindowTitle {
    If( $host.ui.RawUI.WindowTitle -and $global:myOffice365Services['TenantID']) {
        $local:PromptPrefix= ''
        $ThisPrincipal= new-object System.Security.principal.windowsprincipal( [System.Security.Principal.WindowsIdentity]::GetCurrent())
        if( $ThisPrincipal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator)) {
	    $local:PromptPrefix= 'Administrator:'
        }
        $local:Title= '{0}{1} connected to Tenant ID {2}' -f $local:PromptPrefix, $global:myOffice365Services['Office365Credentials'].UserName, $global:myOffice365Services['TenantID']
        $host.ui.RawUI.WindowTitle = $local:Title
    }
}

function global:Get-TenantIDfromMail {
    param(
        [string]$mail
    )
    $domainPart= ($mail -split '@')[1]
    If( $domainPart) {
        $res= (Invoke-RestMethod -Uri ('https://login.microsoftonline.com/{0}/v2.0/.well-known/openid-configuration' -f $domainPart)).jwks_uri.split('/')[3]
        If(!( $res)) {
            Write-Warning 'Could not determine Tenant ID using e-mail address'
            $res= $null
        }
    }
    Else {
        Write-Warning 'E-mail address invalid, cannot determine Tenant ID'
        $res= $null
    }
    return $res
}

function global:Get-TenantID {
    $global:myOffice365Services['TenantID']= Get-TenantIDfromMail $global:myOffice365Services['Office365Credentials'].UserName
    If( $global:myOffice365Services['TenantID']) {
        Write-Host ('TenantID: {0}' -f $global:myOffice365Services['TenantID'])
    }
}

function global:Get-Office365ModuleInfo {
'[
    {
        "Module": "ExchangeOnlineManagement",
        "Description": "Exchange Online Management",
        "Repo": "https://www.powershellgallery.com/packages/ExchangeOnlineManagement"
    },
    {
        "Module": "MSOnline",
        "Description": "MSOnline",
        "Repo": "https://www.powershellgallery.com/packages/MSOnline",
        "ReplacedBy": "Microsoft.Entra"
    },
    {
        "Module": "AzureAD",
        "Description": "Azure AD (v2)",
        "Repo": "https://www.powershellgallery.com/packages/azuread",
        "ReplacedBy": "Microsoft.Entra"
    },
    {
        "Module": "AzureADPreview",
        "Description": "Azure AD (v2 Preview)",
        "Repo": "https://www.powershellgallery.com/packages/AzureADPreview",
        "ReplacedBy": "Microsoft.Entra"
    },
    {
        "Module": "AIPService",
        "Description": "Azure Information Protection",
        "Repo": "https://www.powershellgallery.com/packages/AIPService"
    },
    {
        "Module": "Microsoft.Online.Sharepoint.PowerShell",
        "Description": "SharePoint Online",
        "Repo": "https://www.powershellgallery.com/packages/Microsoft.Online.SharePoint.PowerShell"
    },
    {
        "Module": "MicrosoftTeams",
        "Description": "Microsoft Teams",
        "Repo": "https://www.powershellgallery.com/packages/MicrosoftTeams"
    },
    {
        "Module": "MSCommerce",
        "Description": "Microsoft Commerce",
        "Repo": "https://www.powershellgallery.com/packages/MSCommerce"
    },
    {
        "Module": "PnP.PowerShell",
        "Description": "PnP.PowerShell",
        "Repo": "https://www.powershellgallery.com/packages/PnP.PowerShell"
    },
    {
        "Module": "Microsoft.PowerApps.Administration.PowerShell",
        "Description": "PowerApps-Admin-PowerShell",
        "Repo": "https://www.powershellgallery.com/packages/Microsoft.PowerApps.Administration.PowerShell"
    },
    {
        "Module": "Microsoft.PowerApps.PowerShell",
        "Description": "PowerApps-PowerShell",
        "Repo": "https://www.powershellgallery.com/packages/Microsoft.PowerApps.PowerShell"
    },
    {
        "Module": "Microsoft.Graph.Intune",
        "Description": "MSGraph-Intune",
        "Repo": "https://www.powershellgallery.com/packages/Microsoft.Graph.Intune"
    },
    {
        "Module": "Microsoft.Graph",
        "Description": "Microsoft.Graph",
        "Repo": "https://www.powershellgallery.com/packages/Microsoft.Graph"
    },
    {
        "Module": "Microsoft.Graph.Beta",
        "Description": "Microsoft.Graph.Beta",
        "Repo": "https://www.powershellgallery.com/packages/Microsoft.Graph.Beta"
    },
    {
        "Module": "Microsoft.Graph.Entra",
        "Description": "Microsoft.Graph.Entra",
        "Repo": "https://www.powershellgallery.com/packages/Microsoft.Graph.Entra",
        "ReplacedBy": "Microsoft.Entra"
    },
    {
        "Module": "Microsoft.Entra",
        "Description": "Microsoft.Entra",
        "Repo": "https://www.powershellgallery.com/packages/Microsoft.Entra",
        "Replaces": "Microsoft.Graph.Entra"
    },
    {
        "Module": "Microsoft.Graph.Entra.Beta",
        "Description": "Microsoft.Graph.Entra.Beta",
        "Repo": "https://www.powershellgallery.com/packages/Microsoft.Graph.Entra.Beta",
        "ReplacedBy": "Microsoft.Entra.Beta"
    },
    {
        "Module": "Microsoft.Entra.Beta",
        "Description": "Microsoft.Entra.Beta",
        "Repo": "https://www.powershellgallery.com/packages/Microsoft.Entra.Beta",
        "Replaces": "Microsoft.Graph.Entra.Beta"
    },
    {
        "Module": "MicrosoftPlaces",
        "Description": "MicrosoftPlaces",
        "Repo": "https://www.powershellgallery.com/packages/MicrosoftPlaces"
    },
    {
        "Module": "MicrosoftPowerBIMgmt",
        "Description": "MicrosoftPowerBIMgmt",
        "Repo": "https://www.powershellgallery.com/packages/MicrosoftPowerBIMgmt"
    },
    {
        "Module": "Az",
        "Description": "Az",
        "Repo": "https://www.powershellgallery.com/packages/Az"
    },
    {
        "Module": "Microsoft365DSC",
        "Description": "Microsoft365DSC",
        "Repo": "https://www.powershellgallery.com/packages/Microsoft36DSC"
    },
    {
        "Module": "WhiteboardAdmin",
        "Description": "WhiteboardAdmin",
        "Repo": "https://www.powershellgallery.com/packages/WhiteboardAdmin"
    },
    {
        "Module": "MSIdentityTools",
        "Description": "MSIdentityTools",
        "Repo": "https://www.powershellgallery.com/packages/MSIdentityTools"
    },
    {
        "Module": "O365CentralizedAddInDeployment",
        "Description": "O365 Centralized Add-In Deployment Module",
        "Repo": "https://www.powershellgallery.com/packages/O365CentralizedAddInDeployment"
    },
    {
        "Module": "ORCA",
        "Description": "Office 365 Recommended Configuration Analyzer (ORCA)",
        "Repo": "https://www.powershellgallery.com/packages/ORCA"
    }
    ]' | ConvertFrom-Json
}

function global:Select-Office365Modules {
    param(
        [switch]$AllowPrerelease
    )

    $local:IsAdmin= [System.Security.principal.windowsprincipal]::new([System.Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)

    if (-not $local:IsAdmin) {
        Write-Warning 'Script not running with elevated privileges; AllUsers scoped module management might fail'
    }
    Write-Host ''

    $local:ModuleInfo = Get-Office365ModuleInfo
    $local:CurrentSelection = @{}
    $local:SelectedIndex = 0

    # Initialize current selection based on installed modules
    foreach ($module in $local:ModuleInfo) {
        $installedModule = Get-myModule -Name $module.Module -ListAvailable |
            Where-Object { ([System.Uri]($_.RepositorySourceLocation)).Authority -ieq ([System.Uri]($module.Repo)).Authority } |
            Select-Object -First 1
        $local:CurrentSelection[$module.Module] = $null -ne $installedModule
    }

    function Show-ModuleMenu {
        param($ModuleInfo, $CurrentSelection, $SelectedIndex, $maxColumns, $ColumnSpacing)

        $screenWidth = $Host.UI.RawUI.WindowSize.Width
        $columnWidth = [Math]::Floor(($screenWidth - $columnSpacing) / $maxColumns)

        $maxRows = [Math]::Ceiling($ModuleInfo.Count / $maxColumns)

        for ($row = 0; $row -lt $maxRows; $row++) {
            for ($col = 0; $col -lt $maxColumns; $col++) {
                $i = $row * $maxColumns + $col

                if ($i -lt $ModuleInfo.Count) {
                    $module = $ModuleInfo[$i]
                    $isSelected = $CurrentSelection[$module.Module]
                    $checkbox = if ($isSelected) { "[*]" } else { "[ ]" }
                    $prefix = if ($i -eq $SelectedIndex) { ">" } else { " " }

                    $line = "{0} {1} {2}" -f $prefix, $checkbox, $module.Description

                    if ($module.ReplacedBy) {
                        $line += " [Replaced by $($module.ReplacedBy)]"
                    }

                    if ($line.Length -gt ($columnWidth - 2)) {
                        $line = $line.Substring(0, $columnWidth - 4) + ".."
                    }

                    $line = $line.PadRight($columnWidth)

                    if ($i -eq $SelectedIndex) {
                        Write-Host $line -ForegroundColor $Host.PrivateData.VerboseForegroundColor -BackgroundColor $Host.PrivateData.VerboseBackgroundColor -NoNewline
                    } else {
                        if ($isSelected) {
                            Write-Host $line -ForegroundColor $Host.PrivateData.FormatAccentColor -NoNewline
                        } else {
                            Write-Host $line -ForegroundColor $Host.UI.RawUI.ForegroundColor -NoNewline
                        }
                    }
                } else {
                    # Empty space for column alignment
                    Write-Host (" " * $columnWidth) -NoNewline
                }

                # Add spacing between columns (except for the last column)
                if ($col -lt ($maxColumns - 1)) {
                    Write-Host (" " * $columnSpacing) -NoNewline
                }
            }
            Write-Host ''
        }
        Write-Host ''
        Write-Host "Current Scope: $($global:myOffice365Services['Scope'])"
        Write-Host '[Up/Down/Left/Right] Navigate, [Space] Toggle, [S] Scope, [Enter] Commit, [Esc] Cancel'
    }

    # Helper function to determine which column and row an index is in
    function Get-ColumnPosition {
        param($Index, $ModuleCount, $MaxColumns)
        if ($Index -ge $ModuleCount -or $Index -lt 0) {
            return @{ Column = 0; Row = 0 }
        }

        # Calculate position in 2-column row-major layout
        $row = [Math]::Floor($Index / $MaxColumns)
        $column = $Index % $MaxColumns
        return @{ Column = $column; Row = $row }
    }

    # Helper function to get index from column and row
    function Get-IndexFromPosition {
        param($Column, $Row, $ModuleCount, $MaxColumns)

        # Ensure column is within bounds (0 or 1 for 2 columns)
        $Column = [Math]::Max(0, [Math]::Min($Column, $MaxColumns - 1))
        $Row = [Math]::Max(0, $Row)

        # Calculate the index based on row-major layout
        $index = $Row * $MaxColumns + $Column

        # Ensure the index doesn't exceed the module count
        if ($index -ge $ModuleCount) {
            # If we're beyond the last module, stay at the last valid index
            $index = $ModuleCount - 1
        }

        return [Math]::Max(0, [Math]::Min($index, $ModuleCount - 1))
    }

    $exitMenu = $false
    $committed = $false

    $maxColumns = 2
    $columnSpacing= 2
    $maxRows= [Math]::Ceiling($ModuleInfo.Count / $maxColumns)

    # Fill estate where menu gets be displayed, to avoid bottom of console calculation challenges
    1.. ($maxRows + 3) | ForEach-Object { Write-Host ''}
    $menuTopY= [Console]::get_CursorTop() - $maxRows - 3

    while (-not $exitMenu) {

        [Console]::SetCursorPosition( [Console]::get_CursorLeft(), $menuTopY)
        Show-ModuleMenu -ModuleInfo $local:ModuleInfo -CurrentSelection $local:CurrentSelection -SelectedIndex $local:SelectedIndex -maxColumns $maxColumns -ColumnSpacing $columnSpacing

        $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

        switch ($key.VirtualKeyCode) {
            38 { # Up arrow
                $currentPos = Get-ColumnPosition -Index $local:SelectedIndex -ModuleCount $local:ModuleInfo.Count -MaxColumns $maxColumns
                $newRow = [Math]::Max(0, $currentPos.Row - 1)
                $newIndex = Get-IndexFromPosition -Column $currentPos.Column -Row $newRow -ModuleCount $local:ModuleInfo.Count -MaxColumns $maxColumns
                $local:SelectedIndex = $newIndex
            }
            40 { # Down arrow
                $currentPos = Get-ColumnPosition -Index $local:SelectedIndex -ModuleCount $local:ModuleInfo.Count -MaxColumns $maxColumns
                $newRow = $currentPos.Row + 1
                $newIndex = Get-IndexFromPosition -Column $currentPos.Column -Row $newRow -ModuleCount $local:ModuleInfo.Count -MaxColumns $maxColumns
                # Only move if the new index is valid and different
                if ($newIndex -lt $local:ModuleInfo.Count -and $newIndex -ne $local:SelectedIndex) {
                    $local:SelectedIndex = $newIndex
                }
            }
            37 { # Left arrow
                $currentPos = Get-ColumnPosition -Index $local:SelectedIndex -ModuleCount $local:ModuleInfo.Count -MaxColumns $maxColumns
                $newColumn = [Math]::Max(0, $currentPos.Column - 1)
                $newIndex = Get-IndexFromPosition -Column $newColumn -Row $currentPos.Row -ModuleCount $local:ModuleInfo.Count -MaxColumns $maxColumns
                if ($newIndex -lt $local:ModuleInfo.Count) {
                    $local:SelectedIndex = $newIndex
                }
            }
            39 { # Right arrow
                $currentPos = Get-ColumnPosition -Index $local:SelectedIndex -ModuleCount $local:ModuleInfo.Count -MaxColumns $maxColumns
                $newColumn = [Math]::Min($maxColumns - 1, $currentPos.Column + 1)
                $newIndex = Get-IndexFromPosition -Column $newColumn -Row $currentPos.Row -ModuleCount $local:ModuleInfo.Count -MaxColumns $maxColumns
                if ($newIndex -lt $local:ModuleInfo.Count) {
                    $local:SelectedIndex = $newIndex
                }
            }
            32 { # Spacebar
                if ($local:SelectedIndex -ge 0 -and $local:SelectedIndex -lt $local:ModuleInfo.Count) {
                    $currentModule = $local:ModuleInfo[$local:SelectedIndex].Module
                    $local:CurrentSelection[$currentModule] = -not $local:CurrentSelection[$currentModule]
                }
            }
            83 { # S key - Toggle scope
                if ($global:myOffice365Services['Scope'] -eq 'AllUsers') {
                    $global:myOffice365Services['Scope'] = 'CurrentUser'
                } else {
                    $global:myOffice365Services['Scope'] = 'AllUsers'
                }
            }
            13 { # Enter
                $exitMenu = $true
                $committed = $true
            }
            27 { # Escape
                $exitMenu = $true
                $committed = $false
            }
        }
    }

    if (-not $committed) {
        # Operation cancelled
        Return
    }

    $modulesToInstall = [System.Collections.ArrayList]@()
    $modulesToUninstall = [System.Collections.ArrayList]@()

    foreach ($module in $local:ModuleInfo) {
        $moduleName = $module.Module
        $shouldBeInstalled = $local:CurrentSelection[$moduleName]

        $installedModule = Get-myModule -Name $moduleName -ListAvailable | Where-Object { ([System.Uri]($_.RepositorySourceLocation)).Authority -ieq ([System.Uri]($module.Repo)).Authority } | Select-Object -First 1
        $isCurrentlyInstalled = $null -ne $installedModule

        if ($shouldBeInstalled -and -not $isCurrentlyInstalled) {
            $modulesToInstall.Add( $module) | Out-Null
        }
        else {
            if (-not $shouldBeInstalled -and $isCurrentlyInstalled) {
                $modulesToUninstall.Add( $module) | Out-Null
            }
        }
    }

    # Install selected modules
    foreach ($module in $modulesToInstall) {
        Write-Host ('Installing {0}' -f $module.Description)
        try {
            Install-myModule -Name $module.Module -AllowPrerelease:$AllowPrerelease -AllowClobber
            $allVersions = Get-myModule -Name $module.Module -ListAvailable | Where-Object { ([System.Uri]($_.RepositorySourceLocation)).Authority -ieq ([System.Uri]($module.Repo)).Authority } | Select-Object -First 1
            if ($allVersions) {
                Write-Host ('Installed {0} v{1}' -f $allVersions.Name, $allVersions.Version) -ForegroundColor Green
            }
        }
        catch {
            Write-Error ('Failed to install {0}: {1}' -f $module.Name, $_.Exception.Message)
        }
    }

    # Uninstall deselected modules
    foreach ($module in $modulesToUninstall) {
        $requiredModules= (Get-myModule -Name $module.module -ListAvailable | Select-Object -First 1).Dependencies | Sort-Object -Unique Name
        If( $requiredModules) {
            ForEach( $reqModule in $requiredModules) {
                try {
                    Write-Host ('Uninstalling dependency {0}' -f $reqmodule.Name) -ForegroundColor White
                    Uninstall-myModule -Name $reqmodule.Name -Version 'All' -IsPrerelease:$reqmodule.IsPrerelease
                }
                catch {
                    Write-Error ('Failed to uninstall {0}: {1}' -f $reqmodule.Name, $_.Exception.Message)
                }
            }
        }
return
        try {
            Write-Host ('Uninstalling {0}' -f $module.module) -ForegroundColor White
            Uninstall-myModule -Name $module.module -Version 'All' -IsPrerelease:$module.IsPrerelease
        }
        catch {
            Write-Error ('Failed to uninstall {0}: {1}' -f $module.Name, $_.Exception.Message)
        }
    }

    if ($modulesToInstall.Count -eq 0 -and $modulesToUninstall.Count -eq 0) {
        Write-Host "No changes were made."
    }
}

function global:Set-Office365Environment {
    param(
        [ValidateSet('Germany', 'China', 'AzurePPE', 'USGovernment', 'Default')]
        [string]$Environment
    )
    Switch ( $Environment) {
        'Germany' {
            $global:myOffice365Services['ConnectionEndpointUri'] = 'https://outlook.office.de/PowerShell-LiveID'
            $global:myOffice365Services['SCCConnectionEndpointUri'] = 'https://ps.compliance.protection.outlook.de/PowerShell-LiveId'
            $global:myOffice365Services['EOPConnectionEndpointUri'] = 'https://ps.protection.protection.outlook.de/PowerShell-LiveId'
            $global:myOffice365Services['AzureADAuthorizationEndpointUri'] = 'https://login.microsoftonline.de/common'
            $global:myOffice365Services['SharePointRegion'] = 'Germany'
            $global:myOffice365Services['AzureEnvironment'] = 'AzureGermanyCloud'
            $global:myOffice365Services['TeamsEnvironment'] = ''
        }
        'China' {
            $global:myOffice365Services['ConnectionEndpointUri'] = 'https://partner.outlook.cn/PowerShell-LiveID'
            $global:myOffice365Services['SCCConnectionEndpointUri'] = 'https://ps.compliance.protection.outlook.com/PowerShell-LiveId'
            $global:myOffice365Services['EOPConnectionEndpointUri'] = 'https://ps.protection.protection.outlook.com/PowerShell-LiveId'
            $global:myOffice365Services['AzureADAuthorizationEndpointUri'] = 'https://login.chinacloudapi.cn/common'
            $global:myOffice365Services['SharePointRegion'] = 'China'
            $global:myOffice365Services['AzureEnvironment'] = 'AzureChinaCloud'
            $global:myOffice365Services['TeamsEnvironment'] = ''
        }
        'AzurePPE' {
            $global:myOffice365Services['ConnectionEndpointUri'] = ''
            $global:myOffice365Services['SCCConnectionEndpointUri'] = 'https://ps.compliance.protection.outlook.com/PowerShell-LiveId'
            $global:myOffice365Services['EOPConnectionEndpointUri'] = 'https://ps.protection.protection.outlook.com/PowerShell-LiveId'
            $global:myOffice365Services['AzureADAuthorizationEndpointUri'] = ''
            $global:myOffice365Services['SharePointRegion'] = ''
            $global:myOffice365Services['AzureEnvironment'] = 'AzurePPE'
        }
        'USGovernment' {
            $global:myOffice365Services['ConnectionEndpointUri'] = 'https://outlook.office365.com/PowerShell-LiveId'
            $global:myOffice365Services['SCCConnectionEndpointUri'] = 'https://ps.compliance.protection.outlook.com/PowerShell-LiveId'
            $global:myOffice365Services['EOPConnectionEndpointUri'] = 'https://ps.protection.protection.outlook.com/PowerShell-LiveId'
            $global:myOffice365Services['AzureADAuthorizationEndpointUri'] = 'https://login.microsoftonline.com/common'
            $global:myOffice365Services['SharePointRegion'] = 'ITAR'
            $global:myOffice365Services['AzureEnvironment'] = 'AzureUSGovernment'
        }
        default {
            $global:myOffice365Services['ConnectionEndpointUri'] = 'https://outlook.office365.com/PowerShell-LiveId'
            $global:myOffice365Services['SCCConnectionEndpointUri'] = 'https://ps.compliance.protection.outlook.com/PowerShell-LiveId'
            $global:myOffice365Services['EOPConnectionEndpointUri'] = 'https://ps.protection.protection.outlook.com/PowerShell-LiveId'
            $global:myOffice365Services['AzureADAuthorizationEndpointUri'] = 'https://login.microsoftonline.com/common'
            $global:myOffice365Services['SharePointRegion'] = 'Default'
            $global:myOffice365Services['AzureEnvironment'] = 'AzureCloud'
        }
    }
}

function global:Connect-ExchangeOnline {
    [CmdletBinding()]
    param(
        [string]$ConnectionUri,
        [string]$AzureADAuthorizationEndpointUri,
        [string]$ExchangeEnvironmentName,
        [System.Management.Automation.Remoting.PSSessionOption]$PSSessionOption= $null,
        [switch]$BypassMailboxAnchoring= $false,
        [string]$DelegatedOrganization,
        [string]$Prefix,
        [switch]$ShowBanner,
        [string]$UserPrincipalName,
        [System.Management.Automation.PSCredential]$Credential,
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
        [string]$CertificateFilePath,
        [System.Security.SecureString]$CertificatePassword,
        [string]$CertificateThumbprint,
        [string]$AppId,
        [string]$Organization,
        [string]$AccessToken = '',
        [switch]$ManagedIdentity,
        [string]$ManagedIdentityAccountId,
        [switch]$EnableErrorReporting,
        [string]$LogDirectoryPath,
        $LogLevel,
        [bool]$TrackPerformance,
        [bool]$ShowProgress= $true,
        [bool]$UseMultithreading,
        [uint32]$PageSize,
        [switch]$Device,
        [switch]$InlineCredential,
        [string[]]$CommandName = @("*"),
        [string[]]$FormatTypeName = @("*"),
        [switch] $SkipLoadingFormatData = $false,
        [switch] $SkipLoadingCmdletHelp = $true,
        [switch] $LoadCmdletHelp = $false,
        [System.Security.Cryptography.X509Certificates.X509Certificate2] $SigningCertificate = $null,
        [switch] $DisableWAM = $false,
        [switch] $UseRPSSession = $false
    )
    if (!( $PSBoundParameters.ContainsKey('ConnectionUri'))) {
        $PSBoundParameters['ConnectionUri']= $global:myOffice365Services['ConnectionEndpointUri']
    }
    if (!( $PSBoundParameters.ContainsKey('AzureADAuthorizationEndpointUri'))) {
        $PSBoundParameters['AzureADAuthorizationEndpointUri']= $global:myOffice365Services['AzureADAuthorizationEndpointUri']
    }
    if (!( $PSBoundParameters.ContainsKey('PSSessionOption'))) {
        $PSBoundParameters['PSSessionOption']= $global:myOffice365Services['SessionExchangeOptions']
    }
    If ( $PSBoundParameters.ContainsKey('UserPrincipalName') -or $PSBoundParameters.ContainsKey('Certificate') -or
        $PSBoundParameters.ContainsKey('CertificateFilePath') -or $PSBoundParameters.ContainsKey('CertificateThumbprint') -or
        $PSBoundParameters.ContainsKey('AppId')) {
        Write-Host ('Connecting to Exchange Online ..')
    }
    Else {
        If ( $PSBoundParameters.ContainsKey('Credential')) {
            Write-Host ('Connecting to Exchange Online using {0} ..' -f ($Credential).UserName)
            $global:myOffice365Services['Office365Credentials']= $Credential
        }
        Else {
            If ( $global:myOffice365Services['Office365Credentials']) {
                Write-Host ('Connecting to Exchange Online using {0} ..' -f $global:myOffice365Services['Office365Credentials'].UserName)
                $PSBoundParameters['Credential']= $global:myOffice365Services['Office365Credentials']
            }
            Else {
                Get-Office365Credentials
                If ( $global:myOffice365Services['Office365Credentials']) {
                    Write-Host ('Connecting to Exchange Online using {0} ..' -f $global:myOffice365Services['Office365Credentials'].UserName)
                    $PSBoundParameters['Credential']= $global:myOffice365Services['Office365Credentials']
                }
                Else {
                    Write-Host ('Connecting to Exchange Online ..')
                }
            }
        }
    }

    If(!( Get-Module -Name ExchangeOnlineManagement -ErrorAction SilentlyContinue)) {
        Import-Module -Name ExchangeOnlineManagement -ErrorAction SilentlyContinue
    }
    If( Get-Command -Name Connect-ExchangeOnline -ErrorAction SilentlyContinue) {
        $global:myOffice365Services['Session365'] = ExchangeOnlineManagement\Connect-ExchangeOnline @PSBoundParameters
        If ( $global:myOffice365Services['Session365'] ) {
            Import-PSSession -Session $global:myOffice365Services['Session365'] -AllowClobber
        }
    }
    Else {
        Write-Error -Message 'Cannot connect to Exchange Online - module not installed or not loading.'
    }
}

function global:Connect-ExchangeOnPremises {
    If ( !($global:myOffice365Services['OnPremisesCredentials'])) { Get-OnPremisesCredentials }
    If ( !($global:myOffice365Services['ExchangeOnPremisesFQDN'])) { Get-ExchangeOnPremisesFQDN }
    If ( !($global:myOffice365Services['OnPremisesCredentials'])) {
        Write-Host ('Connecting to Exchange On-Premises {0} using {1} ..' -f $global:myOffice365Services['ExchangeOnPremisesFQDN'], $global:myOffice365Services['OnPremisesCredentials'].UserName)
        $global:myOffice365Services['SessionExchange'] = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$($global:myOffice365Services['ExchangeOnPremisesFQDN'])/PowerShell" -Credential $global:myOffice365Services['OnPremisesCredentials'] -Authentication Kerberos -AllowRedirection -SessionOption $global:myOffice365Services['SessionExchangeOptions']
        If ( $global:myOffice365Services['SessionExchange']) {Import-PSSession -Session $global:myOffice365Services['SessionExchange'] -AllowClobber}
    }
}

Function global:Get-ExchangeOnPremisesFQDN {
    $global:myOffice365Services['ExchangeOnPremisesFQDN'] = Read-Host -Prompt 'Enter Exchange On-Premises endpoint, e.g. exchange1.contoso.com'
}

function global:Connect-IPPSSession {
    If(!( Get-Module -Name ExchangeOnlineManagement -ErrorAction SilentlyContinue)) {
        Import-Module -Name ExchangeOnlineManagement -ErrorAction SilentlyContinue
    }
    If( Get-Command -Name Connect-ExchangeOnline -ErrorAction SilentlyContinue) {
        Write-Host ('Connecting to Security & Compliance Center ..')
        $global:myOffice365Services['SessionCC'] = ExchangeOnlineManagement\Connect-IPPSSession -ConnectionUri $global:myOffice365Services['SCCConnectionEndpointUri'] -UserPrincipalName ($global:myOffice365Services['Office365Credentials']).UserName -PSSessionOption $global:myOffice365Services['SessionExchangeOptions']
        If ( $global:myOffice365Services['SessionCC'] ) {
            Import-PSSession -Session $global:myOffice365Services['SessionCC'] -AllowClobber
        }
    }
    Else {
        Write-Error -Message 'Cannot connect to Security & Compliance Center - module not installed or not loading.'
    }
}


function global:Connect-MSTeams {
    If(!( Get-Module -Name MicrosoftTeams -ListAvailable)) {
        Import-Module -Name MicrosoftTeams -ErrorAction SilentlyContinue
    }
    If ( Get-Command -Name Connect-MicrosoftTeams -ErrorAction SilentlyContinue) {
        If ( !($global:myOffice365Services['Office365Credentials'])) { Get-Office365Credentials }
        Write-Host ('Connecting to Microsoft Teams using {0} ..' -f $global:myOffice365Services['Office365Credentials'].UserName)
        Connect-MicrosoftTeams -AccountId ($global:myOffice365Services['Office365Credentials']).UserName -TenantId $global:myOffice365Services['TenantId']
    }
    Else {
        Write-Error -Message 'Cannot connect to Microsoft Teams - module not installed or not loading.'
    }
}

function global:Connect-AIP {
    If(!( Get-Module -Name AIPService -ListAvailable)) {
        Import-Module -Name AIPService -ErrorAction SilentlyContinue
    }
    If ( Get-Command -Name Connect-AipService -ErrorAction SilentlyContinue) {
        If ( !($global:myOffice365Services['Office365Credentials'])) { Get-Office365Credentials }
        Write-Host ('Connecting to Azure Information Protection using {0}' -f $global:myOffice365Services['Office365Credentials'].UserName)
        Connect-AipService -Credential $global:myOffice365Services['Office365Credentials']
    }
    Else {
        Write-Error -Message 'Cannot connect to Azure Information Protection - module not installed or not loading.'
    }
}
function global:Connect-SharePointOnline {
    If(!( Get-Module -Name Microsoft.Online.SharePoint.PowerShell -ListAvailable)) {
        Import-Module -Name Microsoft.Online.SharePoint.PowerShell -ErrorAction SilentlyContinue
    }
    If ( Get-Command -Name Connect-SPOService -ErrorAction SilentlyContinue) {
        If ( !($global:myOffice365Services['Office365Credentials'])) { Get-Office365Credentials }
        If (($global:myOffice365Services['Office365Credentials']).UserName -like '*.onmicrosoft.com') {
            $global:myOffice365Services['Office365Tenant'] = ($global:myOffice365Services['Office365Credentials']).UserName.Substring(($global:myOffice365Services['Office365Credentials']).UserName.IndexOf('@') + 1).Replace('.onmicrosoft.com', '')
        }
        Else {
            If ( !($global:myOffice365Services['Office365Tenant'])) { Get-Office365Tenant }
        }
        Write-Host 'Connecting to SharePoint Online  ..'
        $Parms = @{
            url= 'https://{0}-admin.sharepoint.com' -f $($global:myOffice365Services['Office365Tenant'])
            region= $global:myOffice365Services['SharePointRegion']
        }
        Connect-SPOService @Parms
    }
    Else {
        Write-Error -Message 'Cannot connect to SharePoint Online - module not installed or not loading.'
    }
}
function global:Connect-PowerApps {
    If(!( Get-Module -Name Microsoft.PowerApps.PowerShell -ListAvailable)) {
        Import-Module -Name Microsoft.PowerApps.PowerShell -ErrorAction SilentlyContinue
    }
    If(!( Get-Module -Name Microsoft.PowerApps.Administration.PowerShell -ListAvailable)) {
        Import-Module -Name Microsoft.PowerApps.Administration.PowerShell -ErrorAction SilentlyContinue
    }
    If ( Get-Command -Name Add-PowerAppsAccount -ErrorAction SilentlyContinue) {
        If ( !($global:myOffice365Services['Office365Credentials'])) { Get-Office365Credentials }
        Write-Host "Connecting to PowerApps using $($global:myOffice365Services['Office365Credentials'].UserName) .."
        $Parms = @{'Username' = $global:myOffice365Services['Office365Credentials'].UserName }
        Add-PowerAppsAccount @Parms
    }
    Else {
        Write-Error -Message 'Cannot connect to SharePoint Online - problem loading module.'
    }
}

Function global:Get-Office365Credentials {
    $global:myOffice365Services['Office365Credentials'] = $host.ui.PromptForCredential('Office 365 Credentials', 'Please enter your Office 365 credentials', $global:myOffice365Services['Office365Credentials'].UserName, '')
    Get-TenantID
    Set-WindowTitle
}

Function global:Get-OnPremisesCredentials {
    $global:myOffice365Services['OnPremisesCredentials'] = $host.ui.PromptForCredential('On-Premises Credentials', 'Please Enter Your On-Premises Credentials', '', '')
}

Function global:Get-Office365Tenant {
    If( $global:myOffice365Services['Office365Credentials']) {
        $local:OpenIdInfo= Invoke-RestMethod ('https://login.windows.net/{0}/.well-known/openid-configuration' -f ($global:myOffice365Services['Office365Credentials'].UserName.Split('@')[1])) -Method GET
        $global:myOffice365Services['Office365Tenant']= $local:OpenIdInfo.userinfo_endpoint.Split('/')[3]
    }
    Else {
        $global:myOffice365Services['Office365Tenant'] = Read-Host -Prompt 'Enter tenant ID, e.g. contoso for contoso.onmicrosoft.com'
    }
}

Function global:Get-ModuleScope {
    param(
        $Module
    )
    If( $Module.ModuleBase -ilike ('{0}*' -f (Join-Path -Path $ENV:HOMEDRIVE -ChildPath $ENV:HOMEPATH))) {
        'CurrentUser'
    }
    Else {
        'AllUsers'
    }
}

function global:Get-ModuleVersionInfo {
    param(
        $Module
    )
    $Module= $Module | Select-Object -First 1
    $ModuleManifestPath = $Module.Path
    If( $ModuleManifestPath) {
        $isModuleManifestPathValid = Test-Path -Path $ModuleManifestPath
        If(!( $isModuleManifestPathValid)) {
            # Module manifest path invalid, skipping extracting prerelease info
            $ModuleVersion= $Module.Version.ToString()
        }
        Else {
            $ModuleManifestContent = Get-Content -Path $ModuleManifestPath
            $preReleaseInfo = $ModuleManifestContent -match "Prerelease = '(.*)'"
            If( $preReleaseInfo) {
                $preReleaseVersion= $preReleaseInfo[0].Split('=')[1].Trim().Trim("'")
                If( $preReleaseVersion) {
                    $ModuleVersion= ('{0}-{1}' -f $Module.Version.ToString(), $preReleaseVersion)
                }
                Else {
                    $ModuleVersion= $Module.Version.ToString()
                }
            }
            Else {
                $ModuleVersion= $Module.Version.ToString()
            }
        }
    }
    Else {
        $ModuleVersion= $Module.Version.ToString()
    }
    $ModuleVersion
}

Function global:Update-Office365Modules {
    param (
        [switch]$AllowPrerelease
    )

    $local:Functions= Get-Office365ModuleInfo
    $global:myOffice365Services['AllowPrerelease']= $AllowPrerelease

    $local:IsAdmin= [System.Security.principal.windowsprincipal]::new([System.Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
    If( $local:IsAdmin) {
        If( (Get-Process -Name powershell, pwsh -ErrorAction SilentlyContinue | Measure-Object).Count -gt 1) {
            Write-Warning ('Running multiple PowerShell sessions, successful updating might be problematic.')
        }
        ForEach ( $local:Item in $local:Functions) {

            $local:Module= Get-myModule -Name ('{0}' -f $local:Item.Module) | Sort-Object -Property Version -Descending | Select-Object -First 1

            If( ($local:Module).RepositorySourceLocation) {

                $local:Version = Get-ModuleVersionInfo -Module $local:Module
                Write-Host ('Checking {0}' -f $local:Item.Description) -NoNewLine

                $local:NewerAvailable= $false
                $OnlineModule = Find-myModule -Name $local:Item.Module -ErrorAction SilentlyContinue
                If( $OnlineModule) {
                    Write-Host (': Local:{0}, Online:{1}' -f $local:Version, $OnlineModule.version)
                    If( (Compare-TextVersionNumber -Version $local:Version -CompareTo $OnlineModule.version) -eq 1) {
                        $local:NewerAvailable= $true
                    }
                }
                Else {
                        # Not installed from online or cannot determine
                        Write-Host ('Local:{0} Online:N/A' -f $local:Version)
                }

                If( $local:NewerAvailable) {
                    $local:UpdateSuccess= $false
                    Try {
                        Update-myModule -Name $local:Item.Module
                        $local:UpdateSuccess= $true
                    }
                    Catch {
                        Write-Error ('Problem updating {0}:{1}' -f $local:Item.Module, $Error[0].Exception.Message)
                    }

                    If( $local:UpdateSuccess) {

                        $local:ModuleVersions= Get-myModule -Name $local:Item.Module -ListAvailable 

                        $local:Module = $local:ModuleVersions | Sort-Object -Property @{e={ [System.Version]($_.Version -replace '[^\d\.]','')}} -Descending | Select-Object -First 1
                        $local:LatestVersion = ($local:Module).Version
                        Write-Host ('Updated {0} to version {1}' -f $local:Item.Description, $local:LatestVersion) -ForegroundColor Green

                        # Uninstall all old versions of module & dependencies
                        If( $OnlineModule) {
                            ForEach( $DependencyModule in $Module.Dependencies) {

                                $local:DepModuleVersions= Get-myModule -Name $DependencyModule.Name -ListAvailable
                                $local:DepModule = $local:DepModuleVersions | Sort-Object -Property @{e={ [System.Version]($_.Version -replace '[^\d\.]','')}} -Descending | Select-Object -First 1
                                $local:DepLatestVersion = ($local:DepModule).Version
                                $local:OldDepModules= $local:DepModuleVersions | Where-Object {$_.Version -ne $local:DepLatestVersion}
                                $local:OldDepModules | ForEach-Object {
                                    $DepModule= $_
                                    Write-Host ('Uninstalling dependency {0} version {1}' -f $DepModule.Name, $DepModule.Version)
                                    Try {
                                        Uninstall-myModule -Name $DepModule.Name -Version $DepModule.Version -IsPrerelease:$DepModule.IsPrerelease
                                    }
                                    Catch {
                                        Write-Error ('Problem uninstalling {0} version {1}' -f $DepModule.Name, $DepModule.Version)
                                    }
                                }
                            }
                            $local:OldModules= $local:ModuleVersions | Where-Object {$_.Version -ne $local:LatestVersion}
                            If( $local:OldModules) {
                                ForEach( $OldModule in $local:OldModules) {
                                    Write-Host ('Uninstalling {0} version {1}' -f $local:Item.Description, $OldModule.Version) -ForegroundColor White
                                    Try {
                                        Uninstall-myModule -Name $OldModule.Name -Version $OldModule.Version -IsPrerelease:$OldModule.IsPrerelease
                                    }
                                    Catch {
                                        Write-Error ('Problem uninstalling {0} version {1}' -f $OldModule.Name, $OldModule.Version)
                                    }
                                }
                            }
                        }
                    }
                    Else {
                        # Problem during update
                    }
                }
                Else {
                    # No update available
                }
            }
            Else {
                # Not installed
            }
        }
    }
    Else {
        Write-Warning ('Script not running with elevated privileges; module management might fail')
    }
}

Function global:Clean-Office365Modules {
    param (
        [switch]$AllowPrerelease
    )

    $local:Functions= Get-Office365ModuleInfo
    $global:myOffice365Services['AllowPrerelease']= $AllowPrerelease

    $local:IsAdmin= [System.Security.principal.windowsprincipal]::new([System.Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
    If( $local:IsAdmin) {
        If( (Get-Process -Name powershell, pwsh -ErrorAction SilentlyContinue | Measure-Object).Count -gt 1) {
            Write-Warning ('Running multiple PowerShell sessions, successful cleanup might be problematic.')
        }
        ForEach ( $local:Item in $local:Functions) {

            $local:Module= Get-Module -Name ('{0}' -f $local:Item.Module) -ListAvailable | Sort-Object -Property Version -Descending

            If( $local:Module) {
                Write-Host ('Checking {0} .. ' -f $local:Item.Description) -NoNewline

                $local:ModuleVersions= Get-myModule -Name $local:Item.Module -ListAvailable -AllScopes -ErrorAction SilentlyContinue
                $local:LatestModule = $local:ModuleVersions | Sort-Object -Property @{e={ [System.Version]($_.Version -replace '[^\d\.]','')}} -Descending | Select-Object -First 1
                $local:LatestVersion = ($local:LatestModule).Version

                $local:OldModules= $local:ModuleVersions | Where-Object {$_.Version -ne $local:LatestVersion}
                If( $local:OldModules) {

                    Write-Host ('Previous versions found') -ForegroundColor Yellow

                    ForEach( $OldModule in $local:OldModules) {

                        # Uninstall all old versions of the module
                        Write-Host ('Uninstalling {0} v{1}' -f $OldModule.Name, $OldModule.Version) -ForegroundColor White
                        Try {
                            Uninstall-myModule -Name $OldModule.Name -Version $OldModule.Version -IsPrerelease:$OldModule.IsPrerelease
                        }
                        Catch {
                            Write-Error ('Problem uninstalling {0} v{1}: {2}' -f $OldModule.Name, $OldModule.Version, $Error[0].Exception.Message)
                        }
                    }
                }
                Else {
                    Write-Host ('OK') -ForegroundColor Green
                }

                # Cleanup required modules as well
                $local:RequiredModules= $local:Module.RequiredModules | Sort-Object -Unique Name

                ForEach( $RequiredModule in $local:RequiredModules) {

                    Write-Host ('Checking {0} .. ' -f $RequiredModule.Name) -NoNewline

                    $local:ModuleVersions= Get-myModule -Name $RequiredModule.Name -ListAvailable -AllScopes -ErrorAction SilentlyContinue
                    $local:LatestModule = $local:ModuleVersions | Sort-Object -Property @{e={ [System.Version]($_.Version -replace '[^\d\.]','')}} -Descending | Select-Object -First 1
                    $local:LatestVersion = ($local:LatestModule).Version

                    $local:OldModules= $local:ModuleVersions | Where-Object {$_.Version -ne $local:LatestVersion}
                    If( $local:OldModules) {

                        Write-Host ('needs cleanup')

                        ForEach( $OldModule in $local:OldModules) {

                            Write-Host ('Uninstalling {0} v{1}' -f $OldModule.Name, $OldModule.Version)
                            Try {
                                Uninstall-myModule -Name $OldModule.Name -Version $OldModule.Version -IsPrerelease:$OldModule.IsPrerelease
                            }
                            Catch {
                                Write-Error ('Problem uninstalling {0} v{1}: {2}' -f $OldModule.Name, $OldModule.Version, $Error[0].Exception.Message)
                            }
                        }
                    }
                    Else {
                        Write-Host ('OK') -ForegroundColor Green
                    }
                }
            }
        }
    }
    Else {
        Write-Host ('Script not running with elevated privileges; cannot remove modules') -ForegroundColor Yellow
    }
}

# Compare-TextVersionNumber to handle (rich) version comparison, similar to [System.Version]'s CompareTo method
# 1=CompareTo is newer, 0 = Equal, -1 = Version is Newer
Function global:Compare-TextVersionNumber {
    param(
        [string]$Version,
        [string]$CompareTo
    )
    $res= 0
    $null= $Version -match '^(?<version>[\d\.]+)(\-)?([a-zA-Z]*(?<preview>[\d]*))?$'
    $VersionVer= [System.Version]($matches.Version)
    If( $matches.Preview) {
        # Suffix .0 to satisfy SystemVersion as '#' won't initialize
        $VersionPreviewVer= [System.Version]('{0}.0' -f $matches.Preview)
    }
    Else {
        $VersionPreviewVer= [System.Version]'99999.99999'
    }
    $null= $CompareTo -match '^(?<version>[\d\.]+)(\-)?([a-zA-Z]*(?<preview>[\d]*))?$'
    $CompareToVer= [System.Version]($matches.Version)
    If( $matches.Preview) {
        $CompareToPreviewVer= [System.Version]('{0}.0' -f $matches.Preview)
    }
    Else {
        $CompareToPreviewVer= [System.Version]'99999.99999'
    }

    If( $VersionVer -gt $CompareToVer) {
        $res= -1
    }
    Else {
        If( $VersionVer -lt $CompareToVer) {
            $res= 1
        }
        Else {
            # Equal - Check Preview Tag
            If( $VersionPreviewVer -gt $CompareToPreviewVer) {
                $res= -1
            }
            Else {
                If( $VersionPreviewVer -lt $CompareToPreviewVer) {
                    $res= 1
                }
                Else {
                    # Really Equal
                    $res= 0
                }
            }

        }
    }
    $res
}

Function global:Report-Office365Modules {
    param(
        [switch]$AllowPrerelease
    )

    $local:Functions= Get-Office365ModuleInfo
    $global:myOffice365Services['AllowPrerelease']= $AllowPrerelease

    ForEach ( $local:Item in $local:Functions) {

        $local:Module= Get-myModule -Name ('{0}' -f $local:Item.Module) -ListAvailable | Sort-Object -Property Version -Descending
        $local:Module= $local:Module | Where-Object {([System.Uri]($_.RepositorySourceLocation)).Authority -ieq ([System.Uri]($local:Item.Repo)).Authority } | Select-Object -First 1

        If( $local:Module) {

            $local:Version = Get-ModuleVersionInfo -Module $local:Module
            Write-Host ('{0}: Local v{1}' -f $local:Item.Description, $Local:Version) -NoNewline -ForegroundColor Gray
            $OnlineModule = Find-myModule -Name $local:Item.Module -ErrorAction SilentlyContinue

            If( $OnlineModule) {
                Write-Host (', Online v{0}' -f $OnlineModule.version) -NoNewline -ForegroundColor Gray
            }
            Else {
                Write-Host (', Online N/A') -NoNewline -ForegroundColor Gray
            }
            Write-Host (', Scope:{0} Status:' -f (Get-ModuleScope -Module $local:Module)) -NoNewline -ForegroundColor Gray

            If( [string]::IsNullOrEmpty( $local:Version) -or [string]::IsNullOrEmpty( $OnlineModule.version)) {
                Write-Host ('Unknown') -ForegroundColor Yellow
            }
            Else {
                If( (Compare-TextVersionNumber -Version $local:Version -CompareTo $OnlineModule.version) -eq 1) {
                    Write-Host ('Outdated') -ForegroundColor Red
                }
                Else {
                    Write-Host ('OK') -ForegroundColor Green
                }
            }
            If( $local:Item.ReplacedBy) {
                Write-Warning ('{0} has been replaced by {1}' -f $local:Item.Module, $local:Item.ReplacedBy) -ForegroundColor Yellow
            }
        }
        Else {
            # Module not installed
        }
    }
}

function global:Connect-Office365 {
    Connect-AzureActiveDirectory
    Connect-AzureRMS
    Connect-ExchangeOnline
    Connect-MSTeams
    Connect-IPPSSession
    Connect-SharePointOnline
}

$PSGetModule= Get-Module -Name Microsoft.PowerShell.PSResourceGet -ListAvailable -ErrorAction SilentlyContinue | Sort-Object -Property Version -Descending | Select-Object -First 1
If(! $PSGetModule) {
    Write-Warning ('Microsoft.PowerShell.PSResourceGet module not found; install it for faster package management')
    $PSGetVer= 'N/A'
}
Else {
    $PSGetVer= $PSGetModule.Version
}
$PackageManagementModule= Get-Module -Name PackageManagement -ListAvailable -ErrorAction SilentlyContinue | Sort-Object -Property Version -Descending | Select-Object -First 1
If(! $PackageManagementModule) {
    $PMMVer= 'N/A'
}
Else {
    $PMMVer= $PackageManagementModule.Version
}

Write-Host ('*' * 78)
Write-Host ('Connect-Office365Services v{0}' -f $local:ScriptVersion)
Write-Host ('Source: https://github.com/michelderooij/Connect-Office365Services')

# See if the Administator built-in role is part of your role
$local:IsAdmin= [System.Security.principal.windowsprincipal]::new([System.Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)

# Initialize global state variable when needed
If( -not( Get-Variable myOffice365Services -ErrorAction SilentlyContinue )) { $global:myOffice365Services=@{} }

# Local Exchange session options
$global:myOffice365Services['SessionExchangeOptions'] = New-PSSessionOption -ProxyAccessType None

# Default install scope
$global:myOffice365Services['Scope'] = 'AllUsers'

# Initialize environment & endpoints
Set-Office365Environment -AzureEnvironment 'Default'

Write-Host ('Environment:{0}, Administrator:{1}, Scope:{2}' -f $global:myOffice365Services['AzureEnvironment'], $local:IsAdmin, $global:myOffice365Services['Scope'])
Write-Host ('PS:{0}, PSResourceGet:{1}, PackageManagement:{2}' -f ($PSVersionTable).PSVersion, $PSGetVer, $PMMVer )
Write-Host ('*' * 78)

$local:Functions= Get-Office365ModuleInfo

# Determine if we can use PSResourceGet or need to use Module cmdlets
Get-myPSResourceGetInstalled

$local:Functions | ForEach-Object -Process {

    $local:Item = $_

    $local:Module= Get-MyModule -Name ('{0}' -f $local:Item.Module) -ListAvailable | Sort-Object -Property Version -Descending
    $local:ModuleMatch= ([System.Uri]($local:Module | Select-Object -First 1).RepositorySourceLocation).Authority -eq ([System.Uri]$local:Item.Repo).Authority
    If( $local:ModuleMatch) {
        $local:Module = $local:Module | Sort-Object -Property @{e= { [System.Version]($_.Version -replace '[^\d\.]','')}} -Descending
        $local:Module= $local:Module | Where-Object {([System.Uri]($_.RepositorySourceLocation)).Authority -ieq ([System.Uri]($local:Item.Repo)).Authority } | Select-Object -First 1
        $local:Version = Get-ModuleVersionInfo -Module $local:Module
        Write-Host ('Found {0} (v{1})' -f $local:Item.Description, $local:Version)
        If( $local:Item.ReplacedBy) {
            Write-Warning ('{0} replaced by {1}' -f $local:Item.Module, $local:Item.ReplacedBy)
        }
    }
    Else {
        # Module not found
    }
}

#Get random text
$local:Quotes= @(
    "You are standing in an open field west of a white house, with a boarded front door. There is a small mailbox here.",
    "You wake up. The room is spinning very gently round your head.`nOr at least it would be if you could see it which you can't. It is pitch black.",
    "You are in a comfortable tunnel like hall. To the east there is the round green door.",
    "You are standing at the end of a road before a small brick building. Around you is a forest.`nA small stream flows out of the building and down a gully.",
    "Shall we play a game?",
    "Request access to CLU program.",
    "You are in a clearing, with a forest surrounding you on all sides. A path leads north."
)
Write-Host ( '{0}{1}' -f [System.Environment]::NewLine, ($local:Quotes | Get-Random))

