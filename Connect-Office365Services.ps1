<#
    .SYNOPSIS
    Connect-Office365Services

    PowerShell script defining functions to connect to Office 365 online services
    or Exchange On-Premises. Call manually or alternatively embed or call from $profile
    (Shell or ISE) to make functions available in your session. If loaded from
    PowerShell_ISE, menu items are defined for the functions. To surpress creation of
    menu items, hold 'Shift' while Powershell ISE loads.

    Michel de Rooij
    michel@eightwone.com
    http://eightwone.com

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    Version 3.19, January 24th, 2024

    Get latest version from GitHub:
    https://github.com/michelderooij/Connect-Office365Services

    KNOWN LIMITATIONS:
    - When specifying PSSessionOptions for Modern Authentication, authentication fails (OAuth).
      Therefor, no PSSessionOptions are used for Modern Authentication.
           
    .DESCRIPTION
    The functions are listed below. Note that functions may call eachother, for example to
    connect to Exchange Online the Office 365 Credentials the user is prompted to enter these credentials.
    Also, the credentials are persistent in the current session, there is no need to re-enter credentials
    when connecting to Exchange Online Protection for example. Should different credentials be required,
    call Get-Office365Credentials or Get-OnPremisesCredentials again. 

    Helper Functions:
    =================
    - Connect-AzureActiveDirectory  Connects to Azure Active Directory
    - Connect-AzureRMS              Connects to Azure Rights Management
    - Connect-ExchangeOnline        Connects to Exchange Online (Graph module)
    - Connect-SkypeOnline           Connects to Skype for Business Online
    - Connect-AIP                   Connects to Azure Information Protection
    - Connect-PowerApps             Connects to PowerApps
    - Connect-ComplianceCenter      Connects to Compliance Center
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

    Functions to connect to other services provided by the module, e.g. Connect-MSGraph or Connect-MSTeams.

    To register the PowerShell Test Gallery and install modules from there, use:
    Register-PSRepository -Name PSGalleryInt -SourceLocation https://www.poshtestgallery.com/ -InstallationPolicy Trusted
    Install-Module -Name MicrosoftTeams -Repository PSGalleryInt -Force -Scope AllUsers

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


#>

#Requires -Version 3.0
$local:ScriptVersion= '3.18'

function global:Set-WindowTitle {
    If( $host.ui.RawUI.WindowTitle -and $global:myOffice365Services['TenantID']) {
        $local:PromptPrefix= ''
        $ThisPrincipal= new-object System.Security.principal.windowsprincipal( [System.Security.Principal.WindowsIdentity]::GetCurrent())
        if( $ThisPrincipal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator)) { 
	    $local:PromptPrefix= 'Administrator:'
        }
        $local:Title= '{0}{1} connected to Tenant ID {2}' -f $local:PromptPrefix, $myOffice365Services['Office365Credentials'].UserName, $global:myOffice365Services['TenantID']
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
    $global:myOffice365Services['TenantID']= Get-TenantIDfromMail $myOffice365Services['Office365Credentials'].UserName
    If( $global:myOffice365Services['TenantID']) {
        Write-Host ('TenantID: {0}' -f $global:myOffice365Services['TenantID'])
    }
}

function global:Get-Office365ModuleInfo {
    # Menu | Submenu | Menu ScriptBlock | ModuleName | Description | (Repo)Link 
    @(
        'Connect|Exchange Online|Connect-ExchangeOnline|ExchangeOnlineManagement|Exchange Online Management|https://www.powershellgallery.com/packages/ExchangeOnlineManagement',
        'Connect|Exchange Security & Compliance Center|Connect-ComplianceCenter|ExchangeOnlineManagement|Exchange Online Management|https://www.powershellgallery.com/packages/ExchangeOnlineManagement',
        'Connect|MSOnline|Connect-MSOnline|MSOnline|MSOnline|https://www.powershellgallery.com/packages/MSOnline',
        'Connect|Azure AD (v2)|Connect-AzureAD|AzureAD|Azure Active Directory (v2)|https://www.powershellgallery.com/packages/azuread',
        'Connect|Azure AD (v2 Preview)|Connect-AzureAD|AzureADPreview|Azure Active Directory (v2 Preview)|https://www.powershellgallery.com/packages/AzureADPreview',
        'Connect|Azure AD (Adapter)|Connect-MgGraph|Microsoft.Graph.Compatibility.AzureAD|Compatibility Adapter for AzureAD PowerShell (Preview)|https://www.powershellgallery.com/packages/Microsoft.Graph.Compatibility.AzureAD',
        'Connect|Azure Information Protection|Connect-AIP|AIPService|Azure Information Protection|https://www.powershellgallery.com/packages/AIPService',
        'Connect|SharePoint Online|Connect-SharePointOnline|Microsoft.Online.Sharepoint.PowerShell|SharePoint Online|https://www.powershellgallery.com/packages/Microsoft.Online.SharePoint.PowerShell',
        'Connect|Microsoft Teams|Connect-MSTeams|MicrosoftTeams|Microsoft Teams|https://www.powershellgallery.com/packages/MicrosoftTeams',
        'Connect|Microsoft Commerce|Connect-MSCommerce|MSCommerce|Microsoft Commerce|https://www.powershellgallery.com/packages/MSCommerce',
        'Connect|PnP.PowerShell|Connect-PnPOnline|PnP.PowerShell|PnP.PowerShell|https://www.powershellgallery.com/packages/PnP.PowerShell',
        'Connect|PowerApps-Admin-PowerShell|Connect-PowerApps|Microsoft.PowerApps.Administration.PowerShell|PowerApps-Admin-PowerShell|https://www.powershellgallery.com/packages/Microsoft.PowerApps.Administration.PowerShell',
        'Connect|PowerApps-PowerShell|Connect-PowerApps|Microsoft.PowerApps.PowerShell|PowerApps-PowerShell|https://www.powershellgallery.com/packages/Microsoft.PowerApps.PowerShell',
        'Connect|MSGraph-Intune|Connect-MSGraph|Microsoft.Graph.Intune|MSGraph-Intune|https://www.powershellgallery.com/packages/Microsoft.Graph.Intune',
        'Connect|Microsoft.Graph|Connect-MSGraph|Microsoft.Graph|Microsoft.Graph|https://www.powershellgallery.com/packages/Microsoft.Graph',
        'Connect|Microsoft.Graph.Beta|Connect-MSGraph|Microsoft.Graph.Beta|Microsoft.Graph.Beta|https://www.powershellgallery.com/packages/Microsoft.Graph.Beta',
        'Connect|MicrosoftPowerBIMgmt|Connect-PowerBIServiceAccount|MicrosoftPowerBIMgmt|MicrosoftPowerBIMgmt|https://www.powershellgallery.com/packages/MicrosoftPowerBIMgmt',
        'Connect|Az|Connect-AzAccount|Az|Az|https://www.powershellgallery.com/packages/Az',
        'Connect|Microsoft365DSC|New-M365DSCConnection|Microsoft365DSC|Microsoft365DSC|https://www.powershellgallery.com/packages/Microsoft36DSC',
        'Connect|Whiteboard|Get-Whiteboard|WhiteboardAdmin|WhiteboardAdmin|https://www.powershellgallery.com/packages/WhiteboardAdmin',
        'Connect|Microsoft Identity|Connect-MgGraph|MSIdentityTools|MSIdentityTools|https://www.powershellgallery.com/packages/MSIdentityTools',
        'Connect|Centralized Add-In Deployment|Connect-OrganizationAddInService|O365CentralizedAddInDeployment|O365 Centralized Add-In Deployment Module|https://www.powershellgallery.com/packages/O365CentralizedAddInDeployment',
        'Report|ORCA|Get-ORCAReport|ORCA|Office 365 Recommended Configuration Analyzer (ORCA)|https://www.powershellgallery.com/packages/ORCA',
        'Settings|Office 365 Credentials|Get-Office365Credentials',
        'Connect|Exchange On-Premises|Connect-ExchangeOnPremises',
        'Settings|On-Premises Credentials|Get-OnPremisesCredentials',
        'Settings|Exchange On-Premises FQDN|Get-ExchangeOnPremisesFQDN'
    )
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

function global:Get-MultiFactorAuthenticationUsage {
    $Answer = Read-host  -Prompt 'Would you like to use Modern Authentication? (Y/n) '
    Switch ($Answer.ToUpper()) {
        'N' { $rval = $false }
        Default { $rval = $true}
    }
    return $rval
}

function global:Get-ExchangeOnlineClickOnceVersion {
    Try {
        $ManifestURI= 'https://cmdletpswmodule.blob.core.windows.net/exopsmodule/Microsoft.Online.CSE.PSModule.Client.application'
        $res= Invoke-WebRequest -Uri $ManifestURI -UseBasicParsing
        $xml= [xml]($res.rawContent.substring( $res.rawContent.indexOf('<?xml')))
	    $xml.assembly.assemblyIdentity.version
    }
    Catch {
        Write-Error 'Cannot access or determine version of Microsoft.Online.CSE.PSModule.Client.application'
    }
}

function global:Connect-ExchangeOnline {
    [CmdletBinding()]
    Param(
        [string]$ConnectionUri,
        [string]$AzureADAuthorizationEndpointUri,
        [System.Management.Automation.Remoting.PSSessionOption]$PSSessionOption,
        [switch]$BypassMailboxAnchoring= $false,
        [string]$DelegatedOrganization,
        [string]$Prefix,
        [switch]$ShowBanner= $False,
        [string]$UserPrincipalName,
        [System.Management.Automation.PSCredential]$Credential,
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
        [string]$CertificateFilePath,
        [System.Security.SecureString]$CertificatePassword,
        [string]$CertificateThumbprint,
        [string]$AppId,
        [string]$Organization,
        [switch]$EnableErrorReporting,
        [string]$LogDirectoryPath,
        $LogLevel,
        [bool]$TrackPerformance,
        [bool]$ShowProgress= $True,
        [bool]$UseMultithreading,
        [uint32]$PageSize,
        [switch]$Device,
        [switch]$InlineCredential,
        [string[]]$CommandName = @("*"),
        [string[]]$FormatTypeName = @("*"),
        [switch]$UseRPSSession = $false
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
    If ( $PSBoundParameters.ContainsKey('UserPrincipalName') -or $PSBoundParameters.ContainsKey('Certificate') -or $PSBoundParameters.ContainsKey('CertificateFilePath') -or $PSBoundParameters.ContainsKey('CertificateThumbprint') -or $PSBoundParameters.ContainsKey('AppId')) {
        $global:myOffice365Services['Office365CredentialsMFA']= $True
        Write-Host ('Connecting to Exchange Online with specified Modern Authentication method ..')
    }
    Else {
        If ( $PSBoundParameters.ContainsKey('Credential')) {
            If ( !($global:myOffice365Services['Office365Credentials'])) { Get-Office365Credentials }
            If ( $global:myOffice365Services['Office365CredentialsMFA']) {
                Write-Host ('Connecting to Exchange Online with {0} using Modern Authentication ..' -f $global:myOffice365Services['Office365Credentials'].UserName)
                $PSBoundParameters['UserPrincipalName']= ($global:myOffice365Services['Office365Credentials']).UserName
            }
            Else {
                Write-Host ('Connecting to Exchange Online with {0} ..' -f $global:myOffice365Services['Office365Credentials'].username)
                $PSBoundParameters['Credential']= $global:myOffice365Services['Office365Credentials'] 
            }
        }
        Else {
            Write-Host ('Connecting to Exchange Online with {0} using Legacy Authentication..' -f $PSBoundParameters['Credential'].UserName)
            $global:myOffice365Services['Office365CredentialsMFA']= $False
            $global:myOffice365Services['Office365Credentials']= $PSBoundParameters['Credential']
        }
    }
    $global:myOffice365Services['Session365'] = ExchangeOnlineManagement\Connect-ExchangeOnline @PSBoundParameters
    If ( $global:myOffice365Services['Session365'] ) {
        Import-PSSession -Session $global:myOffice365Services['Session365'] -AllowClobber
    }
}

function global:Connect-ExchangeOnPremises {
    If ( !($global:myOffice365Services['OnPremisesCredentials'])) { Get-OnPremisesCredentials }
    If ( !($global:myOffice365Services['ExchangeOnPremisesFQDN'])) { Get-ExchangeOnPremisesFQDN }
    Write-Host ('Connecting to Exchange On-Premises {0} using {1} ..' -f $global:myOffice365Services['ExchangeOnPremisesFQDN'], $global:myOffice365Services['OnPremisesCredentials'].username)
    $global:myOffice365Services['SessionExchange'] = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$($global:myOffice365Services['ExchangeOnPremisesFQDN'])/PowerShell" -Credential $global:myOffice365Services['OnPremisesCredentials'] -Authentication Kerberos -AllowRedirection -SessionOption $global:myOffice365Services['SessionExchangeOptions']
    If ( $global:myOffice365Services['SessionExchange']) {Import-PSSession -Session $global:myOffice365Services['SessionExchange'] -AllowClobber}
}

Function global:Get-ExchangeOnPremisesFQDN {
    $global:myOffice365Services['ExchangeOnPremisesFQDN'] = Read-Host -Prompt 'Enter Exchange On-Premises endpoint, e.g. exchange1.contoso.com'
}

function global:Connect-IPPSession {
    If ( !($global:myOffice365Services['Office365Credentials'])) { Get-Office365Credentials }
    If ( $global:myOffice365Services['Office365CredentialsMFA']) {
        Write-Host ('Connecting to Security & Compliance Center using {0} with Modern Authentication ..' -f $global:myOffice365Services['Office365Credentials'].username)
        $global:myOffice365Services['SessionCC'] = ExchangeOnlineManagement\Connect-IPPSSession -ConnectionUri $global:myOffice365Services['SCCConnectionEndpointUri'] -UserPrincipalName ($global:myOffice365Services['Office365Credentials']).UserName -PSSessionOption $global:myOffice365Services['SessionExchangeOptions']
    }
    Else {
        Write-Host ('Connecting to Security & Compliance Center using {0} ..' -f $global:myOffice365Services['Office365Credentials'].username)
        $global:myOffice365Services['SessionCC'] = ExchangeOnlineManagement\Connect-IPPSSession -ConnectionUrl $global:myOffice365Services['SCCConnectionEndpointUri'] -Credential $global:myOffice365Services['Office365Credentials'] -PSSessionOption $global:myOffice365Services['SessionExchangeOptions']
    }
    If ( $global:myOffice365Services['SessionCC'] ) {
        Import-PSSession -Session $global:myOffice365Services['SessionCC'] -AllowClobber
    }
}


function global:Connect-MSTeams {
    If ( !($global:myOffice365Services['Office365Credentials'])) { Get-Office365Credentials }
    If ( $global:myOffice365Services['Office365CredentialsMFA']) {
        Write-Host ('Connecting to Microsoft Teams using {0} with Modern Authentication ..' -f $global:myOffice365Services['Office365Credentials'].username)
        Connect-MicrosoftTeams -AccountId ($global:myOffice365Services['Office365Credentials']).UserName -TenantId $myOffice365Services['TenantId']
    }
    Else {
        Write-Host ('Connecting to Exchange Online Protection using {0} ..' -f $global:myOffice365Services['Office365Credentials'].username)
        Connect-MicrosoftTeams -Credential $global:myOffice365Services['Office365Credentials']
    }
}

function global:Connect-SkypeOnline {
    If ( !($global:myOffice365Services['Office365Credentials'])) { Get-Office365Credentials }
    Write-Host ('Connecting to Skype Online using {0}' -f $global:myOffice365Services['Office365Credentials'].username)
    $global:myOffice365Services['SessionSFBO']= New-CsOnlineSession -Credential $global:myOffice365Services['Office365Credentials']
    If ( $global:myOffice365Services['SessionSFBO'] ) {
        Import-PSSession -Session $global:myOffice365Services['SessionSFBO'] -AllowClobber
    }    
}

function global:Connect-AzureActiveDirectory {
    If ( !(Get-Module -Name AzureAD)) {Import-Module -Name AzureAD -ErrorAction SilentlyContinue}
    If ( !(Get-Module -Name AzureADPreview)) {Import-Module -Name AzureADPreview -ErrorAction SilentlyContinue}
    If ( (Get-Module -Name AzureAD) -or (Get-Module -Name AzureADPreview)) {
        If ( !($global:myOffice365Services['Office365Credentials'])) { Get-Office365Credentials }
        If ( $global:myOffice365Services['Office365CredentialsMFA']) {
            Write-Host 'Connecting to Azure Active Directory with Modern Authentication ..'
            $Parms = @{AccountId= $global:myOffice365Services['Office365Credentials'].UserName; AzureEnvironment= $global:myOffice365Services['AzureEnvironment']}
        }
        Else {
            Write-Host ('Connecting to Azure Active Directory using {0} ..' -f $global:myOffice365Services['Office365Credentials'].username)
            $Parms = @{'Credential' = $global:myOffice365Services['Office365Credentials']; 'AzureEnvironment' = $global:myOffice365Services['AzureEnvironment']}
        }
        Connect-AzureAD @Parms
    }
    Else {
        If ( !(Get-Module -Name MSOnline)) {Import-Module -Name MSOnline -ErrorAction SilentlyContinue}
        If ( Get-Module -Name MSOnline) {
            If ( !($global:myOffice365Services['Office365Credentials'])) { Get-Office365Credentials }
            Write-Host ('Connecting to Azure Active Directory using {0} ..' -f $global:myOffice365Services['Office365Credentials'].username)
            Connect-MsolService -Credential $global:myOffice365Services['Office365Credentials'] -AzureEnvironment $global:myOffice365Services['AzureEnvironment']
        }
        Else {Write-Error -Message 'Cannot connect to Azure Active Directory - problem loading module.'}
    }
}

function global:Connect-AIP {
    If ( !(Get-Module -Name AIPService)) {Import-Module -Name AIPService -ErrorAction SilentlyContinue}
    If ( Get-Module -Name AIPService) {
        If ( !($global:myOffice365Services['Office365Credentials'])) { Get-Office365Credentials }
        Write-Host ('Connecting to Azure Information Protection using {0}' -f $global:myOffice365Services['Office365Credentials'].username)
        Connect-AipService -Credential $global:myOffice365Services['Office365Credentials'] 
    }
    Else {Write-Error -Message 'Cannot connect to Azure Information Protection - problem loading module.'}
}

function global:Connect-SharePointOnline {
    If ( !(Get-Module -Name Microsoft.Online.Sharepoint.PowerShell)) {Import-Module -Name Microsoft.Online.Sharepoint.PowerShell -ErrorAction SilentlyContinue}
    If ( Get-Module -Name Microsoft.Online.Sharepoint.PowerShell) {
        If ( !($global:myOffice365Services['Office365Credentials'])) { Get-Office365Credentials }
        If (($global:myOffice365Services['Office365Credentials']).username -like '*.onmicrosoft.com') {
            $global:myOffice365Services['Office365Tenant'] = ($global:myOffice365Services['Office365Credentials']).username.Substring(($global:myOffice365Services['Office365Credentials']).username.IndexOf('@') + 1).Replace('.onmicrosoft.com', '')
        }
        Else {
            If ( !($global:myOffice365Services['Office365Tenant'])) { Get-Office365Tenant }
        }
        If ( $global:myOffice365Services['Office365CredentialsMFA']) {
            Write-Host 'Connecting to SharePoint Online with Modern Authentication ..'
            $Parms = @{
                url= 'https://{0}-admin.sharepoint.com' -f $($global:myOffice365Services['Office365Tenant'])
                region= $global:myOffice365Services['SharePointRegion']
            }
        }
        Else {
            Write-Host "Connecting to SharePoint Online using $($global:myOffice365Services['Office365Credentials'].username) .."
            $Parms = @{
                url= 'https://{0}-admin.sharepoint.com' -f $global:myOffice365Services['Office365Tenant']
                credential= $global:myOffice365Services['Office365Credentials']
                region= $global:myOffice365Services['SharePointRegion']
            }
        }
        Connect-SPOService @Parms
    }
    Else {
        Write-Error -Message 'Cannot connect to SharePoint Online - problem loading module.'
    }
}
function global:Connect-PowerApps {
    If ( !(Get-Module -Name Microsoft.PowerApps.PowerShell)) {Import-Module -Name Microsoft.PowerApps.PowerShell -ErrorAction SilentlyContinue}
    If ( !(Get-Module -Name Microsoft.PowerApps.Administration.PowerShell)) {Import-Module -Name Microsoft.PowerApps.Administration.PowerShell -ErrorAction SilentlyContinue}
    If ( Get-Module -Name Microsoft.PowerApps.PowerShell) {
        If ( !($global:myOffice365Services['Office365Credentials'])) { Get-Office365Credentials }
        Write-Host "Connecting to PowerApps using $($global:myOffice365Services['Office365Credentials'].username) .."
        If ( $global:myOffice365Services['Office365CredentialsMFA']) {
            $Parms = @{'Username' = $global:myOffice365Services['Office365Credentials'].UserName }
        }
        Else {
            $Parms = @{'Username' = $global:myOffice365Services['Office365Credentials'].UserName; 'Password'= $global:myOffice365Services['Office365Credentials'].Password }
        }
        Add-PowerAppsAccount @Parms
    }
    Else {
        Write-Error -Message 'Cannot connect to SharePoint Online - problem loading module.'
    }
}

Function global:Get-Office365Credentials {

    $global:myOffice365Services['Office365Credentials'] = $host.ui.PromptForCredential('Office 365 Credentials', 'Please enter your Office 365 credentials', $global:myOffice365Services['Office365Credentials'].UserName, '')
    $global:myOffice365Services['Office365CredentialsMFA'] = Get-MultiFactorAuthenticationUsage
    Get-TenantID
    Set-WindowTitle
}

Function global:Get-OnPremisesCredentials {
    $global:myOffice365Services['OnPremisesCredentials'] = $host.ui.PromptForCredential('On-Premises Credentials', 'Please Enter Your On-Premises Credentials', '', '')
}

Function global:Get-Office365Tenant {
    $global:myOffice365Services['Office365Tenant'] = Read-Host -Prompt 'Enter tenant ID, e.g. contoso for contoso.onmicrosoft.com'
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
    $ModuleManifestPath = $Module.Path
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
    $ModuleVersion
}

Function global:Update-Office365Modules {
    param (
        [switch]$AllowPrerelease
    )

    $local:Functions= Get-Office365ModuleInfo
    $local:ReposChecked= [System.Collections.ArrayList]::new()

    $local:IsAdmin= [System.Security.principal.windowsprincipal]::new([System.Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
    If( $local:IsAdmin) {
        If( (Get-Process -Name powershell, pwsh -ErrorAction SilentlyContinue | Measure-Object).Count -gt 1) {
            Write-Warning ('Running multiple PowerShell sessions, successful updating might be problematic.') 
        }
        ForEach ( $local:Function in $local:Functions) {
            $local:Item = ($local:Function).split('|')
            If( $local:Item[3] -and -not $local:ReposChecked.Contains( $local:Item[3])) {

                $local:Module= Get-Module -Name ('{0}' -f $local:Item[3]) -ListAvailable | Sort-Object -Property Version -Descending 

                $local:CheckThisModule= $false

                If( ([System.Uri]($local:Module | Select-Object -First 1).RepositorySourceLocation).Authority -eq (([System.Uri]$local:Item[5])).Authority) {
                    $local:CheckThisModule= $true
                }

                If( $local:CheckThisModule) {

                    If( $local:Item[5]) {
                       $local:Module= $local:Module | Where-Object {([System.Uri]($_.RepositorySourceLocation)).Authority -ieq ([System.Uri]($local:Item[5])).Authority } | Select-Object -First 1
                    }
                    Else {
                        $local:Module= $local:Module | Select-Object -First 1
                    }

                    If( ($local:Module).RepositorySourceLocation) {

                        $local:Version = Get-ModuleVersionInfo -Module $local:Module
                        Write-Host ('Checking {0}' -f $local:Item[4]) -NoNewLine

                        $local:NewerAvailable= $false
                        If( $local:Item[5]) {
                            $local:Repo= $local:Repos | Where-Object {([System.Uri]($_.SourceLocation)).Authority -eq (([System.Uri]$local:Item[5])).Authority}            
                        }
                        If( [string]::IsNullOrEmpty( $local:Repo )) { 
                            $local:Repo = 'PSGallery'
                        }
                        Else {
                            $local:Repo= ($local:Repo).Name
                        }
                        $OnlineModule = Find-Module -Name $local:Item[3] -Repository $local:Repo -AllowPrerelease:$AllowPrerelease -ErrorAction SilentlyContinue
                        If( $OnlineModule) {
                            Write-Host (': Local:{0}, Online:{1}' -f $local:Version, $OnlineModule.version)
                            If( (Compare-TextVersionNumber -Version $local:Version -CompareTo $OnlineModule.version) -eq 1) {
                                $local:NewerAvailable= $true
                            }
                            Else {
                                 # Local module up to date or newer
                            }
                        }
                        Else {
                             # Not installed from online or cannot determine
                             Write-Host ('Local:{0} Online:N/A' -f $local:Version)
                        }

                        If( $local:NewerAvailable) {

                            $local:UpdateSuccess= $false
                            Try {
                                $Parm= @{
                                    AllowPrerelease= $AllowPrerelease
                                    Force= $True
                                    Confirm= $False
                                    Scope= Get-ModuleScope -Module $local:Module
                                    AllowClobber= $True
                                }
                                # Pass AcceptLicense if current version of UpdateModule supports it
                                If( ( Get-Command -name Update-Module).Parameters['AcceptLicense']) {
                                    $Parm.AcceptLicense= $True
                                }
                                If( Get-Command Install-Package -ErrorAction SilentlyContinue) {
                                    If( ( Get-Command -name Install-Package).Parameters['SkipPublisherCheck']) {
                                        $Parm.SkipPublisherCheck= $True
                                    }
                                    Install-Package -Name $local:Item[3] -Source $local:Repo @Parm | Out-Null
                                }
                                Else{
                                    Update-Module -Name $local:Item[3] @Parm
                                }
                                $local:UpdateSuccess= $true
                            }
                            Catch {
                                Write-Error ('Problem updating module {0}:{1}' -f $local:Item[3], $Error[0].Message)
                            }

                            If( $local:UpdateSuccess) {

                                If( Get-Command -Name Get-InstalledModule -ErrorAction SilentlyContinue) {
                                    $local:ModuleVersions= Get-InstalledModule -Name $local:Item[3] -AllVersions 
                                }
                                Else {
                                    $local:ModuleVersions= Get-Module -Name $local:Item[3] -ListAvailable -All
                                }

                                $local:Module = $local:ModuleVersions | Sort-Object -Property @{e={ [System.Version]($_.Version -replace '[^\d\.]','')}} -Descending | Select-Object -First 1
                                $local:LatestVersion = ($local:Module).Version
                                Write-Host ('Updated {0} to version {1}' -f $local:Item[4], $local:LatestVersion) -ForegroundColor Green

                                # Uninstall all old versions of dependencies
                                If( $OnlineModule) {
                                    ForEach( $DependencyModule in $OnlineModule.Dependencies) {

                                        # Unload
                                        Remove-Module -Name $DependencyModule.Name -Force -Confirm:$False -ErrorAction SilentlyContinue

                                        $local:DepModuleVersions= Get-Module -Name $DependencyModule.Name -ListAvailable
                                        $local:DepModule = $local:DepModuleVersions | Sort-Object -Property @{e={ [System.Version]($_.Version -replace '[^\d\.]','')}} -Descending | Select-Object -First 1
                                        $local:DepLatestVersion = ($local:DepModule).Version
                                        $local:OldDepModules= $local:DepModuleVersions | Where-Object {$_.Version -ne $local:DepLatestVersion}
                                        ForEach( $DepModule in $local:OldDepModules) {
                                            Write-Host ('Uninstalling dependency module {0} version {1}' -f $DepModule.Name, $DepModule.Version)
                                            Try {
                                                $DepModule | Uninstall-Module -Confirm:$false -Force
                                            }
                                            Catch {
                                                Write-Error ('Problem uninstalling module {0} version {1}' -f $DepModule.Name, $DepModule.Version) 
                                            }
                                        }
                                    }
                                }

                                # Uninstall all old versions of the module
                                $local:OldModules= $local:ModuleVersions | Where-Object {$_.Version -ne $local:LatestVersion}
                                If( $local:OldModules) {

                                    # Unload module when currently loaded
                                    Remove-Module -Name $local:Item[3] -Force -Confirm:$False -ErrorAction SilentlyContinue

                                    ForEach( $OldModule in $local:OldModules) {
                                        Write-Host ('Uninstalling {0} version {1}' -f $local:Item[4], $OldModule.Version) -ForegroundColor White
                                        Try {
                                            $OldModule | Uninstall-Module -Confirm:$false -Force
                                        }
                                        Catch {
                                            Write-Error ('Problem uninstalling module {0} version {1}' -f $OldModule.Name, $OldModule.Version) 
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
                        Write-Host ('Skipping {0}: Not installed using PowerShellGet/Install-Module' -f $local:Item[4]) -ForegroundColor Yellow
                    }
                }
            }
            $null= $local:ReposChecked.Add( $local:Item[3])
        }
    }
    Else {
        Write-Host ('Script not running with elevated privileges; cannot update modules') -ForegroundColor Yellow
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
    $local:Repos= Get-PSRepository
    $local:ReposChecked= [System.Collections.ArrayList]::new()

    ForEach ( $local:Function in $local:Functions) {

        $local:Item = ($local:Function).split('|')
        If( $local:Item[3] -and -not $local:ReposChecked.Contains( $local:Item[3])) {
            $local:Module= Get-Module -Name ('{0}' -f $local:Item[3]) -ListAvailable | Sort-Object -Property Version -Descending

            # Use specific or default repository
            If( $local:Item[5]) {
                $local:Repo= $local:Repos | Where-Object {([System.Uri]($_.SourceLocation)).Authority -eq (([System.Uri]$local:Item[5])).Authority}
            }
            If( [string]::IsNullOrEmpty( $local:Repo )) { 
                $local:Repo = 'PSGallery'
            }
            Else {
                $local:Repo= ($local:Repo).Name
            }

            If( $local:Item[5]) {
                $local:Module= $local:Module | Where-Object {([System.Uri]($_.RepositorySourceLocation)).Authority -ieq ([System.Uri]($local:Item[5])).Authority } | Select-Object -First 1
            }
            Else {
                $local:Module= $local:Module | Select-Object -First 1
            }

            If( $local:Module) {

                $local:Version = Get-ModuleVersionInfo -Module $local:Module

                Write-Host ('Module {0}: Local v{1}' -f $local:Item[4], $Local:Version) -NoNewline
   
                $OnlineModule = Find-Module -Name $local:Item[3] -Repository $local:Repo -AllowPrerelease:$AllowPrerelease -ErrorAction SilentlyContinue
                If( $OnlineModule) {
                    Write-Host (', Online v{0}' -f $OnlineModule.version) -NoNewline
                }
                Else {
                    Write-Host (', Online N/A') -NoNewline
                }
                Write-Host (', Scope:{0} Status:' -f (Get-ModuleScope -Module $local:Module)) -NoNewline

                If( [string]::IsNullOrEmpty( $local:Version) -or [string]::IsNullOrEmpty( $OnlineModule.version)) {
                    Write-Host ('Unknown')
                }
                Else {
                    If( (Compare-TextVersionNumber -Version $local:Version -CompareTo $OnlineModule.version) -eq 1) {
                        Write-Host ('Outdated') -ForegroundColor Red
                    }
                    Else {
                        Write-Host ('OK') -ForegroundColor Green
                    }
                }
            }
            Else {
                Write-Host ('{0} module not found ({1})' -f $local:Item[4], $local:Item[5])
            }
        }
        $null= $local:ReposChecked.Add( $local:Item[3]) 
    }
}

function global:Connect-Office365 {
    Connect-AzureActiveDirectory
    Connect-AzureRMS
    Connect-ExchangeOnline
    Connect-MSTeams
    Connect-SkypeOnline
    Connect-ComplianceCenter
    Connect-SharePointOnline
}

$PSGetModule= Get-Module -Name PowerShellGet -ListAvailable -ErrorAction SilentlyContinue | Sort-Object -Property Version -Descending | Select-Object -First 1
If(! $PSGetModule) {
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

# See if the Administator built-in role is part of your role
$local:IsAdmin= [System.Security.principal.windowsprincipal]::new([System.Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)

$local:CreateISEMenu = $psISE -and -not [System.Windows.Input.Keyboard]::IsKeyDown( [System.Windows.Input.Key]::LeftShift)
If ( $local:CreateISEMenu) {Write-Host 'ISE detected, adding ISE menu options'}

# Initialize global state variable when needed
If( -not( Get-Variable myOffice365Services -ErrorAction SilentlyContinue )) { $global:myOffice365Services=@{} }

# Local Exchange session options
$global:myOffice365Services['SessionExchangeOptions'] = New-PSSessionOption -ProxyAccessType None

# Initialize environment & endpoints
Set-Office365Environment -AzureEnvironment 'Default'

Write-Host ('Environment:{0}, Administrator:{1}' -f $global:myOffice365Services['AzureEnvironment'], $local:IsAdmin)
Write-Host ('Architecture:{0}, PS:{1}, PSGet:{2}, PackageManagement:{3}' -f ($ENV:PROCESSOR_ARCHITECTURE), ($PSVersionTable).PSVersion, $PSGetVer, $PMMVer )
Write-Host ('*' * 78)

$local:Functions= Get-Office365ModuleInfo
$local:Repos= Get-PSRepository

Write-Host ('Collecting Module information ..')

$local:ReposChecked= [System.Collections.ArrayList]::new() 

ForEach ( $local:Function in $local:Functions) {

    $local:Item = ($local:Function).split('|')
    $local:CreateMenuItem= $False
    If( $local:Item[3] -and -not $local:ReposChecked.Contains( $local:Item[3])) {
        $local:Module= Get-Module -Name ('{0}' -f $local:Item[3]) -ListAvailable | Sort-Object -Property Version -Descending
        $local:ModuleMatch= ([System.Uri]($local:Module | Select-Object -First 1).RepositorySourceLocation).Authority -eq ([System.Uri]$local:Item[5]).Authority
        If( $local:ModuleMatch) {
            $local:Module = $local:Module | Sort-Object -Property @{e= { [System.Version]($_.Version -replace '[^\d\.]','')}} -Descending
            If( $local:Item[5]) {
                $local:Module= $local:Module | Where-Object {([System.Uri]($_.RepositorySourceLocation)).Authority -ieq ([System.Uri]($local:Item[5])).Authority } | Select-Object -First 1
            }
            Else {
                $local:Module= $local:Module | Select-Object -First 1
            }
            $local:Version = Get-ModuleVersionInfo -Module $local:Module
            Write-Host ('Found {0} module (v{1})' -f $local:Item[4], $local:Version) -ForegroundColor Green
            $local:CreateMenuItem= $True
        }
        Else {
            # Module not found
        }
        $null= $local:ReposChecked.Add( $local:Item[3])
    }
    Else {
        # Local function
        $local:CreateMenuItem= $True
    }

    If( $local:CreateMenuItem -and $local:CreateISEMenu) {
        # Create menu item when module found or local function 
        $local:MenuObj = $psISE.CurrentPowerShellTab.AddOnsMenu.SubMenus | Where-Object -FilterScript { $_.DisplayName -eq $local:Item[0] }
        If ( !( $local:MenuObj)) {
            Try {$local:MenuObj = $psISE.CurrentPowerShellTab.AddOnsMenu.SubMenus.Add( $local:Item[0], $null, $null)}
            Catch {Write-Warning -Message $_}
        }
        Try {
            $local:RemoveItems = $local:MenuObj.Submenus |  Where-Object -FilterScript { $_.DisplayName -eq $local:Item[1] -or $_.Action -eq $local:Item[2] }
            $null = $local:RemoveItems | ForEach-Object -Process { $local:MenuObj.Submenus.Remove( $_) }
            $null = $local:MenuObj.SubMenus.Add( $local:Item[1], [ScriptBlock]::Create( $local:Item[2]), $null)
        }
        Catch {
            Write-Warning -Message $_
        }
    }
}
