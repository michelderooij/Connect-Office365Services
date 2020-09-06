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

    Version 2.45, September 6th, 2020

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
    - Connect-AzureActiveDirectory	Connects to Azure Active Directory
    - Connect-AzureRMS           	Connects to Azure Rights Management
    - Connect-ExchangeOnline     	Connects to Exchange Online
    - Connect-ExchangeOnlinev2          Connects to Exchange Online using REST module 
    - Connect-SkypeOnline        	Connects to Skype for Business Online
    - Connect-EOP                	Connects to Exchange Online Protection
    - Connect-ComplianceCenter   	Connects to Compliance Center
    - Connect-SharePointOnline   	Connects to SharePoint Online
    - Connect-MSTeams                   Connects to Microsoft Teams
    - Get-Office365Credentials    	Gets Office 365 credentials
    - Connect-ExchangeOnPremises 	Connects to Exchange On-Premises
    - Get-OnPremisesCredentials    	Gets On-Premises credentials
    - Get-ExchangeOnPremisesFQDN        Gets FQDN for Exchange On-Premises
    - Get-Office365Tenant		Gets Office 365 tenant name
    - Set-Office365Environment		Configures Uri's and region to use
    - Update-Office365Modules           Updates supported Office 365 modules
    - Report-Office365Modules           Report on known vs online module versions

    Functions to connect to other services provided by the module, e.g. Connect-MSGraph or Connect-MSTeams.

    To register the PowerShell Test Gallery and install modules from there, use:
    Register-PSRepository -Name PSGalleryInt -SourceLocation https://www.poshtestgallery.com/ -InstallationPolicy Trusted
    Install-Module -Name MicrosoftTeams -Repository PSGalleryInt -Force

    To load the helper functions from your PowerShell profile, put Connect-Office365Services.ps1 in the same location
    as your $profile file, and edit $profile as follows:
    & (Join-Path $PSScriptRoot "Connect-Office365Services.ps1")

    .HISTORY
    1.2	    Community release
    1.3     Updated required version of Online Sign-In Assistant
    1.4	    Added (in-code) AzureEnvironment (Connect-AzureAD)
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
#>

#Requires -Version 3.0
$local:ScriptVersion= '2.45'

function global:Set-WindowTitle {
    If( $host.ui.RawUI.WindowTitle -and $global:myOffice365Services['TenantID']) {
        $local:PromptPrefix= ''
        $ThisPrincipal= new-object System.Security.principal.windowsprincipal( [System.Security.Principal.WindowsIdentity]::GetCurrent())
        if( $ThisPrincipal.IsInRole("Administrators")) { 
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
        $doc= Invoke-WebRequest -Uri ('https://login.microsoft.com/{0}/FederationMetadata/2007-06/FederationMetadata.xml' -f $domainPart)
        $GUIDmatch= $doc -match 'sts\.windows\.net\/(?<TenantID>[0-9a-f-]*)'
        If( $doc -and $GUIDmatch) {
             $res= $matches.TenantID
        }
        Else {
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
        'Connect|Exchange Online|Connect-ExchangeOnline||Exchange Online (v1)|',
        'Connect|Exchange Online (v2)|Connect-ExchangeOnlinev2|ExchangeOnlineManagement|Exchange Online Management (v2)|https://www.powershellgallery.com/packages/ExchangeOnlineManagement',
        'Connect|Exchange Online Protection|Connect-EOP',
        'Connect|Exchange Compliance Center|Connect-ComplianceCenter',
        'Connect|Azure AD (v1)|Connect-MSOnline|MSOnline|Azure Active Directory (v1)|https://www.powershellgallery.com/packages/MSOnline',
        'Connect|Azure AD (v2)|Connect-AzureAD|AzureAD|Azure Active Directory (v2)|https://www.powershellgallery.com/packages/azuread',
        'Connect|Azure AD (v2 Preview)|Connect-AzureAD|AzureADPreview|Azure Active Directory (v2 Preview)|https://www.powershellgallery.com/packages/AzureADPreview',
        'Connect|Azure Information Protection|Connect-AIP|AIPService|Azure Information Protection|https://www.powershellgallery.com/packages/AIPService',
        'Connect|Skype for Business Online|Connect-SkypeOnline|SkypeOnlineConnector|Skype for Business Online|https://www.microsoft.com/en-us/download/details.aspx?id=39366',
        'Connect|SharePoint Online|Connect-SharePointOnline|Microsoft.Online.Sharepoint.PowerShell|SharePoint Online|https://www.powershellgallery.com/packages/Microsoft.Online.SharePoint.PowerShell',
        'Connect|Microsoft Teams|Connect-MSTeams|MicrosoftTeams|Microsoft Teams (GA)|https://www.powershellgallery.com/packages/MicrosoftTeams',
        'Connect|Microsoft Teams|Connect-MSTeams|MicrosoftTeams|Microsoft Teams (Test)|https://www.poshtestgallery.com/packages/MicrosoftTeams',
        'Connect|Microsoft Commerce|Connect-MSCommerce|MSCommerce|Microsoft Commerce|https://www.powershellgallery.com/packages/MSCommerce',
        'Connect|Microsoft.Graph.Teams|Connect-Graph|Microsoft.Graph.Teams.Team|Microsoft.Graph.Teams.Team|https://www.powershellgallery.com/packages/Microsoft.Graph.Teams.Team',
        'Connect|SharePoint PnP Online|Connect-PnPOnline|SharePointPnPPowerShellOnline|SharePointPnP Online|https://www.powershellgallery.com/packages/SharePointPnPPowerShellOnline',
        'Connect|PowerApps-Admin-PowerShell|Connect-PowerApps|Microsoft.PowerApps.Administration.PowerShell|PowerApps-Admin-PowerShell|https://www.powershellgallery.com/packages/Microsoft.PowerApps.Administration.PowerShell',
        'Connect|PowerApps-PowerShell|Connect-PowerApps|Microsoft.PowerApps.PowerShell|PowerApps-PowerShell|https://www.powershellgallery.com/packages/Microsoft.PowerApps.PowerShell',
        'Connect|MSGraph-Intune|Connect-MSGraph|Microsoft.Graph.Intune|MSGraph-Intune|https://www.powershellgallery.com/packages/Microsoft.Graph.Intune',
        'Connect|Microsoft.Graph|Connect-Graph|Microsoft.Graph|Microsoft.Graph|https://www.powershellgallery.com/packages/Microsoft.Graph',
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
            $global:myOffice365Services['SCCConnectionEndpointUri'] = 'https://ps.compliance.protection.outlook.com/PowerShell-LiveId'
            $global:myOffice365Services['AzureADAuthorizationEndpointUri'] = 'https://login.microsoftonline.de/common'
            $global:myOffice365Services['SharePointRegion'] = 'Germany'
            $global:myOffice365Services['AzureEnvironment'] = 'AzureGermanyCloud'
        }
        'China' {
            $global:myOffice365Services['ConnectionEndpointUri'] = 'https://partner.outlook.cn/PowerShell-LiveID'
            $global:myOffice365Services['SCCConnectionEndpointUri'] = 'https://ps.compliance.protection.outlook.com/PowerShell-LiveId'
            $global:myOffice365Services['AzureADAuthorizationEndpointUri'] = 'https://login.chinacloudapi.cn/common'
            $global:myOffice365Services['SharePointRegion'] = 'China'
            $global:myOffice365Services['AzureEnvironment'] = 'AzureChinaCloud'
        }
        'AzurePPE' {
            $global:myOffice365Services['ConnectionEndpointUri'] = ''
            $global:myOffice365Services['SCCConnectionEndpointUri'] = 'https://ps.compliance.protection.outlook.com/PowerShell-LiveId'
            $global:myOffice365Services['AzureADAuthorizationEndpointUri'] = ''
            $global:myOffice365Services['SharePointRegion'] = ''
            $global:myOffice365Services['AzureEnvironment'] = 'AzurePPE'
        }
        'USGovernment' {
            $global:myOffice365Services['ConnectionEndpointUri'] = 'https://outlook.office365.com/PowerShell-LiveId'
            $global:myOffice365Services['SCCConnectionEndpointUri'] = 'https://ps.compliance.protection.outlook.com/PowerShell-LiveId'
            $global:myOffice365Services['AzureADAuthorizationEndpointUri'] = 'https://login-us.microsoftonline.com/'
            $global:myOffice365Services['SharePointRegion'] = 'ITAR'
            $global:myOffice365Services['AzureEnvironment'] = 'AzureUSGovernment'
        }
        default {
            $global:myOffice365Services['ConnectionEndpointUri'] = 'https://outlook.office365.com/PowerShell-LiveId'
            $global:myOffice365Services['SCCConnectionEndpointUri'] = 'https://ps.compliance.protection.outlook.com/PowerShell-LiveId'
            $global:myOffice365Services['AzureADAuthorizationEndpointUri'] = 'https://login.windows.net/common'
            $global:myOffice365Services['SharePointRegion'] = 'Default'
            $global:myOffice365Services['AzureEnvironment'] = 'AzureCloud'
        }
    }
}

function global:Get-MultiFactorAuthenticationUsage {
    $Answer = Read-host  -Prompt 'Would you like to use Modern Authentication? (y/N) '
    Switch ($Answer.ToUpper()) {
        'Y' { $rval = $true }
        Default { $rval = $false}
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
    If ( !($global:myOffice365Services['Office365Credentials'])) { Get-Office365Credentials }
    If ( $global:myOffice365Services['Office365CredentialsMFA']) {
        Write-Host "Connecting to Exchange Online using $($global:myOffice365Services['Office365Credentials'].username) with Modern Authentication .."
        $global:myOffice365Services['Session365'] = New-ExoPSSession -ConnectionUri $global:myOffice365Services['ConnectionEndpointUri'] -UserPrincipalName ($global:myOffice365Services['Office365Credentials']).UserName -AzureADAuthorizationEndpointUri $global:myOffice365Services['AzureADAuthorizationEndpointUri'] -PSSessionOption $global:myOffice365Services['SessionExchangeOptions']
    }
    Else {
        Write-Host "Connecting to Exchange Online using $($global:myOffice365Services['Office365Credentials'].username) .."
        $global:myOffice365Services['Session365'] = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $global:myOffice365Services['ConnectionEndpointUri'] -Credential $global:myOffice365Services['Office365Credentials'] -Authentication Basic -AllowRedirection -SessionOption $global:myOffice365Services['SessionExchangeOptions']
    }
    If ( $global:myOffice365Services['Session365'] ) {
        Import-PSSession -Session $global:myOffice365Services['Session365'] -AllowClobber
    }
}

function global:Connect-ExchangeOnlinev2 {
    If ( !($global:myOffice365Services['Office365Credentials'])) { Get-Office365Credentials }
    # Other connect options: TrackPerformance, UseMultithreading, Showprogress, EnableEXOTelemetry, LogDirectoryPath
    If ( $global:myOffice365Services['Office365CredentialsMFA']) {
        Write-Host "Connecting to Exchange Online (v2) using $($global:myOffice365Services['Office365Credentials'].username) with Modern Authentication .."
        $global:myOffice365Services['Session365'] = ExchangeOnlineManagement\Connect-ExchangeOnline -ConnectionUri $global:myOffice365Services['ConnectionEndpointUri'] -UserPrincipalName ($global:myOffice365Services['Office365Credentials']).UserName -AzureADAuthorizationEndpointUri $global:myOffice365Services['AzureADAuthorizationEndpointUri'] -PSSessionOption $global:myOffice365Services['SessionExchangeOptions']
    }
    Else {
        Write-Host "Connecting to Exchange Online (v2) using $($global:myOffice365Services['Office365Credentials'].username) .."
        $global:myOffice365Services['Session365'] = ExchangeOnlineManagement\Connect-ExchangeOnline -ConnectionUrl $global:myOffice365Services['ConnectionEndpointUri'] -Credential $global:myOffice365Services['Office365Credentials'] -SessionOption $global:myOffice365Services['SessionExchangeOptions']
    }
    If ( $global:myOffice365Services['Session365'] ) {
        Import-PSSession -Session $global:myOffice365Services['Session365'] -AllowClobber
    }
}

function global:Connect-ExchangeOnPremises {
    If ( !($global:myOffice365Services['OnPremisesCredentials'])) { Get-OnPremisesCredentials }
    If ( !($global:myOffice365Services['ExchangeOnPremisesFQDN'])) { Get-ExchangeOnPremisesFQDN }
    Write-Host "Connecting to Exchange On-Premises $($global:myOffice365Services['ExchangeOnPremisesFQDN']) using $($global:myOffice365Services['OnPremisesCredentials'].username) .."
    $global:myOffice365Services['SessionExchange'] = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$($global:myOffice365Services['ExchangeOnPremisesFQDN'])/PowerShell" -Credential $global:myOffice365Services['OnPremisesCredentials'] -Authentication Kerberos -AllowRedirection -SessionOption $global:myOffice365Services['SessionExchangeOptions']
    If ( $global:myOffice365Services['SessionExchange']) {Import-PSSession -Session $global:myOffice365Services['SessionExchange'] -AllowClobber}
}

Function global:Get-ExchangeOnPremisesFQDN {
    $global:myOffice365Services['ExchangeOnPremisesFQDN'] = Read-Host -Prompt 'Enter Exchange On-Premises endpoint, e.g. exchange1.contoso.com'
}

function global:Connect-ComplianceCenter {
    If ( !($global:myOffice365Services['Office365Credentials'])) { Get-Office365Credentials }
    If ( $global:myOffice365Services['Office365CredentialsMFA']) {
        Write-Host "Connecting to Security & Compliance Center using $($global:myOffice365Services['Office365Credentials'].username) with Modern Authentication .."
        $global:myOffice365Services['SessionCC'] = New-ExoPSSession -ConnectionUri $global:myOffice365Services['SCCConnectionEndpointUri'] -UserPrincipalName ($global:myOffice365Services['Office365Credentials']).UserName -AzureADAuthorizationEndpointUri $global:myOffice365Services['AzureADAuthorizationEndpointUri'] -PSSessionOption $global:myOffice365Services['SessionExchangeOptions']
    }
    Else {
        Write-Host "Connecting to Security & Compliance Center using $($global:myOffice365Services['Office365Credentials'].username) .."
        $global:myOffice365Services['SessionCC'] = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $global:myOffice365Services['SCCConnectionEndpointUri'] -Credential $global:myOffice365Services['Office365Credentials'] -Authentication Basic -AllowRedirection
        If ( $global:myOffice365Services['SessionCC'] ) {Import-PSSession -Session $global:myOffice365Services['SessionCC'] -AllowClobber}
    }
    If ( $global:myOffice365Services['SessionCC'] ) {
        Import-PSSession -Session $global:myOffice365Services['SessionCC'] -AllowClobber
    }    
}

function global:Connect-EOP {
    If ( !($global:myOffice365Services['Office365Credentials'])) { Get-Office365Credentials }
    Write-Host  -InputObject "Connecting to Exchange Online Protection using $($global:myOffice365Services['Office365Credentials'].username) .."
    $global:myOffice365Services['SessionEOP'] = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri 'https://ps.protection.outlook.com/powershell-liveid/' -Credential $global:myOffice365Services['Office365Credentials'] -Authentication Basic -AllowRedirection
    If ( $global:myOffice365Services['SessionEOP'] ) {Import-PSSession -Session $global:myOffice365Services['SessionEOP'] -AllowClobber}
}

function global:Connect-MSTeams {
    If ( !($global:myOffice365Services['Office365Credentials'])) { Get-Office365Credentials }
    If (($global:myOffice365Services['Office365Credentials']).username -like '*.onmicrosoft.com') {
        $global:myOffice365Services['Office365Tenant'] = ($global:myOffice365Services['Office365Credentials']).username.Substring(($global:myOffice365Services['Office365Credentials']).username.IndexOf('@') + 1).Replace('.onmicrosoft.com', '')
    }
    If ( $global:myOffice365Services['Office365CredentialsMFA']) {
        Write-Host "Connecting to Microsoft Teams using $($global:myOffice365Services['Office365Credentials'].username) with Modern Authentication .."
        $Parms = @{'AccountId' = ($global:myOffice365Services['Office365Credentials']).username}
    }
    Else {
        Write-Host "Connecting to Microsoft Teams using $($global:myOffice365Services['Office365Credentials'].username) .."
        $Parms = @{Credential = $global:myOffice365Services['Office365Credentials'] }
    }
    If ( $global:myOffice365Services['Office365Tenant']) { $Parms['TenantId'] = $global:myOffice365Services['Office365Tenant'] }
    Connect-MicrosoftTeams @Parms
}

function global:Connect-AzureActiveDirectory {
    If ( !(Get-Module -Name AzureAD)) {Import-Module -Name AzureAD -ErrorAction SilentlyContinue}
    If ( !(Get-Module -Name AzureADPreview)) {Import-Module -Name AzureADPreview -ErrorAction SilentlyContinue}
    If ( (Get-Module -Name AzureAD) -or (Get-Module -Name AzureADPreview)) {
        If ( !($global:myOffice365Services['Office365Credentials'])) { Get-Office365Credentials }
        If ( $global:myOffice365Services['Office365CredentialsMFA']) {
            Write-Host 'Connecting to Azure Active Directory with Modern Authentication ..'
            $Parms = @{'AzureEnvironment' = $global:myOffice365Services['AzureEnvironment']}
        }
        Else {
            Write-Host "Connecting to Azure Active Directory using $($global:myOffice365Services['Office365Credentials'].username) .."
            $Parms = @{'Credential' = $global:myOffice365Services['Office365Credentials']; 'AzureEnvironment' = $global:myOffice365Services['AzureEnvironment']}
        }
        Connect-AzureAD @Parms
    }
    Else {
        If ( !(Get-Module -Name MSOnline)) {Import-Module -Name MSOnline -ErrorAction SilentlyContinue}
        If ( Get-Module -Name MSOnline) {
            If ( !($global:myOffice365Services['Office365Credentials'])) { Get-Office365Credentials }
            Write-Host "Connecting to Azure Active Directory using $($global:myOffice365Services['Office365Credentials'].username) .."
            Connect-MsolService -Credential $global:myOffice365Services['Office365Credentials'] -AzureEnvironment $global:myOffice365Services['AzureEnvironment']
        }
        Else {Write-Error -Message 'Cannot connect to Azure Active Directory - problem loading module.'}
    }
}

function global:Connect-AIP {
    If ( !(Get-Module -Name AIPService)) {Import-Module -Name AIPService -ErrorAction SilentlyContinue}
    If ( Get-Module -Name AIPService) {
        If ( !($global:myOffice365Services['Office365Credentials'])) { Get-Office365Credentials }
        Write-Host "Connecting to Azure Information Protection using $($global:myOffice365Services['Office365Credentials'].username) .."
        Connect-AipService -Credential $global:myOffice365Services['Office365Credentials'] 
    }
    Else {Write-Error -Message 'Cannot connect to Azure Information Protection - problem loading module.'}
}

function global:Connect-SkypeOnline {
    If ( !(Get-Module -Name SkypeOnlineConnector)) {Import-Module -Name SkypeOnlineConnector -ErrorAction SilentlyContinue}
    If ( Get-Module -Name SkypeOnlineConnector) {
        If ( !($global:myOffice365Services['Office365Credentials'])) { Get-Office365Credentials }
        If ( $global:myOffice365Services['Office365CredentialsMFA']) {
            Write-Host "Connecting to Skype for Business Online using $($global:myOffice365Services['Office365Credentials'].username) with Modern Authentication .."
            $Parms = @{'Username' = ($global:myOffice365Services['Office365Credentials']).username}
        }
        Else {
            Write-Host "Connecting to Skype for Business Online using $($global:myOffice365Services['Office365Credentials'].username) .."
            $Parms = @{'Credential' = $global:myOffice365Services['Office365Credentials']}
        }
        $global:myOffice365Services['SessionSFB'] = New-CsOnlineSession @Parms
        If ( $global:myOffice365Services['SessionSFB'] ) {
            Import-PSSession -Session $global:myOffice365Services['SessionSFB'] -AllowClobber
        }
    }
    Else {
        Write-Error -Message 'Cannot connect to Skype for Business Online - problem loading module.'
    }
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
            $Parms = @{'url' = "https://$($global:myOffice365Services['Office365Tenant'])-admin.sharepoint.com"; 'Region' = $global:myOffice365Services['SharePointRegion']}
        }
        Else {
            Write-Host "Connecting to SharePoint Online using $($global:myOffice365Services['Office365Credentials'].username) .."
            $Parms = @{'url' = "https://$($global:myOffice365Services['Office365Tenant'])-admin.sharepoint.com"; 'Credential' = $global:myOffice365Services['Office365Credentials']; 'Region' = $global:myOffice365Services['SharePointRegion']}
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
    $global:myOffice365Services['Office365Credentials'] = $host.ui.PromptForCredential('Office 365 Credentials', 'Please enter your Office 365 credentials', '', '')
    $local:MFAenabledModulePresence= $false
    # Check for MFA-enabled modules 
    If ( (Get-Module -Name 'Microsoft.Exchange.Management.ExoPowershellModule') -or (Get-Module -Name 'MicrosoftTeams')) {
        $local:MFAenabledModulePresence= $true
    }
    Else {
        # Check for MFA-enabled modules with version dependency
        $MFAMods= @('SkypeOnlineConnector|7.0', 'Microsoft.Online.Sharepoint.PowerShell|16.0')
	ForEach( $MFAMod in $MFAMods) {
            $local:Item = ($local:MFAMod).split('|')
            If( (Get-Module -Name $local:Item[0] -ListAvailable)) {
                $local:MFAenabledModulePresence= $local:MFAenabledModulePresence -or ((Get-Module -Name $local:Item[0] -ListAvailable).Version -ge [System.Version]($local:Item[1] -replace '[^\d\.]','') )
            }
        }
    }
    If( $local:MFAEnabledModulePresence) {
        $global:myOffice365Services['Office365CredentialsMFA'] = Get-MultiFactorAuthenticationUsage
    }
    Else {
        $global:myOffice365Services['Office365CredentialsMFA'] = $false
    }
    Get-TenantID
    Set-WindowTitle
}

Function global:Get-AllowPrereleaseModule {
    If( $global:myOffice365Services['AllowPrerelease']) {
        # Already asked
    }
    Else {
        $Answer = Read-host  -Prompt 'Would you like to check for pre-release modules? (y/N) '
        Switch ($Answer.ToUpper()) {
            'Y' { $rval = $true }
            Default { $rval = $false}
        }
        $global:myOffice365Services['AllowPrerelease']= $rval
    }
}

Function global:Get-OnPremisesCredentials {
    $global:myOffice365Services['OnPremisesCredentials'] = $host.ui.PromptForCredential('On-Premises Credentials', 'Please Enter Your On-Premises Credentials', '', '')
}

Function global:Get-Office365Tenant {
    $global:myOffice365Services['Office365Tenant'] = Read-Host -Prompt 'Enter tenant ID, e.g. contoso for contoso.onmicrosoft.com'
}

Function global:Update-Office365Modules {
    Get-AllowPrereleaseModule
    $local:Functions= Get-Office365ModuleInfo
    $local:AvailableModules= Get-Module -ListAvailable
    $local:IsAdmin= [System.Security.principal.windowsprincipal]::new([System.Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
    If( $local:IsAdmin) {
        ForEach ( $local:Function in $local:Functions) {
            $local:Item = ($local:Function).split('|')
            If( $local:Item[3]) {
                $local:CheckThisModule= $false
                If( ([System.Uri](Get-Module -Name ('{0}' -f $local:Item[3]) -ListAvailable | Select -First 1).RepositorySourceLocation).Authority -eq (([System.Uri]$local:Item[5])).Authority) {
                    $local:CheckThisModule= $true
                }

                If( $local:CheckThisModule) {
                    $local:Module = $local:AvailableModules | Where {$_.Name -ieq $local:Item[3] } | Sort-Object -Property Version -Descending | Select-Object -First 1
                    If( ($local:Module).RepositorySourceLocation) {
                        $local:Version = ($local:Module).Version
                        Write-Host ('Checking {0}' -f $local:Item[4]) 

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
                        $OnlineModule = Find-Module -Name $local:Item[3] -Repository $local:Repo -AllowPrerelease:$global:myOffice365Services['AllowPrerelease'] -ErrorAction SilentlyContinue
                        If( $OnlineModule) {
                            If( [System.Version]($local:Version -replace '[^\d\.]','') -lt [System.Version]($OnlineModule.version -replace '[^\d\.]','')) {
                                $local:NewerAvailable= $true
                            }
                            Else {
                                 # Local module up to date or newer
                            }
                        }
                        Else {
                             # Not installed from online or cannot determine
                        }

                        If( $local:NewerAvailable) {
                            $local:UpdateSuccess= $false
                            Try {
                                # Pass AcceptLicense if current version of UpdateModule supports it
                                If( ( Get-Command -name Update-Module).Parameters['AcceptLicense']) {
                                    Update-Module -Name $local:Item[3] -AllowPrerelease:$global:myOffice365Services['AllowPrerelease'] -Force -Confirm:$false -AcceptLicense
                                }
                                Else {
                                    Update-Module -Name $local:Item[3] -AllowPrerelease:$global:myOffice365Services['AllowPrerelease'] -Force -Confirm:$false
                                }
                                $local:UpdateSuccess= $true
                            }
                            Catch {
                                Write-Error ('Problem updating module {0}:{1}' -f $local:Item[3], $Error[0].Message)
                            }

                            If( $local:UpdateSuccess) {
                                # Uninstall all old versions of the module
                                $local:OldModules= Get-Module -Name $local:Item[3] -ListAvailable | Sort-Object -Property Version -Descending | Select-Object -Skip 1
                                If( $local:OldModules) {
                                    ForEach( $Module in $local:OldModules) {
                                        Write-Host ('Uninstalling {0} version {1}' -f $local:Item[4], $Module.Version) -ForegroundColor White
                                        $Module | Uninstall-Module -Confirm:$false -Force 
                                    }
                                }

                                $local:Module = Get-Module -Name $local:Item[3] -ListAvailable | Sort-Object -Property Version -Descending | Select-Object -First 1
                                $local:NewVersion = ($local:Module).Version[0]
                                If( [System.Version]($local:NewVersion -replace '[^\d\.]','') -gt [System.Version]($local:Version -replace '[^\d\.]','')) {
                                    Write-Host ('Installed {0} version {1}' -f $local:Item[4], [System.Version]($local:NewVersion -replace '[^\d\.]','')) -ForegroundColor Green
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
                        Write-Host ('Skipping {0}: Not installed using Install-Module' -f $local:Item[4]) -ForegroundColor Yellow
                    }
                }
            }
        }
    }
    Else {
        Write-Host ('Script not running with elevated privileges; cannot update modules') -ForegroundColor Yellow
    }
}

Function global:Report-Office365Modules {

    Get-AllowPrereleaseModule

    $local:Functions= Get-Office365ModuleInfo
    $local:Repos= Get-PSRepository
    $local:AvailableModules= Get-Module -ListAvailable

    ForEach ( $local:Function in $local:Functions) {

        $local:Item = ($local:Function).split('|')
        If( $local:Item[3]) {

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

            $local:Module= $local:AvailableModules | Where {$_.Name -ieq $local:Item[3] } | Where {[System.Uri]($_.RepositorySourceLocation).Authority -eq ([System.Uri]($local:Item[5])).Authority } | Select-Object -First 1
            If( $local:Module) {
                $local:Version = ($local:Module).Version

                Write-Host ('Module: {0} - Checked: v{1}, Online: ' -f $local:Item[4], $local:Version) -NoNewLine
                $OnlineModule = Find-Module -Name $local:Item[3] -Repository $local:Repo -AllowPrerelease:$global:myOffice365Services['AllowPrerelease'] -ErrorAction SilentlyContinue
                If( $OnlineModule) {
                    Write-Host ('v{0}' -f [System.Version]($OnlineModule.version -replace '[^\d\.]','')) -NoNewLine
                }
                Else {
                    Write-Host ('N/A') -NoNewLine
                }
                If( [string]::IsNullOrEmpty( $local:Version) -or [string]::IsNullOrEmpty( $OnlineModule.version)) {
                    Write-Host (' Unknown') -ForegroundColor Yellow
                }
                Else {
                    If( [System.Version]($local:Version -replace '[^\d\.]','') -ige [System.Version]($OnlineModule.version -replace '[^\d\.]','')) {
                        Write-Host (' OK') -ForegroundColor Green
                    }
                    Else {
                        Write-Host (' Outdated') -ForegroundColor RED
                    }
                }
            }
        }
        Else {

            Switch( $local:Item[1]) {

                'Exchange Online' {

                    #Scan for Exchange & SCC MFA PowerShell module presence
                    $local:ExchangeMFAModule = 'Microsoft.Exchange.Management.ExoPowershellModule'
                    $local:ExchangeADALModule = 'Microsoft.IdentityModel.Clients.ActiveDirectory'
                    $local:ModuleList = @(Get-ChildItem -Path "$($env:LOCALAPPDATA)\Apps\2.0" -Filter "$($local:ExchangeMFAModule).manifest" -Recurse -ErrorAction SilentlyContinue ) | Sort-Object LastWriteTime -Desc | Select-Object -First 1

                    If ( $local:ModuleList) {
                        $local:ModuleName = Join-path -Path $local:ModuleList[0].Directory.FullName -ChildPath "$($local:ExchangeMFAModule).dll"
                        $local:ModuleVersion = (Get-Item -Path $local:ModuleName).VersionInfo.ProductVersion
                        Write-Host ('Module: ExoPowershellModule - Checked v{0}, Online: ' -f $local:ModuleVersion) -NoNewLine

                        $OnlineModuleVersion= Get-ExchangeOnlineClickOnceVersion
                        Write-Host ('v{0}' -f [System.Version]($OnlineModuleVersion)) -NoNewLine

                        If( [System.Version]($local:ModuleVersion -replace '[^\d\.]','') -ieq [System.Version]($OnlineModuleVersion -replace '[^\d\.]','')) {
                            Write-Host (' OK') -ForegroundColor Green
                        }
                        Else {
                            Write-Host (' Outdated') -ForegroundColor RED
                        }
                    }
                }
            }
        }
    }
}

function global:Connect-Office365 {
    Connect-AzureActiveDirectory
    Connect-AzureRMS
    Connect-ExchangeOnline
    Connect-MSTeams
    Connect-SkypeOnline
    Connect-EOP
    Connect-ComplianceCenter
    Connect-SharePointOnline
}

Write-Host ('*' * 78)
Write-Host ('Connect-Office365Services v{0}' -f $local:ScriptVersion)

If( $ENV:PROCESSOR_ARCHITECTURE -eq 'AMD64') {
    $local:Platform= 'x64'
}
Else {
    $local:Platform= 'x86 (impacts module availalability)'
}

$local:HasInternetAccess = ([Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]'{DCB00C01-570F-4A9B-8D69-199FDBA5723B}')).IsConnectedToInternet)

# See if the Administator built-in role is part of your role
$local:IsAdmin= [System.Security.principal.windowsprincipal]::new([System.Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)

$local:CreateISEMenu = $psISE -and -not [System.Windows.Input.Keyboard]::IsKeyDown( [System.Windows.Input.Key]::LeftShift)
If ( $local:CreateISEMenu) {Write-Host 'ISE detected, adding ISE menu options'}

# Initialize global state variable when needed
If( -not( Get-Variable myOffice365Services -ErrorAction SilentlyContinue )) { $global:myOffice365Services=@{} }

# Local Exchange session options
$global:myOffice365Services['SessionExchangeOptions'] = New-PSSessionOption

# Initialize environment & endpoints
Set-Office365Environment -AzureEnvironment 'Default'

Write-Host ('Environment:{0}, IsAdmin:{1}, InternetAccess:{2}, Platform:{3}' -f $global:myOffice365Services['AzureEnvironment'], $local:IsAdmin, $local:HasInternetAccess, $local:Platform)
Write-Host ('*' * 78)

$local:Functions= Get-Office365ModuleInfo
$local:Repos= Get-PSRepository

Write-Host ('Collecting Module information ..')
$local:AvailableModules= Get-Module -ListAvailable

ForEach ( $local:Function in $local:Functions) {

    $local:Item = ($local:Function).split('|')
    If( [string]::IsNullOrEmpty($local:Item[3])) {
        # Non-repo module

        Switch( $local:Item[1]) {
            'Exchange Online' {

                #Scan for Exchange & SCC MFA PowerShell module presence
                $local:ExchangeMFAModule = 'Microsoft.Exchange.Management.ExoPowershellModule'
                $local:ExchangeADALModule = 'Microsoft.IdentityModel.Clients.ActiveDirectory'
                $local:ModuleList = @(Get-ChildItem -Path ('{0}\Apps\2.0' -f $env:LOCALAPPDATA) -Filter ('{0}.manifest' -f $local:ExchangeMFAModule) -Recurse -ErrorAction SilentlyContinue ) | Sort-Object LastWriteTime -Desc | Select-Object -First 1

                If ( $local:ModuleList) {
                    $local:ModuleName = Join-path -Path $local:ModuleList[0].Directory.FullName -ChildPath ('{0}.dll' -f $local:ExchangeMFAModule)
                    $local:ModuleVersion = (Get-Item -Path $local:ModuleName).VersionInfo.ProductVersion
                    Write-Host ('Module {0} installed (v{1})' -f $local:Item[4], $local:ModuleVersion) -ForegroundColor Green
                    Import-Module -FullyQualifiedName $local:ModuleName -Force
                    $local:ModuleName = Join-path -Path $local:ModuleList[0].Directory.FullName -ChildPath ('{0}.dll' -f $local:ExchangeADALModule)
                    If ( Test-Path -Path $local:ModuleName) {
                        $local:ModuleVersion = (Get-Item -Path $local:ModuleName).VersionInfo.FileVersion
                        Write-Host "ADAL module found (version $($local:ModuleVersion))" -ForegroundColor Green
                        Add-Type -Path $local:ModuleName
                    }
                    $local:ModuleMatch= $true
                }
                Else {
                    Write-Verbose 'Exchange Modern Authentication PowerShell Module not found (http://bit.ly/ExOPSModule)'
                }

            }
            default {
                $local:ModuleMatch= $true
            }
        }
    }
    Else {
        # Match module from specific repo
        $local:ModuleMatch= ($local:AvailableModules | Where {$_.Name -ieq $local:Item[3] }).RepositorySourceLocation.Authority -eq ([System.Uri]$local:Item[5]).Authority
    }

    If( $local:ModuleMatch) {
        If ( $local:CreateISEMenu) {
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
        If ( $local:Item[3]) {
            $local:Module = $local:AvailableModules | Where {$_.Name -ieq $local:Item[3] } | Sort-Object -Property Version -Descending
            If( $local:Item[5]) {
                $local:Module= $local:Module | Where {([System.Uri]($_.RepositorySourceLocation)).Authority -ieq ([System.Uri]($local:Item[5])).Authority } | Select-Object -First 1
            }
            Else {
                $local:Module= $local:Module | Select-Object -First 1
            }
            $local:Version = ($local:Module).Version[0]
            Write-Host ('Found {0} module (v{1})' -f $local:Item[4], $local:Version) -ForegroundColor Green
        }
    }
    Else {
        Write-Host "$($local:Item[4]) module not found ($($local:Item[5]))" 
    }
}
